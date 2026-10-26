#requires -Version 7.2
<#
.SYNOPSIS
Creates or updates a ConnectWise PSA opportunity, ensures the correct SharePoint
the flat opportunity folder exists, updates opportunity-folder metadata, and writes the opportunity folder URL back to
ConnectWise.

.ROUTING
NewClient      -> https://palaisparc.sharepoint.com/sales
                  Company Opportunities
ExistingClient -> https://palaisparc.sharepoint.com/sales/AccountManagement
                  Client Opportunities

.FOLDER STRUCTURE
<Library>/<Opportunity Number - Opportunity Name>

.SHAREPOINT FIELDS UPDATED
OpportunityName
OpportunityNumber
OpportunityType
ConnectWiseLink
OpportunityBriefDescription
CompanyName
CompanyNumber

.REQUIRED FUNCTION APP SETTINGS
ConnectWisePsa_ApiBaseUrl
ConnectWisePsa_ApiCompanyId
ConnectWisePsa_ApiPublicKey
ConnectWisePsa_ApiPrivateKey
ConnectWisePsa_ApiClientId
Ms365_TenantId
Ms365_AuthAppId
Ms365_CertBase64
Ms365_CertPassword

.OPTIONAL FUNCTION APP SETTINGS
ConnectWisePsa_ApiVersion
OpportunityWorkspace_CustomFieldCaption
OpportunityWorkspace_CustomFieldId
OpportunityWorkspace_NewClientTypeRegex
OpportunityWorkspace_ExistingClientTypeRegex
OpportunityWorkspace_DefaultCategory
OpportunityWorkspace_DryRun
#>
param($Request, $TriggerMetadata)

Set-StrictMode -Version 1
$ErrorActionPreference = 'Stop'

function Get-Setting {
    param([string]$Name,[string]$Default=$null)
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
    return $value.Trim()
}
function Test-SettingTrue {
    param([string]$Name)
    return (Get-Setting $Name '') -match '^(?i:true|1|yes)$'
}
function Assert-RequiredValue {
    param([string]$Name,[AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw "Required value '$Name' was not supplied." }
}
function Get-PropertyValue {
    param([object]$Object,[string[]]$Paths)
    foreach ($path in $Paths) {
        $current = $Object; $found = $true
        foreach ($segment in $path.Split('.')) {
            if ($null -eq $current -or $current.PSObject.Properties.Name -notcontains $segment) { $found=$false; break }
            $current = $current.$segment
        }
        if ($found -and $null -ne $current -and -not [string]::IsNullOrWhiteSpace([string]$current)) { return [string]$current }
    }
    return $null
}
function ConvertTo-SafeFolderName {
    param([Parameter(Mandatory)][string]$Name)
    $safe = $Name -replace '["*:<>?/\\|#%]', '-'
    $safe = $safe -replace '[\x00-\x1F]', ''
    $safe = $safe -replace '\s+', ' '
    $safe = $safe.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safe)) { throw 'Opportunity name produced an empty SharePoint folder name.' }
    if ($safe.Length -gt 180) { $safe = $safe.Substring(0,180).Trim().TrimEnd('.') }
    return $safe
}
function Get-RequestBody {
    param($Request)
    if ($null -eq $Request) { throw 'Request is required.' }
    $body = $Request.Body
    if ($body -is [string]) { return $body | ConvertFrom-Json -Depth 100 }
    if ($body -is [byte[]]) { return ([Text.Encoding]::UTF8.GetString($body) | ConvertFrom-Json -Depth 100) }
    return $body
}

# Fixed SharePoint routing requested by CloudFirst.
[string]$SalesSiteUrl = 'https://palaisparc.sharepoint.com/sales'
[string]$SalesLibraryName = 'Company Opportunities'
[string]$AccountSiteUrl = 'https://palaisparc.sharepoint.com/sales/AccountManagement'
[string]$AccountLibraryName = 'Client Opportunities'

[string]$ConnectWiseApiHost = Get-Setting 'ConnectWisePsa_ApiBaseUrl'
[string]$ConnectWiseCompanyId = Get-Setting 'ConnectWisePsa_ApiCompanyId'
[string]$ConnectWisePublicKey = Get-Setting 'ConnectWisePsa_ApiPublicKey'
[string]$ConnectWisePrivateKey = Get-Setting 'ConnectWisePsa_ApiPrivateKey'
[string]$ConnectWiseClientId = Get-Setting 'ConnectWisePsa_ApiClientId'
[string]$ConnectWiseApiVersion = Get-Setting 'ConnectWisePsa_ApiVersion' '2022.1'
[string]$TenantId = Get-Setting 'Ms365_TenantId'
[string]$SharePointClientId = Get-Setting 'Ms365_AuthAppId'
[string]$CertificateBase64 = Get-Setting 'Ms365_CertBase64'
[string]$CertificatePassword = Get-Setting 'Ms365_CertPassword'
[string]$WorkspaceCustomFieldCaption = Get-Setting 'OpportunityWorkspace_CustomFieldCaption' 'SharePoint Links'
[string]$WorkspaceCustomFieldId = Get-Setting 'OpportunityWorkspace_CustomFieldId'
[string]$NewClientTypeRegex = Get-Setting 'OpportunityWorkspace_NewClientTypeRegex' '(?i)new client|new customer|prospect'
[string]$ExistingClientTypeRegex = Get-Setting 'OpportunityWorkspace_ExistingClientTypeRegex' '(?i)existing client|existing customer|account management|expansion|renewal'
[string]$DefaultCategory = Get-Setting 'OpportunityWorkspace_DefaultCategory' ''
$DryRun = Test-SettingTrue 'OpportunityWorkspace_DryRun'

foreach ($pair in @{
    ConnectWisePsa_ApiBaseUrl=$ConnectWiseApiHost; ConnectWisePsa_ApiCompanyId=$ConnectWiseCompanyId;
    ConnectWisePsa_ApiPublicKey=$ConnectWisePublicKey; ConnectWisePsa_ApiPrivateKey=$ConnectWisePrivateKey;
    ConnectWisePsa_ApiClientId=$ConnectWiseClientId; Ms365_TenantId=$TenantId;
    Ms365_AuthAppId=$SharePointClientId; Ms365_CertBase64=$CertificateBase64;
    Ms365_CertPassword=$CertificatePassword
}.GetEnumerator()) { Assert-RequiredValue $pair.Key ([string]$pair.Value) }

$apiBase = $ConnectWiseApiHost.Trim()
if ($apiBase -notmatch '^https://') { $apiBase = "https://$apiBase" }
$apiBaseUri = $null
if (-not [Uri]::TryCreate($apiBase,[UriKind]::Absolute,[ref]$apiBaseUri) -or $apiBaseUri.Scheme -ne 'https') {
    throw 'ConnectWisePsa_ApiBaseUrl must be a valid HTTPS host or base URL.'
}
$script:CWBase = '{0}/v4_6_release/apis/3.0' -f $apiBaseUri.AbsoluteUri.TrimEnd('/')
$authPlain = "$ConnectWiseCompanyId+$ConnectWisePublicKey`:$ConnectWisePrivateKey"
$authEncoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($authPlain))
$script:CWHeaders = @{
    Authorization = "Basic $authEncoded"
    ClientId = $ConnectWiseClientId
    Accept = "application/vnd.connectwise.com+json; version=$ConnectWiseApiVersion"
    'Content-Type' = 'application/json'
}

function Invoke-CWRequest {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PUT','PATCH')][string]$Method,
        [Parameter(Mandatory)][string]$RelativePath,
        [object]$Body,
        [ValidateRange(1,10)][int]$MaxAttempts=5
    )
    $uri = "$script:CWBase/$($RelativePath.TrimStart('/'))"
    for ($attempt=1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $p = @{ Method=$Method; Uri=$uri; Headers=$script:CWHeaders; TimeoutSec=180 }
            if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) { $p.Body = $Body | ConvertTo-Json -Depth 100 -Compress }
            return Invoke-RestMethod @p
        } catch {
            $status=$null; $responseText=$null
            try { $status=[int]$_.Exception.Response.StatusCode } catch {}
            try { if ($_.ErrorDetails.Message) { $responseText=$_.ErrorDetails.Message } } catch {}
            $transient = ($null -eq $status) -or $status -in @(408,429) -or ($status -ge 500 -and $status -le 599)
            if (-not $transient -or $attempt -eq $MaxAttempts) {
                throw "ConnectWise $Method failed. Path: $RelativePath; Status: $status; Error: $($_.Exception.Message); Response: $responseText"
            }
            $delay=[Math]::Min(60,[Math]::Pow(2,$attempt)+(Get-Random -Minimum 0 -Maximum 1000)/1000)
            Start-Sleep -Milliseconds ([int]($delay*1000))
        }
    }
}
function Get-CWOpportunity { param([int]$Id) Invoke-CWRequest GET "sales/opportunities/$Id" }
function New-CWOpportunity { param([object]$Opportunity) Invoke-CWRequest POST 'sales/opportunities' $Opportunity }
function Update-CWOpportunity {
    param([int]$Id,[object]$Opportunity)
    $items=@($Opportunity)
    $isPatch=$items.Count -gt 0 -and $items[0].PSObject.Properties.Name -contains 'op' -and $items[0].PSObject.Properties.Name -contains 'path'
    if ($isPatch) { Invoke-CWRequest PATCH "sales/opportunities/$Id" $items }
    else { Invoke-CWRequest PUT "sales/opportunities/$Id" $Opportunity }
}
function Get-WorkspaceFieldValue {
    param([object]$Opportunity)
    $fields=@($Opportunity.customFields); $target=$null
    if ($WorkspaceCustomFieldId) { $target=$fields | Where-Object { [string]$_.id -eq $WorkspaceCustomFieldId } | Select-Object -First 1 }
    if (-not $target) { $target=$fields | Where-Object { [string]$_.caption -eq $WorkspaceCustomFieldCaption } | Select-Object -First 1 }
    if ($target -and -not [string]::IsNullOrWhiteSpace([string]$target.value)) { return [string]$target.value }
    return $null
}
function Set-CWWorkspaceUrl {
    param([object]$Opportunity,[string]$Url)
    $fields=@($Opportunity.customFields); $target=$null
    if ($WorkspaceCustomFieldId) { $target=$fields | Where-Object { [string]$_.id -eq $WorkspaceCustomFieldId } | Select-Object -First 1 }
    if (-not $target) { $target=$fields | Where-Object { [string]$_.caption -eq $WorkspaceCustomFieldCaption } | Select-Object -First 1 }
    if (-not $target) { throw "Opportunity custom field '$WorkspaceCustomFieldCaption' was not returned by ConnectWise." }
    $target.value=$Url
    if (-not $DryRun) { Invoke-CWRequest PATCH "sales/opportunities/$($Opportunity.id)" @(@{op='replace';path='customFields';value=$fields}) | Out-Null }
}
function Resolve-WorkspaceCategory {
    param([object]$RequestBody,[object]$Opportunity)
    $explicit=Get-PropertyValue $RequestBody @('workspaceCategory')
    if ($explicit -match '^(?i)newclient|new client|newcustomer|new customer|sales$') { return 'NewClient' }
    if ($explicit -match '^(?i)existingclient|existing client|existingcustomer|existing customer|accountmanagement|account management$') { return 'ExistingClient' }
    if ($explicit) { throw "workspaceCategory '$explicit' is invalid. Use NewClient or ExistingClient." }
    $typeText=Get-PropertyValue $Opportunity @('type.name','type.identifier')
    if ($typeText -match $NewClientTypeRegex) { return 'NewClient' }
    if ($typeText -match $ExistingClientTypeRegex) { return 'ExistingClient' }
    if ($DefaultCategory -in @('NewClient','ExistingClient')) { return $DefaultCategory }
    throw "Could not classify opportunity type '$typeText'. Supply workspaceCategory or configure the type regex settings."
}
function Get-WorkspaceRoute {
    param([string]$Category)
    if ($Category -eq 'NewClient') { return @{SiteUrl=$SalesSiteUrl;LibraryName=$SalesLibraryName} }
    return @{SiteUrl=$AccountSiteUrl;LibraryName=$AccountLibraryName}
}
function Connect-WorkspaceSite {
    param([string]$SiteUrl)
    $securePassword=ConvertTo-SecureString $CertificatePassword -AsPlainText -Force
    Connect-PnPOnline -Url $SiteUrl -ClientId $SharePointClientId -Tenant $TenantId -CertificateBase64Encoded $CertificateBase64 -CertificatePassword $securePassword -ReturnConnection
}
function Get-ConnectWiseLink {
    param([object]$RequestBody,[object]$Opportunity)
    $provided=Get-PropertyValue $RequestBody @('connectWiseLink')
    if ($provided) { return $provided }
    $returned=Get-PropertyValue $Opportunity @('_info.link','_info.href')
    if ($returned) { return $returned }
    return "$script:CWBase/sales/opportunities/$($Opportunity.id)"
}
function Ensure-WorkspaceLibraryFields {
    param([object]$Connection,[string]$LibraryName)
    $definitions = @(
        @{InternalName='OpportunityName';DisplayName='Opportunity Name';Type='Text'},
        @{InternalName='OpportunityNumber';DisplayName='Opportunity Number';Type='Text'},
        @{InternalName='OpportunityType';DisplayName='Opportunity Type';Type='Text'},
        @{InternalName='CompanyName';DisplayName='Company Name';Type='Text'},
        @{InternalName='CompanyNumber';DisplayName='ConnectWise Company Number';Type='Text'},
        @{InternalName='ConnectWiseLink';DisplayName='ConnectWise Link';Type='URL'},
        @{InternalName='OpportunityBriefDescription';DisplayName='Brief Description of the Opportunity';Type='Note'}
    )
    foreach ($definition in $definitions) {
        $field=Get-PnPField -List $LibraryName -Identity $definition.InternalName -Connection $Connection -ErrorAction SilentlyContinue
        if (-not $field) {
            Add-PnPField -List $LibraryName -DisplayName $definition.DisplayName -InternalName $definition.InternalName -Type $definition.Type -AddToDefaultView -Connection $Connection | Out-Null
        }
    }
}
function Set-WorkspaceMetadata {
    param([object]$Connection,[string]$LibraryName,[object]$Folder,[object]$Opportunity,[object]$RequestBody)
    $type=Get-PropertyValue $Opportunity @('type.name','type.identifier')
    $description=Get-PropertyValue $RequestBody @('briefDescription')
    if (-not $description) { $description=Get-PropertyValue $Opportunity @('notes','description','summary') }
    $values=@{
        OpportunityName = (Get-PropertyValue $Opportunity @('name','summary'))
        OpportunityNumber = (Get-PropertyValue $Opportunity @('id','opportunityNumber'))
        OpportunityType = $type
        CompanyName = (Get-PropertyValue $Opportunity @('company.name'))
        CompanyNumber = (Get-PropertyValue $Opportunity @('company.identifier','company.id'))
        ConnectWiseLink = (Get-ConnectWiseLink $RequestBody $Opportunity)
        OpportunityBriefDescription = $description
    }
    $Folder.Context.Load($Folder.ListItemAllFields)
    $Folder.Context.ExecuteQuery()
    Set-PnPListItem -List $LibraryName -Identity $Folder.ListItemAllFields.Id -Values $values -Connection $Connection | Out-Null
    return $values
}
function Ensure-Workspace {
    param([object]$Opportunity,[string]$Category,[object]$RequestBody)
    $route=Get-WorkspaceRoute $Category
    $number=Get-PropertyValue $Opportunity @('id','opportunityNumber')
    $name=Get-PropertyValue $Opportunity @('name','summary')
    Assert-RequiredValue 'Opportunity number/id' $number
    Assert-RequiredValue 'Opportunity name' $name

    # Flat hierarchy:
    # <Opportunity Library>/<Opportunity Number - Opportunity Name>
    $opportunityFolderName=ConvertTo-SafeFolderName "$number - $name"
    $targetPath="$($route.LibraryName)/$opportunityFolderName"
    $existingUrl=Get-WorkspaceFieldValue $Opportunity
    $connection=Connect-WorkspaceSite $route.SiteUrl
    try {
        if (-not $DryRun) { Ensure-WorkspaceLibraryFields $connection $route.LibraryName }
        if ($existingUrl) {
            $existingUri=$null
            if (-not [Uri]::TryCreate($existingUrl,[UriKind]::Absolute,[ref]$existingUri) -or $existingUri.Scheme -ne 'https') { throw "Invalid existing SharePoint URL: $existingUrl" }
            $folderPath=[Uri]::UnescapeDataString($existingUri.AbsolutePath)
            $folder=Get-PnPFolder -Url $folderPath -Connection $connection -Includes ServerRelativeUrl,ListItemAllFields -ErrorAction Stop
            if (-not $DryRun) { $metadata=Set-WorkspaceMetadata $connection $route.LibraryName $folder $Opportunity $RequestBody }
            return @{Url=$existingUrl;Created=$false;Path=$folder.ServerRelativeUrl;Metadata=$metadata;Note='Existing opportunity folder verified and metadata refreshed. Existing linked folders are not automatically moved.'}
        }
        if ($DryRun) { return @{Url=$null;Created=$false;Path=$targetPath;Metadata=$null;Note='Dry run: opportunity folder not created.'} }
        $folder=Resolve-PnPFolder -SiteRelativePath $targetPath -Connection $connection -Includes ServerRelativeUrl,ListItemAllFields
        $web=Get-PnPWeb -Connection $connection -Includes Url
        $url=([Uri]::new([Uri]$web.Url,$folder.ServerRelativeUrl)).AbsoluteUri
        $metadata=Set-WorkspaceMetadata $connection $route.LibraryName $folder $Opportunity $RequestBody
        return @{Url=$url;Created=$true;Path=$folder.ServerRelativeUrl;Metadata=$metadata;Note='Flat opportunity folder ensured and metadata updated.'}
    } finally { Disconnect-PnPOnline -Connection $connection -ErrorAction SilentlyContinue }
}

$statusCode=200; $result=$null
try {
    Import-Module PnP.PowerShell -ErrorAction Stop
    $body=Get-RequestBody $Request
    $mode=Get-PropertyValue $body @('mode','action')
    if ($mode -notmatch '^(?i:create|update)$') { throw "mode must be 'Create' or 'Update'." }
    if ($null -eq $body.opportunity) { throw 'The request body must include opportunity.' }

    if ($mode -match '^(?i:create)$') {
        if ($DryRun) { throw 'Dry-run Create is blocked because ConnectWise must assign the opportunity number.' }
        $opportunity=New-CWOpportunity $body.opportunity
        $operation='Created'; $statusCode=201
    } else {
        $idText=Get-PropertyValue $body @('opportunityId','id')
        if (-not $idText -or $idText -notmatch '^\d+$') { throw 'Update requires a numeric opportunityId.' }
        if (-not $DryRun) { Update-CWOpportunity ([int]$idText) $body.opportunity | Out-Null }
        $opportunity=Get-CWOpportunity ([int]$idText)
        $operation=if ($DryRun) {'UpdateDryRun'} else {'Updated'}
    }

    $category=Resolve-WorkspaceCategory $body $opportunity
    $workspace=Ensure-Workspace $opportunity $category $body
    $currentUrl=Get-WorkspaceFieldValue $opportunity
    if ($workspace.Url -and $currentUrl -ne $workspace.Url) { Set-CWWorkspaceUrl $opportunity $workspace.Url }
    $finalOpportunity=if ($DryRun) {$opportunity} else {Get-CWOpportunity ([int]$opportunity.id)}

    $result=[ordered]@{
        success=$true; operation=$operation; dryRun=$DryRun
        opportunityId=$finalOpportunity.id; opportunityName=$finalOpportunity.name
        workspaceCategory=$category; workspaceUrl=$workspace.Url
        workspaceCreated=$workspace.Created; workspacePath=$workspace.Path
        sharePointMetadata=$workspace.Metadata; note=$workspace.Note
    }
} catch {
    $statusCode=400
    $result=[ordered]@{success=$false;error=$_.Exception.Message;type=$_.Exception.GetType().FullName}
    Write-Error $_
}
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode=$statusCode
    Headers=@{'Content-Type'='application/json'}
    Body=($result | ConvertTo-Json -Depth 20)
})
