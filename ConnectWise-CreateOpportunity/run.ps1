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
    param([string]$Name,[AllowEmptyString()][string]$Value,[int]$StatusCode=400)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw (New-RequestException "Required value '$Name' was not supplied." $StatusCode) }
}
function New-RequestException {
    param([string]$Message,[int]$StatusCode=400)
    $exception=[InvalidOperationException]::new($Message)
    $exception.Data['HttpStatusCode']=$StatusCode
    return $exception
}
function Assert-HttpsUrl {
    param([string]$Name,[string]$Value)
    $uri=$null
    if (-not [Uri]::TryCreate($Value,[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.UserInfo) {
        throw (New-RequestException "$Name must be an absolute HTTPS URL without embedded credentials.")
    }
    return $uri
}
function Test-WorkspaceCategory {
    param([string]$Category)
    return $Category -in @('NewClient','ExistingClient')
}
function Assert-WorkspaceConfiguration {
    foreach ($setting in @(
        @{Name='OpportunityWorkspace_NewClientTypeRegex';Value=$NewClientTypeRegex},
        @{Name='OpportunityWorkspace_ExistingClientTypeRegex';Value=$ExistingClientTypeRegex}
    )) {
        try { [regex]::new($setting.Value) | Out-Null }
        catch { throw "Invalid regular expression in $($setting.Name): $($_.Exception.Message)" }
    }
    if ($DefaultCategory -and -not (Test-WorkspaceCategory $DefaultCategory)) {
        throw "OpportunityWorkspace_DefaultCategory must be NewClient or ExistingClient."
    }
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
    if ([string]::IsNullOrWhiteSpace($safe)) { throw (New-RequestException 'Opportunity name produced an empty SharePoint folder name.') }
    if ($safe.Length -gt 180) { $safe = $safe.Substring(0,180).Trim().TrimEnd('.') }
    return $safe
}
function Get-RequestBody {
    param($Request)
    if ($null -eq $Request) { throw 'Request is required.' }
    $body = $Request.Body
    if ($body -is [string]) {
        try { return $body | ConvertFrom-Json -Depth 100 }
        catch { throw (New-RequestException "Request body is not valid JSON: $($_.Exception.Message)") }
    }
    if ($body -is [byte[]]) {
        try { return ([Text.Encoding]::UTF8.GetString($body) | ConvertFrom-Json -Depth 100) }
        catch { throw (New-RequestException "Request body is not valid UTF-8 JSON: $($_.Exception.Message)") }
    }
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
[string]$WorkspaceCustomFieldCaption = Get-Setting 'OpportunityWorkspace_CustomFieldCaption' 'Sharepoint Link'
[string]$WorkspaceCustomFieldId = Get-Setting 'OpportunityWorkspace_CustomFieldId' '76'
[string]$NewClientTypeRegex = Get-Setting 'OpportunityWorkspace_NewClientTypeRegex' '(?i)new client|new customer|prospect'
[string]$ExistingClientTypeRegex = Get-Setting 'OpportunityWorkspace_ExistingClientTypeRegex' '(?i)existing client|existing customer|account management|expansion|renewal'
[string]$DefaultCategory = Get-Setting 'OpportunityWorkspace_DefaultCategory' ''
$DryRun = Test-SettingTrue 'OpportunityWorkspace_DryRun'

Assert-WorkspaceConfiguration
foreach ($pair in @{
    ConnectWisePsa_ApiBaseUrl=$ConnectWiseApiHost; ConnectWisePsa_ApiCompanyId=$ConnectWiseCompanyId;
    ConnectWisePsa_ApiPublicKey=$ConnectWisePublicKey; ConnectWisePsa_ApiPrivateKey=$ConnectWisePrivateKey;
    ConnectWisePsa_ApiClientId=$ConnectWiseClientId; Ms365_TenantId=$TenantId;
    Ms365_AuthAppId=$SharePointClientId; Ms365_CertBase64=$CertificateBase64;
    Ms365_CertPassword=$CertificatePassword
}.GetEnumerator()) { Assert-RequiredValue $pair.Key ([string]$pair.Value) 500 }

$apiBase = $ConnectWiseApiHost.Trim()
if ($apiBase -notmatch '^https://') { $apiBase = "https://$apiBase" }
$apiBaseUri = $null
if (-not [Uri]::TryCreate($apiBase,[UriKind]::Absolute,[ref]$apiBaseUri) -or $apiBaseUri.Scheme -ne 'https' -or
    $apiBaseUri.UserInfo -or $apiBaseUri.Query -or $apiBaseUri.Fragment) {
    throw 'ConnectWisePsa_ApiBaseUrl must be a valid HTTPS host or base URL.'
}
$apiPath=$apiBaseUri.AbsolutePath.TrimEnd('/')
if ($apiPath -notmatch '(?i)/v4_6_release/apis/3\.0$') {
    $apiPath="$apiPath/v4_6_release/apis/3.0"
}
$script:CWBase = '{0}://{1}{2}' -f $apiBaseUri.Scheme,$apiBaseUri.Authority,$apiPath
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
        [ValidateRange(1,3)][int]$MaxAttempts=3
    )
    $uri = "$script:CWBase/$($RelativePath.TrimStart('/'))"
    $retryableMethod=$Method -in @('GET','PUT')
    if (-not $retryableMethod) { $MaxAttempts=1 }
    for ($attempt=1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $p = @{ Method=$Method; Uri=$uri; Headers=$script:CWHeaders; TimeoutSec=30 }
            if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) { $p.Body = $Body | ConvertTo-Json -Depth 100 -Compress }
            return Invoke-RestMethod @p
        } catch {
            $status=$null; $responseText=$null
            try { $status=[int]$_.Exception.Response.StatusCode } catch {}
            try { if ($_.ErrorDetails.Message) { $responseText=$_.ErrorDetails.Message } } catch {}
            $transient = ($null -eq $status) -or $status -in @(408,429) -or ($status -ge 500 -and $status -le 599)
            if (-not $transient -or -not $retryableMethod -or $attempt -eq $MaxAttempts) {
                $exception=New-RequestException "ConnectWise $Method failed. Path: $RelativePath; Status: $status; Error: $($_.Exception.Message); Response: $responseText" 502
                throw $exception
            }
            $delay=[Math]::Min(8,[Math]::Pow(2,$attempt-1)+(Get-Random -Minimum 0 -Maximum 1000)/1000)
            Start-Sleep -Milliseconds ([int]($delay*1000))
        }
    }
}
function Get-CWOpportunity { param([int]$Id) Invoke-CWRequest GET "sales/opportunities/$Id" }
function New-CWOpportunity { param([object]$Opportunity) Invoke-CWRequest POST 'sales/opportunities' $Opportunity }
function Merge-OpportunityObject {
    param([object]$Current,[object]$Updates)
    foreach ($property in $Updates.PSObject.Properties) {
        $currentProperty=$Current.PSObject.Properties[$property.Name]
        if ($null -ne $currentProperty -and $null -ne $currentProperty.Value -and
            $null -ne $property.Value -and
            $currentProperty.Value -isnot [string] -and $property.Value -isnot [string] -and
            $currentProperty.Value -isnot [System.Collections.IEnumerable] -and $property.Value -isnot [System.Collections.IEnumerable] -and
            $currentProperty.Value.PSObject.Properties.Count -gt 0 -and $property.Value.PSObject.Properties.Count -gt 0) {
            $null=Merge-OpportunityObject $currentProperty.Value $property.Value
        } else {
            $Current | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value -Force
        }
    }
    return $Current
}
function Update-CWOpportunity {
    param([int]$Id,[object]$Opportunity,[object]$CurrentOpportunity)
    $items=@($Opportunity)
    $isPatch=$Opportunity -is [System.Collections.IEnumerable] -and $Opportunity -isnot [string] -and
        $items.Count -gt 0 -and $items[0].PSObject.Properties.Name -contains 'op' -and $items[0].PSObject.Properties.Name -contains 'path'
    if ($isPatch) { Invoke-CWRequest PATCH "sales/opportunities/$Id" $items; return }
    $merged=Merge-OpportunityObject $CurrentOpportunity $Opportunity
    $merged.id=$Id
    Invoke-CWRequest PUT "sales/opportunities/$Id" $merged
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
    if ($explicit -match '^(?i:newclient|new client|newcustomer|new customer|sales)$') { return 'NewClient' }
    if ($explicit -match '^(?i:existingclient|existing client|existingcustomer|existing customer|accountmanagement|account management)$') { return 'ExistingClient' }
    if ($explicit) { throw (New-RequestException "workspaceCategory '$explicit' is invalid. Use NewClient or ExistingClient.") }
    $typeText=Get-PropertyValue $Opportunity @('type.name','type.identifier')
    if ($typeText -match $NewClientTypeRegex) { return 'NewClient' }
    if ($typeText -match $ExistingClientTypeRegex) { return 'ExistingClient' }
    if (Test-WorkspaceCategory $DefaultCategory) { return $DefaultCategory }
    throw (New-RequestException "Could not classify opportunity type '$typeText'. Supply workspaceCategory or configure the type regex settings.")
}
function Get-WorkspaceRoute {
    param([string]$Category)
    if ($Category -eq 'NewClient') { return @{SiteUrl=$SalesSiteUrl;LibraryName=$SalesLibraryName} }
    return @{SiteUrl=$AccountSiteUrl;LibraryName=$AccountLibraryName}
}
function Assert-WorkspaceUrlForRoute {
    param([string]$Url,[object]$Route)
    $uri=Assert-HttpsUrl 'Existing SharePoint URL' $Url
    $siteUri=[Uri]$Route.SiteUrl
    $libraryPath="$($siteUri.AbsolutePath.TrimEnd('/'))/$([Uri]::EscapeDataString($Route.LibraryName))/"
    if ($uri.Host -ne $siteUri.Host -or -not $uri.AbsolutePath.StartsWith($libraryPath,[StringComparison]::OrdinalIgnoreCase)) {
        throw (New-RequestException 'Existing SharePoint URL must point to a folder in the configured opportunity library.')
    }
    return $uri
}
function Connect-WorkspaceSite {
    param([string]$SiteUrl)
    try {
        $securePassword=ConvertTo-SecureString $CertificatePassword -AsPlainText -Force
        Connect-PnPOnline -Url $SiteUrl -ClientId $SharePointClientId -Tenant $TenantId -CertificateBase64Encoded $CertificateBase64 -CertificatePassword $securePassword -ReturnConnection
    } catch {
        $_.Exception.Data['HttpStatusCode']=502
        throw
    }
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
    Assert-RequiredValue 'Opportunity number/id' $number 502
    Assert-RequiredValue 'Opportunity name' $name 502

    # Flat hierarchy:
    # <Opportunity Library>/<Opportunity Number - Opportunity Name>
    $opportunityFolderName=ConvertTo-SafeFolderName "$number - $name"
    $targetPath="$($route.LibraryName)/$opportunityFolderName"
    $existingUrl=Get-WorkspaceFieldValue $Opportunity
    $connection=Connect-WorkspaceSite $route.SiteUrl
    try {
        if (-not $DryRun) { Ensure-WorkspaceLibraryFields $connection $route.LibraryName }
        if ($existingUrl) {
            $existingUri=Assert-WorkspaceUrlForRoute $existingUrl $route
            $folderPath=[Uri]::UnescapeDataString($existingUri.AbsolutePath)
            $folder=Get-PnPFolder -Url $folderPath -Connection $connection -Includes ServerRelativeUrl,ListItemAllFields -ErrorAction Stop
            $metadata=$null
            if (-not $DryRun) { $metadata=Set-WorkspaceMetadata $connection $route.LibraryName $folder $Opportunity $RequestBody }
            return @{Url=$existingUrl;Created=$false;Path=$folder.ServerRelativeUrl;Metadata=$metadata;Note='Existing opportunity folder verified and metadata refreshed. Existing linked folders are not automatically moved.'}
        }
        if ($DryRun) { return @{Url=$null;Created=$false;Path=$targetPath;Metadata=$null;Note='Dry run: opportunity folder not created.'} }
        $folder=$null
        try {
            $folder=Get-PnPFolder -Url $targetPath -Connection $connection -Includes ServerRelativeUrl,ListItemAllFields -ErrorAction Stop
        } catch {
            $isNotFound=$_.Exception.Message -match '(?i)(404|file not found|folder not found|does not exist)'
            if (-not $isNotFound) { throw }
        }
        $created=$null -eq $folder
        if ($created) { $folder=Resolve-PnPFolder -SiteRelativePath $targetPath -Connection $connection -Includes ServerRelativeUrl,ListItemAllFields }
        $web=Get-PnPWeb -Connection $connection -Includes Url
        $url=([Uri]::new([Uri]$web.Url,$folder.ServerRelativeUrl)).AbsoluteUri
        $metadata=Set-WorkspaceMetadata $connection $route.LibraryName $folder $Opportunity $RequestBody
        $note=if ($created) {'Flat opportunity folder created and metadata updated.'} else {'Existing unlinked opportunity folder reused and metadata updated.'}
        return @{Url=$url;Created=$created;Path=$folder.ServerRelativeUrl;Metadata=$metadata;Note=$note}
    } catch {
        if ($null -eq $_.Exception.Data['HttpStatusCode']) { $_.Exception.Data['HttpStatusCode']=502 }
        throw
    } finally { Disconnect-PnPOnline -Connection $connection -ErrorAction SilentlyContinue }
}

$statusCode=200; $result=$null
$correlationId=[Guid]::NewGuid().ToString()
try {
    Import-Module PnP.PowerShell -ErrorAction Stop
    $body=Get-RequestBody $Request
    if ($null -eq $body -or ($body -isnot [PSCustomObject] -and $body -isnot [System.Collections.IDictionary])) {
        throw (New-RequestException 'Request body must be a JSON object.')
    }
    $mode=Get-PropertyValue $body @('mode','action')
    if ($mode -notmatch '^(?i:create|update)$') { throw (New-RequestException "mode must be 'Create' or 'Update'.") }
    if ($null -eq $body.opportunity) { throw (New-RequestException 'The request body must include opportunity.') }
    $incomingCorrelationId=[Guid]::Empty
    if ($Request.Headers -and [Guid]::TryParse([string]$Request.Headers['x-correlation-id'],[ref]$incomingCorrelationId)) {
        $correlationId=$incomingCorrelationId.ToString()
    }
    Write-Information "CorrelationId=$correlationId; Operation=$mode" -InformationAction Continue
    $providedLink=Get-PropertyValue $body @('connectWiseLink')
    if ($providedLink) { Assert-HttpsUrl 'connectWiseLink' $providedLink | Out-Null }

    if ($mode -match '^(?i:create)$') {
        if ($DryRun) { throw (New-RequestException 'Dry-run Create is blocked because ConnectWise must assign the opportunity number.') }
        if ($body.opportunity -isnot [PSCustomObject] -and $body.opportunity -isnot [System.Collections.IDictionary]) {
            throw (New-RequestException 'Create opportunity must be a non-empty JSON object.')
        }
        Assert-RequiredValue 'opportunity.name' (Get-PropertyValue $body.opportunity @('name'))
        $category=Resolve-WorkspaceCategory $body $body.opportunity
        $opportunity=New-CWOpportunity $body.opportunity
        $operation='Created'; $statusCode=201
    } else {
        $idText=Get-PropertyValue $body @('opportunityId','id')
        $parsedOpportunityId=0
        if (-not $idText -or -not [int]::TryParse($idText,[ref]$parsedOpportunityId) -or $parsedOpportunityId -lt 1) {
            throw (New-RequestException 'Update requires a positive numeric opportunityId within the supported range.')
        }
        $patchInput=$body.opportunity -is [System.Collections.IEnumerable] -and $body.opportunity -isnot [string]
        if ($patchInput) {
            $patchItems=@($body.opportunity)
            if ($patchItems.Count -eq 0) { throw (New-RequestException 'JSON Patch updates must be a non-empty array.') }
            $changesOpportunityType=$false
            foreach ($patchItem in $patchItems) {
                $patchOp=Get-PropertyValue $patchItem @('op')
                $patchPath=Get-PropertyValue $patchItem @('path')
                if ($patchItem -isnot [PSCustomObject] -or $patchOp -notin @('add','remove','replace','move','copy','test') -or -not $patchPath) {
                    throw (New-RequestException 'Each JSON Patch operation must have a supported op and non-empty path.')
                }
                if ($patchOp -in @('add','replace','test') -and $patchItem.PSObject.Properties.Name -notcontains 'value') {
                    throw (New-RequestException "JSON Patch operation '$patchOp' requires a value.")
                }
                if ($patchOp -in @('move','copy') -and -not (Get-PropertyValue $patchItem @('from'))) {
                    throw (New-RequestException "JSON Patch operation '$patchOp' requires a from path.")
                }
                if ($patchPath -match '^(?i:/?type(?:[./]|$))') { $changesOpportunityType=$true }
            }
            if ($changesOpportunityType -and -not (Get-PropertyValue $body @('workspaceCategory'))) {
                throw (New-RequestException 'JSON Patch updates that change opportunity type must include workspaceCategory.')
            }
        } elseif (($body.opportunity -isnot [PSCustomObject] -and $body.opportunity -isnot [System.Collections.IDictionary]) -or $body.opportunity.PSObject.Properties.Count -eq 0) {
            throw (New-RequestException 'Update opportunity must be an object or a non-empty JSON Patch array.')
        }
        $bodyOpportunityId=Get-PropertyValue $body.opportunity @('id')
        if ($bodyOpportunityId -and $bodyOpportunityId -ne [string]$parsedOpportunityId) {
            throw (New-RequestException 'opportunity.id must match opportunityId.')
        }
        $opportunity=Get-CWOpportunity $parsedOpportunityId
        if ($patchInput) {
            $category=Resolve-WorkspaceCategory $body $opportunity
        } else {
            $preview=Merge-OpportunityObject ($opportunity | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100) $body.opportunity
            $category=Resolve-WorkspaceCategory $body $preview
        }
        if (-not $DryRun) {
            $existingWorkspaceUrl=Get-WorkspaceFieldValue $opportunity
            if ($existingWorkspaceUrl) { Assert-WorkspaceUrlForRoute $existingWorkspaceUrl (Get-WorkspaceRoute $category) | Out-Null }
            Update-CWOpportunity $parsedOpportunityId $body.opportunity $opportunity | Out-Null
            $opportunity=Get-CWOpportunity $parsedOpportunityId
        }
        $operation=if ($DryRun) {'UpdateDryRun'} else {'Updated'}
    }

    $workspace=Ensure-Workspace $opportunity $category $body
    $currentUrl=Get-WorkspaceFieldValue $opportunity
    if ($workspace.Url -and $currentUrl -ne $workspace.Url) { Set-CWWorkspaceUrl $opportunity $workspace.Url }
    $finalOpportunity=if ($DryRun) {$opportunity} else {Get-CWOpportunity ([int]$opportunity.id)}

    $result=[ordered]@{
        success=$true; operation=$operation; dryRun=$DryRun
        opportunityId=$finalOpportunity.id; opportunityName=$finalOpportunity.name
        correlationId=$correlationId
        connectWiseCustomFieldId=$WorkspaceCustomFieldId
        workspaceCategory=$category; workspaceUrl=$workspace.Url
        workspaceCreated=$workspace.Created; workspacePath=$workspace.Path
        sharePointMetadata=$workspace.Metadata; note=$workspace.Note
    }
} catch {
    $statusCode=500
    if ($_.Exception.Data['HttpStatusCode']) { $statusCode=[int]$_.Exception.Data['HttpStatusCode'] }
    $result=[ordered]@{success=$false;error=$_.Exception.Message;type=$_.Exception.GetType().FullName;correlationId=$correlationId}
    Write-Error ("CorrelationId={0}; HttpStatus={1}; {2}" -f $correlationId,$statusCode,$_.Exception.Message) -ErrorAction Continue
}
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode=$statusCode
    Headers=@{'Content-Type'='application/json'}
    Body=($result | ConvertTo-Json -Depth 20)
})
