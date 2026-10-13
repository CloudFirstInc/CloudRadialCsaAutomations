using namespace System.Net

param($Request, $TriggerMetadata)

$CorrelationId = [guid]::NewGuid().ToString()
$DebugEnabled = ($env:DebugLogging -eq '1')

function Write-Info { param([string]$Message) Write-Host ('[{0}] [INFO] {1}' -f $CorrelationId, $Message) }
function Write-DebugLog { param([string]$Message) if ($DebugEnabled) { Write-Host ('[{0}] [DEBUG] {1}' -f $CorrelationId, $Message) } }
function Write-ErrorLog { param([string]$Message) Write-Host ('[{0}] [ERROR] {1}' -f $CorrelationId, $Message) }

function Send-JsonResponse {
    param([System.Net.HttpStatusCode]$StatusCode, [object]$Body)
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = $StatusCode
        Body = ($Body | ConvertTo-Json -Depth 100)
        ContentType = 'application/json'
    })
}

function Get-RequestBodyObject {
    param($Request)
    if ($null -eq $Request.Body) { return $null }
    if ($Request.Body -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Request.Body)) { return $null }
        try { return $Request.Body | ConvertFrom-Json }
        catch { throw [System.ArgumentException]::new('Request body is not valid JSON.') }
    }
    return $Request.Body
}

function Get-InputValue {
    param($Request, $Body, [string[]]$Names)
    foreach ($name in $Names) {
        if ($Body) {
            $prop = $Body.PSObject.Properties | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
            if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) { return [string]$prop.Value }
        }
        if ($Request.Query) {
            foreach ($key in $Request.Query.Keys) {
                if ([string]$key -ieq $name -and -not [string]::IsNullOrWhiteSpace([string]$Request.Query[$key])) { return [string]$Request.Query[$key] }
            }
        }
        if ($Request.Headers) {
            foreach ($key in $Request.Headers.Keys) {
                if ([string]$key -ieq $name -and -not [string]::IsNullOrWhiteSpace([string]$Request.Headers[$key])) { return [string]$Request.Headers[$key] }
            }
        }
    }
    return $null
}

function Get-BoolValue {
    param([object]$Value, [bool]$DefaultValue)
    if ($null -eq $Value) { return $DefaultValue }
    if ($Value -is [bool]) { return $Value }
    if ([string]$Value -match '^(true|1|yes|y)$') { return $true }
    if ([string]$Value -match '^(false|0|no|n)$') { return $false }
    return $DefaultValue
}

function Assert-RequiredSettings {
    param([string[]]$Names)
    $missing = @()
    foreach ($name in $Names) {
        $value = [Environment]::GetEnvironmentVariable($name)
        if ([string]::IsNullOrWhiteSpace($value)) { $missing += $name }
    }
    if ($missing.Count -gt 0) {
        $ex = [System.InvalidOperationException]::new('Missing required application settings.')
        $ex.Data['MissingSettings'] = $missing
        throw $ex
    }
}

function Invoke-WithRetry {
    param([scriptblock]$Operation, [int]$MaxRetries = 5, [string]$OperationName = 'HTTP request')
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try { return & $Operation }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
            $retryable = ($statusCode -in @(408, 429, 500, 502, 503, 504))
            if (-not $retryable -or $attempt -ge $MaxRetries) { throw }
            $delay = [Math]::Min(30, [int][Math]::Pow(2, $attempt) + (Get-Random -Minimum 0 -Maximum 3))
            Write-Info ('{0} failed with status {1}; retrying in {2} seconds.' -f $OperationName, $statusCode, $delay)
            Start-Sleep -Seconds $delay
        }
    }
}

function New-ConnectWiseHeaders {
    $auth = '{0}+{1}:{2}' -f $env:ConnectWisePsa_ApiCompanyId, $env:ConnectWisePsa_ApiPublicKey, $env:ConnectWisePsa_ApiPrivateKey
    $encoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($auth))
    return @{
        Authorization = 'Basic {0}' -f $encoded
        ClientId = $env:ConnectWisePsa_ApiClientId
        Accept = 'application/json'
        'Content-Type' = 'application/json'
    }
}

function Invoke-ConnectWiseJson {
    param([string]$Path, [hashtable]$Headers, [switch]$Optional)
    $uri = '{0}/v4_6_release/apis/3.0/{1}' -f $env:ConnectWisePsa_ApiBaseUrl.TrimEnd('/'), $Path.TrimStart('/')
    try {
        return Invoke-WithRetry -OperationName ('ConnectWise GET {0}' -f $Path) -Operation {
            Invoke-RestMethod -Uri $uri -Method Get -Headers $Headers -TimeoutSec 100
        }
    }
    catch {
        if ($Optional) {
            Write-Info ('Optional ConnectWise resource unavailable: {0}' -f $Path)
            return @()
        }
        throw
    }
}

function Get-ConnectWiseDocumentBytes {
    param([int]$DocumentId, [hashtable]$Headers)
    $uri = '{0}/v4_6_release/apis/3.0/system/documents/{1}/download' -f $env:ConnectWisePsa_ApiBaseUrl.TrimEnd('/'), $DocumentId
    $tempFile = Join-Path ([IO.Path]::GetTempPath()) ('cw-{0}-{1}.bin' -f $DocumentId, [guid]::NewGuid())
    try {
        Invoke-WithRetry -OperationName ('ConnectWise document download {0}' -f $DocumentId) -Operation {
            Invoke-WebRequest -Uri $uri -Method Get -Headers $Headers -OutFile $tempFile -TimeoutSec 300 | Out-Null
        } | Out-Null
        return [IO.File]::ReadAllBytes($tempFile)
    }
    finally {
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
    }
}

function Get-GraphToken {
    $tokenUri = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $env:SharePoint_TenantId
    $form = @{
        client_id = $env:SharePoint_ClientId
        client_secret = $env:SharePoint_ClientSecret
        scope = 'https://graph.microsoft.com/.default'
        grant_type = 'client_credentials'
    }
    $token = Invoke-RestMethod -Uri $tokenUri -Method Post -Body $form -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 100
    return $token.access_token
}

function Invoke-GraphJson {
    param([string]$Uri, [string]$Method = 'Get', [string]$Token, [object]$Body = $null)
    $headers = @{ Authorization = 'Bearer {0}' -f $Token; Accept = 'application/json' }
    return Invoke-WithRetry -OperationName ('Graph {0}' -f $Method) -Operation {
        if ($null -eq $Body) { return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $headers -TimeoutSec 100 }
        $json = $Body | ConvertTo-Json -Depth 30
        return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $headers -ContentType 'application/json' -Body $json -TimeoutSec 100
    }
}

function ConvertTo-SafeFolderName {
    param([string]$Name, [int]$MaxLength = 120)
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'Unnamed' }
    $safe = $Name.Trim() -replace '["*:<>?/\\|#%]', '-'
    $safe = $safe -replace '\s+', ' '
    $safe = $safe.Trim('.', ' ')
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength).Trim() }
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'Unnamed' }
    return $safe
}

function Get-GraphChildByName {
    param([string]$DriveId, [string]$ParentId, [string]$Name, [string]$Token)
    $safeName = $Name.Replace("'", "''")
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ParentId/children?`$filter=name eq '$safeName'&`$select=id,name,webUrl,folder,file"
    try {
        $result = Invoke-GraphJson -Uri $uri -Token $Token
        return @($result.value) | Select-Object -First 1
    }
    catch { return $null }
}

function Ensure-GraphFolder {
    param([string]$DriveId, [string]$ParentId, [string]$Name, [string]$Token)
    $existing = Get-GraphChildByName -DriveId $DriveId -ParentId $ParentId -Name $Name -Token $Token
    if ($existing) {
        if (-not $existing.folder) { throw [System.InvalidOperationException]::new("A file already exists where folder '$Name' is required.") }
        return $existing
    }
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ParentId/children"
    $body = @{ name = $Name; folder = @{}; '@microsoft.graph.conflictBehavior' = 'fail' }
    try { return Invoke-GraphJson -Uri $uri -Method Post -Token $Token -Body $body }
    catch {
        $existing = Get-GraphChildByName -DriveId $DriveId -ParentId $ParentId -Name $Name -Token $Token
        if ($existing) { return $existing }
        throw
    }
}

function Ensure-GraphFolderPath {
    param([string]$DriveId, [string]$RootItemId, [string]$Path, [string]$Token)
    $current = [PSCustomObject]@{ id = $RootItemId; webUrl = $null; name = 'root' }
    foreach ($segment in @($Path -split '/' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $current = Ensure-GraphFolder -DriveId $DriveId -ParentId $current.id -Name (ConvertTo-SafeFolderName $segment) -Token $Token
    }
    return $current
}

function Upload-GraphBytes {
    param([string]$DriveId, [string]$ParentId, [string]$FileName, [byte[]]$Bytes, [string]$Token)
    $encodedName = [Uri]::EscapeDataString((ConvertTo-SafeFolderName $FileName -MaxLength 180))
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ParentId`:/`:$encodedName`:/content" -replace '/items/([^:]+):/:', '/items/$1:/'
    $headers = @{ Authorization = 'Bearer {0}' -f $Token; 'Content-Type' = 'application/octet-stream' }
    return Invoke-WithRetry -OperationName ('Graph upload {0}' -f $FileName) -Operation {
        Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $Bytes -TimeoutSec 300
    }
}

function Get-OpportunityDocuments {
    param([int]$OpportunityId, [hashtable]$Headers)
    $recordTypes = @()
    if (-not [string]::IsNullOrWhiteSpace($env:ConnectWisePsa_OpportunityRecordType)) { $recordTypes += $env:ConnectWisePsa_OpportunityRecordType }
    $recordTypes += @('Opportunity', 'SalesOpportunity')
    foreach ($recordType in ($recordTypes | Select-Object -Unique)) {
        $conditions = [Uri]::EscapeDataString("recordType='$recordType' and recordId=$OpportunityId")
        $docs = @(Invoke-ConnectWiseJson -Path ("system/documents?conditions={0}&pageSize=1000" -f $conditions) -Headers $Headers -Optional)
        if ($docs.Count -gt 0) { return [PSCustomObject]@{ recordType = $recordType; documents = $docs } }
    }
    return [PSCustomObject]@{ recordType = $null; documents = @() }
}

try {
    Write-Info 'Starting ConnectWise opportunity package function.'
    $body = Get-RequestBodyObject -Request $Request
    $opportunityIdText = Get-InputValue -Request $Request -Body $body -Names @('OpportunityId', 'opportunityId', 'id')
    $securityKey = Get-InputValue -Request $Request -Body $body -Names @('SecurityKey', 'securityKey')
    $includeAttachments = Get-BoolValue -Value (Get-InputValue -Request $Request -Body $body -Names @('IncludeAttachments', 'includeAttachments')) -DefaultValue $true
    $overwriteExistingFiles = Get-BoolValue -Value (Get-InputValue -Request $Request -Body $body -Names @('OverwriteExistingFiles', 'overwriteExistingFiles')) -DefaultValue $false

    if (-not [string]::IsNullOrWhiteSpace($env:SecurityKey) -and -not [string]::Equals($env:SecurityKey, $securityKey, [StringComparison]::Ordinal)) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::Unauthorized) -Body @{ success = $false; correlationId = $CorrelationId; errorCode = 'InvalidSecurityKey'; error = 'Invalid security key.' }
        return
    }

    $opportunityId = 0
    if (-not [int]::TryParse($opportunityIdText, [ref]$opportunityId) -or $opportunityId -lt 1) {
        throw [System.ArgumentException]::new('OpportunityId must be a positive integer.')
    }

    Assert-RequiredSettings @(
        'ConnectWisePsa_ApiBaseUrl','ConnectWisePsa_ApiCompanyId','ConnectWisePsa_ApiPublicKey','ConnectWisePsa_ApiPrivateKey','ConnectWisePsa_ApiClientId',
        'SharePoint_TenantId','SharePoint_ClientId','SharePoint_ClientSecret','SharePoint_DriveId','SharePoint_RootFolderPath'
    )

    $cwHeaders = New-ConnectWiseHeaders
    $opportunity = Invoke-ConnectWiseJson -Path ("sales/opportunities/{0}" -f $opportunityId) -Headers $cwHeaders

    $details = [ordered]@{
        opportunity = $opportunity
        notes = @(Invoke-ConnectWiseJson -Path ("sales/opportunities/{0}/notes?pageSize=1000" -f $opportunityId) -Headers $cwHeaders -Optional)
        products = @(Invoke-ConnectWiseJson -Path ("sales/opportunities/{0}/products?pageSize=1000" -f $opportunityId) -Headers $cwHeaders -Optional)
        team = @(Invoke-ConnectWiseJson -Path ("sales/opportunities/{0}/team?pageSize=1000" -f $opportunityId) -Headers $cwHeaders -Optional)
        forecast = @(Invoke-ConnectWiseJson -Path ("sales/opportunities/{0}/forecast?pageSize=1000" -f $opportunityId) -Headers $cwHeaders -Optional)
        activities = @(Invoke-ConnectWiseJson -Path ("sales/activities?conditions={0}&pageSize=1000" -f [Uri]::EscapeDataString("opportunity/id=$opportunityId")) -Headers $cwHeaders -Optional)
    }

    $docResult = Get-OpportunityDocuments -OpportunityId $opportunityId -Headers $cwHeaders
    $details.documents = @($docResult.documents)

    $token = Get-GraphToken
    $driveId = $env:SharePoint_DriveId
    $rootItemId = if ([string]::IsNullOrWhiteSpace($env:SharePoint_RootItemId)) { 'root' } else { $env:SharePoint_RootItemId }
    $baseFolder = Ensure-GraphFolderPath -DriveId $driveId -RootItemId $rootItemId -Path $env:SharePoint_RootFolderPath -Token $token

    $companyName = if ($opportunity.company.name) { [string]$opportunity.company.name } else { 'Unknown Client' }
    $opportunityName = if ($opportunity.name) { [string]$opportunity.name } else { 'Opportunity' }
    $folderName = ConvertTo-SafeFolderName ('{0} - {1} - {2}' -f $opportunityId, $companyName, $opportunityName)
    $opportunityFolder = Ensure-GraphFolder -DriveId $driveId -ParentId $baseFolder.id -Name $folderName -Token $token
    $attachmentsFolder = Ensure-GraphFolder -DriveId $driveId -ParentId $opportunityFolder.id -Name 'Attachments' -Token $token

    $attachmentResults = @()
    if ($includeAttachments) {
        foreach ($document in @($docResult.documents)) {
            $documentId = 0
            if (-not [int]::TryParse([string]$document.id, [ref]$documentId)) { continue }
            $fileName = if ($document.fileName) { [string]$document.fileName } elseif ($document.title) { [string]$document.title } else { 'ConnectWise-Document-{0}' -f $documentId }
            $existing = Get-GraphChildByName -DriveId $driveId -ParentId $attachmentsFolder.id -Name $fileName -Token $token
            if ($existing -and -not $overwriteExistingFiles) {
                $attachmentResults += [PSCustomObject]@{ documentId = $documentId; fileName = $fileName; status = 'SkippedExisting'; webUrl = $existing.webUrl; size = $existing.size }
                continue
            }
            try {
                $bytes = Get-ConnectWiseDocumentBytes -DocumentId $documentId -Headers $cwHeaders
                $uploaded = Upload-GraphBytes -DriveId $driveId -ParentId $attachmentsFolder.id -FileName $fileName -Bytes $bytes -Token $token
                $attachmentResults += [PSCustomObject]@{ documentId = $documentId; fileName = $uploaded.name; status = 'Uploaded'; webUrl = $uploaded.webUrl; size = $uploaded.size }
            }
            catch {
                Write-ErrorLog ('Attachment {0} failed: {1}' -f $documentId, $_.Exception.Message)
                $attachmentResults += [PSCustomObject]@{ documentId = $documentId; fileName = $fileName; status = 'Failed'; error = 'The attachment could not be copied.' }
            }
        }
    }

    $metadata = [ordered]@{
        opportunityId = $opportunityId
        companyName = $companyName
        opportunityName = $opportunityName
        connectWiseDocumentRecordType = $docResult.recordType
        sharePointFolderUrl = $opportunityFolder.webUrl
        sharePointAttachmentsFolderUrl = $attachmentsFolder.webUrl
        retrievedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        correlationId = $CorrelationId
    }
    $metadataBytes = [Text.Encoding]::UTF8.GetBytes(([PSCustomObject]@{ metadata = $metadata; opportunityData = $details } | ConvertTo-Json -Depth 100))
    $metadataUpload = Upload-GraphBytes -DriveId $driveId -ParentId $opportunityFolder.id -FileName 'opportunity-data.json' -Bytes $metadataBytes -Token $token

    Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
        success = $true
        correlationId = $CorrelationId
        opportunityId = $opportunityId
        folder = @{ id = $opportunityFolder.id; name = $opportunityFolder.name; webUrl = $opportunityFolder.webUrl; attachmentsWebUrl = $attachmentsFolder.webUrl }
        opportunityData = $details
        attachments = $attachmentResults
        metadataFile = @{ name = $metadataUpload.name; webUrl = $metadataUpload.webUrl }
        retrievedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    }
}
catch {
    Write-ErrorLog $_.Exception.Message
    $status = [HttpStatusCode]::InternalServerError
    $code = 'UnhandledFunctionError'
    $message = 'An unexpected error occurred while processing the opportunity.'
    if ($_.Exception -is [System.ArgumentException]) { $status = [HttpStatusCode]::BadRequest; $code = 'InvalidRequestArgument'; $message = $_.Exception.Message }
    elseif ($_.Exception.Data['MissingSettings']) { $code = 'MissingApplicationSettings'; $message = 'Missing required application settings.' }
    else {
        $remoteStatus = $null
        try { $remoteStatus = [int]$_.Exception.Response.StatusCode } catch { }
        if ($remoteStatus -eq 404) { $status = [HttpStatusCode]::NotFound; $code = 'ResourceNotFound'; $message = 'The ConnectWise opportunity or related resource was not found.' }
        elseif ($remoteStatus -in @(401,403)) { $status = [HttpStatusCode]::Forbidden; $code = 'UpstreamAccessDenied'; $message = 'ConnectWise or SharePoint denied access.' }
    }
    $response = @{ success = $false; correlationId = $CorrelationId; errorCode = $code; error = $message }
    if ($_.Exception.Data['MissingSettings']) { $response.missingSettings = @($_.Exception.Data['MissingSettings']) }
    Send-JsonResponse -StatusCode $status -Body $response
}
