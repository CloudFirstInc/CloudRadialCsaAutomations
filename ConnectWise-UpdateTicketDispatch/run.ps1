<#
.SYNOPSIS
    Creates or updates ConnectWise Manage tickets for the CloudFirst Copilot Dispatch Agent.

.DESCRIPTION
    This Azure Function is the write-side companion to the CloudFirst read-only ticket context function.
    It updates approved ConnectWise Manage service-ticket fields, creates tickets when explicitly requested,
    and adds audit notes. It is designed for Copilot Studio actions where AI recommends changes and a
    human dispatcher/manager approves before execution.

    Copilot enhancements included:
    - Pre-flight dry-run validation ("DryRun": true) without writing to ConnectWise
    - Acceptance of BoardName, StatusName, PriorityName, and CompanyName in addition to numeric IDs
    - Board guardrail validation on both ticket creation and destination board transfers
    - Standardized errorCode properties in all response payloads for easy Power Fx condition branching
    - Exception sanitization to avoid leaking internal error details

.REQUIRED ENVIRONMENT VARIABLES
    ConnectWisePsa_ApiBaseUrl      - Base URL of the ConnectWise Manage API
    ConnectWisePsa_ApiCompanyId    - ConnectWise company ID
    ConnectWisePsa_ApiPublicKey    - ConnectWise public key
    ConnectWisePsa_ApiPrivateKey   - ConnectWise private key
    ConnectWisePsa_ApiClientId     - ConnectWise client ID
    SecurityKey                    - Optional shared secret to secure this function

.OPTIONAL ENVIRONMENT VARIABLES
    DebugLogging                         - Set to 1 for verbose diagnostic logs
    ConnectWisePsa_RequireApproval       - Default 1. Set to 0 only for controlled system automation.
    ConnectWisePsa_AllowCloseTickets     - Default 0. Set to 1 to allow ticket closure (ClosedFlag=true or a StatusId whose board status closes the ticket).
    ConnectWisePsa_AllowedServiceBoards  - Comma-separated board names allowed for updates
    ConnectWisePsa_AllowedProjectBoards  - Comma-separated project board names allowed for updates
    ConnectWisePsa_MaxPatchOperations    - Default 25

.INPUTS
    JSON body examples:

    Update ticket (by ID or Name, with optional DryRun):
    {
        "Action": "UpdateTicket",
        "TicketId": 123456,
        "SecurityKey": "optional",
        "Approved": true,
        "ApprovedBy": "Dispatcher Name",
        "ApprovalReason": "Dispatcher accepted Copilot recommendation",
        "DryRun": false,
        "Changes": {
            "Summary": "Updated summary",
            "BoardName": "Service Desk",
            "StatusName": "In Progress",
            "PriorityName": "Priority 2 - High",
            "TypeId": 34,
            "SubTypeId": 56,
            "ItemId": 78,
            "AgreementId": 1001,
            "OwnerId": 99,
            "TeamId": 4,
            "InitialDescription": "Updated description"
        },
        "AuditNote": "Copilot recommendation approved by dispatcher."
    }

    Add note:
    {
        "Action": "AddNote",
        "TicketId": 123456,
        "Approved": true,
        "ApprovedBy": "Dispatcher Name",
        "NoteText": "Internal dispatcher note",
        "NoteType": "InternalAnalysis",
        "DryRun": false
    }

    Create ticket (accepts CompanyName or CompanyId, BoardName or BoardId):
    {
        "Action": "CreateTicket",
        "Approved": true,
        "ApprovedBy": "Dispatcher Name",
        "DryRun": false,
        "Ticket": {
            "Summary": "New service request",
            "BoardName": "Service Desk",
            "CompanyName": "City of Rye",
            "StatusName": "New",
            "PriorityName": "Priority 3 - Normal",
            "ContactId": 456,
            "TypeId": 11,
            "SubTypeId": 12,
            "ItemId": 13,
            "InitialDescription": "Problem description"
        },
        "AuditNote": "Ticket created from approved Copilot action."
    }

.OUTPUTS
    JSON object with success/failure, correlationId, action, ticketId, before/after values, and audit note result.

.NOTES
    This function performs writes. Keep it separate from the read-only context function.
    Recommended Copilot design: AI recommends, human approves, then this function executes.
#>

using namespace System.Net

param($Request, $TriggerMetadata)

# -------------------------------
# Utility: Logging
# -------------------------------

$CorrelationId = [guid]::NewGuid().ToString()
$DebugEnabled = ($env:DebugLogging -eq "1")

function Write-Info {
    param([string]$Message)
    Write-Host ("[{0}] [INFO] {1}" -f $CorrelationId, $Message)
}

function Write-DebugLog {
    param([string]$Message)
    if ($DebugEnabled) {
        Write-Host ("[{0}] [DEBUG] {1}" -f $CorrelationId, $Message)
    }
}

function Write-ErrorLog {
    param([string]$Message)
    Write-Host ("[{0}] [ERROR] {1}" -f $CorrelationId, $Message)
}

# -------------------------------
# Utility: Response
# -------------------------------

function Send-JsonResponse {
    param(
        [System.Net.HttpStatusCode]$StatusCode,
        [object]$Body
    )

    $json = $Body | ConvertTo-Json -Depth 80

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = $StatusCode
        Body        = $json
        ContentType = "application/json"
    })
}

# -------------------------------
# Utility: Safe property getter
# -------------------------------

function Get-ValueOrNull {
    param(
        [object]$Object,
        [string]$PropertyName
    )

    if ($null -eq $Object) { return $null }

    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) {
            if ([string]$key -eq $PropertyName) { return $Object[$key] }
        }
        return $null
    }

    $prop = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function Test-HasProperty {
    param(
        [object]$Object,
        [string]$PropertyName
    )

    if ($null -eq $Object) { return $false }

    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) {
            if ([string]$key -eq $PropertyName) { return $true }
        }
        return $false
    }

    return ($null -ne $Object.PSObject.Properties[$PropertyName])
}

function Invoke-BadRequest {
    param([string]$Message)

    throw [System.ArgumentException]::new($Message)
}

function ConvertTo-CwConditionString {
    param([string]$Value)

    if ($null -eq $Value) { return "" }
    return $Value.Replace("'", "''")
}

function ConvertTo-CwUrlEncoded {
    param([string]$Value)

    if ($null -eq $Value) { return "" }
    return [System.Uri]::EscapeDataString($Value)
}

# -------------------------------
# Utility: Build headers
# -------------------------------

function New-ConnectWiseHeaders {
    param(
        [string]$CompanyId,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId
    )

    $authString = "{0}+{1}:{2}" -f $CompanyId, $PublicKey, $PrivateKey
    $encodedAuth = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($authString))

    return @{
        "Authorization" = "Basic {0}" -f $encodedAuth
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
        "ClientId"      = $ClientId
    }
}

# -------------------------------
# Utility: Invoke ConnectWise API with retry
# -------------------------------

function Invoke-ConnectWiseApi {
    param(
        [string]$Uri,
        [string]$Method = "Get",
        [hashtable]$Headers,
        [object]$Body = $null,
        [int]$MaxRetries = 5
    )

    $attempt = 0

    while ($true) {
        try {
            $attempt++
            Write-DebugLog ("CW API {0} {1}, attempt {2}" -f $Method, $Uri, $attempt)

            if ($null -ne $Body) {
                $bodyJson = $Body | ConvertTo-Json -Depth 40
                Write-DebugLog ("Body: {0}" -f $bodyJson)
                $webResponse = Invoke-WebRequest -Uri $Uri -Method $Method -Headers $Headers -Body $bodyJson -ContentType "application/json" -TimeoutSec 100 -UseBasicParsing
            }
            else {
                $webResponse = Invoke-WebRequest -Uri $Uri -Method $Method -Headers $Headers -TimeoutSec 100 -UseBasicParsing
            }

            # Parse manually: Invoke-RestMethod can throw "Operation is not valid due to the current state of the object"
            # after a successful write when ConnectWise returns JSON it cannot deserialize (for example, keys differing only by case).
            $content = $webResponse.Content
            if ($content -is [byte[]]) { $content = [System.Text.Encoding]::UTF8.GetString($content) }
            if ([string]::IsNullOrWhiteSpace([string]$content)) { return $null }

            try { return ($content | ConvertFrom-Json -Depth 40) }
            catch {
                try { return ($content | ConvertFrom-Json -Depth 40 -AsHashtable) }
                catch {
                    Write-ErrorLog ("CW API {0} succeeded (HTTP {1}) but the response could not be parsed: {2}" -f $Method, $webResponse.StatusCode, $_.Exception.Message)
                    return [PSCustomObject]@{ rawResponse = [string]$content }
                }
            }
        }
        catch {
            $statusCode = $null
            $retryAfter = $null

            if ($_.Exception.Response) {
                try {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                catch {
                    $statusCode = $null
                }

                try {
                    $retryAfterHeader = $_.Exception.Response.Headers.RetryAfter
                    if ($retryAfterHeader) {
                        if ($retryAfterHeader.Delta) {
                            $retryAfter = [int]$retryAfterHeader.Delta.TotalSeconds
                        }
                        elseif ($retryAfterHeader.Date) {
                            $retryAfter = [int][Math]::Max(1, ($retryAfterHeader.Date - [DateTimeOffset]::UtcNow).TotalSeconds)
                        }
                    }
                }
                catch {
                    $retryAfter = $null
                }
            }

            $message = $_.Exception.Message
            Write-ErrorLog ("CW API call failed. StatusCode={0}. Message={1}" -f $statusCode, $message)

            $retryable = $false
            if ($statusCode -in @(408, 429, 500, 502, 503, 504)) {
                $retryable = $true
            }

            if (-not $retryable -or $attempt -ge $MaxRetries) {
                throw
            }

            if ($retryAfter) {
                $sleepSeconds = [int]$retryAfter
            }
            else {
                $sleepSeconds = [Math]::Pow(2, $attempt) + (Get-Random -Minimum 1 -Maximum 4)
            }

            Write-Info ("Retrying in {0} seconds" -f $sleepSeconds)
            Start-Sleep -Seconds $sleepSeconds
        }
    }
}

# -------------------------------
# Input parsing
# -------------------------------

function Get-RequestBodyObject {
    param($Request)

    if ($null -eq $Request.Body) { return $null }

    if ($Request.Body -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Request.Body)) { return $null }
        try { return $Request.Body | ConvertFrom-Json }
        catch { Invoke-BadRequest "Request body is not valid JSON." }
    }

    return $Request.Body
}

function Get-InputValue {
    param(
        $Request,
        $Body,
        [string[]]$Names
    )

    foreach ($name in $Names) {
        if ($Body) {
            if ($Body -is [System.Collections.IDictionary]) {
                if ($Body.Contains($name)) {
                    $val = $Body[$name]
                    if (-not [string]::IsNullOrWhiteSpace([string]$val)) { return [string]$val }
                }

                $matchedKey = $Body.Keys | Where-Object { $_ -eq $name -or $_.ToString().ToLower() -eq $name.ToLower() } | Select-Object -First 1
                if ($matchedKey) {
                    $val = $Body[$matchedKey]
                    if (-not [string]::IsNullOrWhiteSpace([string]$val)) { return [string]$val }
                }
            }
            else {
                $bodyProp = $Body.PSObject.Properties[$name]
                if ($bodyProp -and -not [string]::IsNullOrWhiteSpace([string]$bodyProp.Value)) {
                    return [string]$bodyProp.Value
                }

                $prop = $Body.PSObject.Properties | Where-Object { $_.Name -eq $name -or $_.Name.ToLower() -eq $name.ToLower() } | Select-Object -First 1
                if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) { return [string]$prop.Value }
            }
        }

        if ($Request.Query) {
            $queryValue = $Request.Query[$name]
            if (-not [string]::IsNullOrWhiteSpace([string]$queryValue)) { return [string]$queryValue }

            foreach ($key in $Request.Query.Keys) {
                if ($key -eq $name -or $key.ToString().ToLower() -eq $name.ToLower()) {
                    $qv = $Request.Query[$key]
                    if (-not [string]::IsNullOrWhiteSpace([string]$qv)) { return [string]$qv }
                }
            }
        }

        if ($Request.Headers) {
            $headerValue = $Request.Headers[$name]
            if (-not [string]::IsNullOrWhiteSpace([string]$headerValue)) { return [string]$headerValue }

            foreach ($key in $Request.Headers.Keys) {
                if ($key -eq $name -or $key.ToString().ToLower() -eq $name.ToLower()) {
                    $hv = $Request.Headers[$key]
                    if (-not [string]::IsNullOrWhiteSpace([string]$hv)) { return [string]$hv }
                }
            }
        }
    }

    return $null
}

function Get-BoolValue {
    param(
        [object]$Value,
        [bool]$DefaultValue
    )

    if ($null -eq $Value) { return $DefaultValue }
    if ($Value -is [bool]) { return $Value }

    $text = [string]$Value
    if ($text -match "^(true|1|yes|y)$") { return $true }
    if ($text -match "^(false|0|no|n)$") { return $false }
    return $DefaultValue
}

# -------------------------------
# ConnectWise ticket functions
# -------------------------------

function Get-ConnectWiseTicket {
    param(
        [string]$ConnectWiseUrl,
        [string]$TicketId,
        [hashtable]$Headers
    )

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/tickets/{1}" -f $ConnectWiseUrl.TrimEnd("/"), $TicketId
    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
}

function Update-ConnectWiseTicket {
    param(
        [string]$ConnectWiseUrl,
        [string]$TicketId,
        [hashtable]$Headers,
        [object[]]$PatchOperations
    )

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/tickets/{1}" -f $ConnectWiseUrl.TrimEnd("/"), $TicketId
    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Patch -Headers $Headers -Body $PatchOperations
}

function New-ConnectWiseTicket {
    param(
        [string]$ConnectWiseUrl,
        [hashtable]$Headers,
        [object]$TicketBody
    )

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/tickets" -f $ConnectWiseUrl.TrimEnd("/")
    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Post -Headers $Headers -Body $TicketBody
}

function Add-ConnectWiseTicketNote {
    param(
        [string]$ConnectWiseUrl,
        [string]$TicketId,
        [hashtable]$Headers,
        [string]$Text,
        [string]$NoteType = "InternalAnalysis"
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/tickets/{1}/notes" -f $ConnectWiseUrl.TrimEnd("/"), $TicketId

    $noteBody = @{
        text = $Text
    }

    switch ($NoteType.ToLower()) {
        "resolution" {
            $noteBody.resolutionFlag = $true
            $noteBody.detailDescriptionFlag = $false
            $noteBody.internalAnalysisFlag = $false
        }
        "discussion" {
            $noteBody.detailDescriptionFlag = $true
            $noteBody.internalAnalysisFlag = $false
            $noteBody.resolutionFlag = $false
        }
        default {
            $noteBody.internalAnalysisFlag = $true
            $noteBody.detailDescriptionFlag = $false
            $noteBody.resolutionFlag = $false
        }
    }

    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Post -Headers $Headers -Body $noteBody
}

function Get-ConnectWiseBoardStatuses {
    param(
        [string]$ConnectWiseUrl,
        [string]$BoardId,
        [hashtable]$Headers
    )

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/boards/{1}/statuses?pageSize=1000" -f $ConnectWiseUrl.TrimEnd("/"), $BoardId
    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
}

# -------------------------------
# Board guardrails
# -------------------------------

function Get-AllowedBoards {
    $allowedServiceEnv = $env:ConnectWisePsa_AllowedServiceBoards
    $allowedProjectEnv = $env:ConnectWisePsa_AllowedProjectBoards
    $allowedBoards = @()

    if (-not [string]::IsNullOrWhiteSpace($allowedServiceEnv)) {
        $allowedBoards += ($allowedServiceEnv -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    }

    if (-not [string]::IsNullOrWhiteSpace($allowedProjectEnv)) {
        $allowedBoards += ($allowedProjectEnv -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    }

    if ($allowedBoards.Count -eq 0) {
        $allowedBoards = @(
            'AutomateBoard',
            'Managed Service Alerts',
            'Managed Services Board',
            'Service Desk',
            'Network',
            'Security & Monitoring',
            'Test Board',
            'Engineering',
            'Engin',
            'Change Management',
            'CWRMM'
        )
    }

    return $allowedBoards
}

function Test-BoardAllowed {
    param(
        [string]$BoardName,
        [string[]]$AllowedBoards
    )

    if ([string]::IsNullOrWhiteSpace($BoardName)) { return $false }

    $boardNormalized = $BoardName.ToString().Trim().ToLower()
    foreach ($b in $AllowedBoards) {
        if ($b.ToString().Trim().ToLower() -eq $boardNormalized) { return $true }
    }

    return $false
}

function Get-ConnectWiseBoards {
    param(
        [string]$ConnectWiseUrl,
        [hashtable]$Headers
    )

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/boards?pageSize=1000" -f $ConnectWiseUrl.TrimEnd("/")
    $response = Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
    return @($response | ForEach-Object { $_ })
}

function Get-ConnectWisePriorities {
    param(
        [string]$ConnectWiseUrl,
        [hashtable]$Headers
    )

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/priorities?pageSize=1000" -f $ConnectWiseUrl.TrimEnd("/")
    $response = Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
    return @($response | ForEach-Object { $_ })
}

function Find-ConnectWiseCompanies {
    param(
        [string]$ConnectWiseUrl,
        [string]$CompanyName,
        [hashtable]$Headers
    )

    $safeName = ConvertTo-CwConditionString $CompanyName
    $conditions = ConvertTo-CwUrlEncoded ("name contains '{0}'" -f $safeName)
    $uri = "{0}/v4_6_release/apis/3.0/company/companies?conditions={1}&orderBy=name&pageSize=1000" -f $ConnectWiseUrl.TrimEnd("/"), $conditions
    $response = Invoke-ConnectWiseApi -Uri $uri -Method Get -Headers $Headers
    return @($response | ForEach-Object { $_ })
}

function Resolve-BoardReference {
    param(
        [string]$ConnectWiseUrl,
        [hashtable]$Headers,
        [object]$BoardId,
        [string]$BoardName,
        [string[]]$AllowedBoards
    )

    if (($null -eq $BoardId -or [string]::IsNullOrWhiteSpace([string]$BoardId)) -and [string]::IsNullOrWhiteSpace($BoardName)) {
        return $null
    }

    $allBoards = Get-ConnectWiseBoards -ConnectWiseUrl $ConnectWiseUrl -Headers $Headers
    $matched = $null

    if ($null -ne $BoardId -and (-not [string]::IsNullOrWhiteSpace([string]$BoardId))) {
        $parsedId = 0
        if ([int]::TryParse([string]$BoardId, [ref]$parsedId)) {
            $matched = @($allBoards) | Where-Object { [int]$_.id -eq $parsedId } | Select-Object -First 1
            if ($null -eq $matched) {
                try {
                    $singleBoard = Invoke-ConnectWiseApi -Uri ("{0}/v4_6_release/apis/3.0/service/boards/{1}" -f $ConnectWiseUrl.TrimEnd("/"), $parsedId) -Method Get -Headers $Headers
                    if ($singleBoard -and $singleBoard.id) {
                        $matched = $singleBoard
                    }
                }
                catch { }
            }
        }

        if ($null -eq $matched) {
            Invoke-BadRequest ("BoardId '{0}' was not found in ConnectWise." -f $BoardId)
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($BoardName)) {
        $trimmedName = $BoardName.Trim().ToLower()
        $matched = @($allBoards) | Where-Object { $_.name -and $_.name.ToString().Trim().ToLower() -eq $trimmedName } | Select-Object -First 1

        if ($null -eq $matched) {
            Invoke-BadRequest ("BoardName '{0}' was not found in ConnectWise." -f $BoardName)
        }
    }

    if ($null -ne $matched) {
        $isAllowed = Test-BoardAllowed -BoardName $matched.name -AllowedBoards $AllowedBoards
        return [PSCustomObject]@{
            Id        = [int]$matched.id
            Name      = [string]$matched.name
            IsAllowed = $isAllowed
        }
    }

    return $null
}

function Resolve-StatusReference {
    param(
        [string]$ConnectWiseUrl,
        [hashtable]$Headers,
        [int]$BoardId,
        [object]$StatusId,
        [string]$StatusName
    )

    if (($null -eq $StatusId -or [string]::IsNullOrWhiteSpace([string]$StatusId)) -and [string]::IsNullOrWhiteSpace($StatusName)) {
        return $null
    }

    $boardStatuses = @(Get-ConnectWiseBoardStatuses -ConnectWiseUrl $ConnectWiseUrl -BoardId $BoardId -Headers $Headers | ForEach-Object { $_ })
    $matched = $null

    if ($null -ne $StatusId -and (-not [string]::IsNullOrWhiteSpace([string]$StatusId))) {
        $parsedId = 0
        if ([int]::TryParse([string]$StatusId, [ref]$parsedId)) {
            $matched = @($boardStatuses) | Where-Object { [int]$_.id -eq $parsedId } | Select-Object -First 1
        }

        if ($null -eq $matched) {
            Invoke-BadRequest ("StatusId '{0}' was not found on board ID {1}." -f $StatusId, $BoardId)
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($StatusName)) {
        $trimmedName = $StatusName.Trim().ToLower()
        $matched = @($boardStatuses) | Where-Object { $_.name -and $_.name.ToString().Trim().ToLower() -eq $trimmedName } | Select-Object -First 1

        if ($null -eq $matched) {
            Invoke-BadRequest ("StatusName '{0}' was not found on board ID {1}." -f $StatusName, $BoardId)
        }
    }

    return $matched
}

function Resolve-PriorityReference {
    param(
        [string]$ConnectWiseUrl,
        [hashtable]$Headers,
        [object]$PriorityId,
        [string]$PriorityName
    )

    if (($null -eq $PriorityId -or [string]::IsNullOrWhiteSpace([string]$PriorityId)) -and [string]::IsNullOrWhiteSpace($PriorityName)) {
        return $null
    }

    $priorities = @(Get-ConnectWisePriorities -ConnectWiseUrl $ConnectWiseUrl -Headers $Headers)
    $matched = $null

    if ($null -ne $PriorityId -and (-not [string]::IsNullOrWhiteSpace([string]$PriorityId))) {
        $parsedId = 0
        if ([int]::TryParse([string]$PriorityId, [ref]$parsedId)) {
            $matched = @($priorities) | Where-Object { [int]$_.id -eq $parsedId } | Select-Object -First 1
        }

        if ($null -eq $matched) {
            Invoke-BadRequest ("PriorityId '{0}' was not found in ConnectWise." -f $PriorityId)
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($PriorityName)) {
        $trimmedName = $PriorityName.Trim().ToLower()
        $matched = @($priorities) | Where-Object { $_.name -and $_.name.ToString().Trim().ToLower() -eq $trimmedName } | Select-Object -First 1

        if ($null -eq $matched) {
            Invoke-BadRequest ("PriorityName '{0}' was not found in ConnectWise." -f $PriorityName)
        }
    }

    return $matched
}

# -------------------------------
# Patch builder
# -------------------------------

function Add-PatchOperation {
    param(
        [System.Collections.ArrayList]$Operations,
        [string]$Path,
        [object]$Value
    )

    if ($null -eq $Value) { return }
    [void]$Operations.Add(@{
        op    = "replace"
        path  = $Path
        value = $Value
    })
}

function New-ReferenceObject {
    param([object]$Id)
    if ($null -eq $Id) { return $null }
    if ([string]::IsNullOrWhiteSpace([string]$Id)) { return $null }

    $parsedId = 0
    if (-not [int]::TryParse([string]$Id, [ref]$parsedId)) {
        Invoke-BadRequest ("Invalid numeric id value: '{0}'." -f $Id)
    }

    return @{ id = $parsedId }
}

function New-TicketPatchOperations {
    param([object]$Changes)

    if ($null -eq $Changes) { Invoke-BadRequest "Missing Changes object." }
    if ($Changes -is [System.Collections.IDictionary]) { $Changes = [PSCustomObject]$Changes }

    $ops = New-Object System.Collections.ArrayList

    if (Test-HasProperty -Object $Changes -PropertyName "Summary") {
        Add-PatchOperation -Operations $ops -Path "/summary" -Value ([string]$Changes.Summary)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "InitialDescription") {
        Add-PatchOperation -Operations $ops -Path "/initialDescription" -Value ([string]$Changes.InitialDescription)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "BoardId") {
        Add-PatchOperation -Operations $ops -Path "/board" -Value (New-ReferenceObject -Id $Changes.BoardId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "StatusId") {
        Add-PatchOperation -Operations $ops -Path "/status" -Value (New-ReferenceObject -Id $Changes.StatusId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "TypeId") {
        Add-PatchOperation -Operations $ops -Path "/type" -Value (New-ReferenceObject -Id $Changes.TypeId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "SubTypeId") {
        Add-PatchOperation -Operations $ops -Path "/subType" -Value (New-ReferenceObject -Id $Changes.SubTypeId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "ItemId") {
        Add-PatchOperation -Operations $ops -Path "/item" -Value (New-ReferenceObject -Id $Changes.ItemId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "PriorityId") {
        Add-PatchOperation -Operations $ops -Path "/priority" -Value (New-ReferenceObject -Id $Changes.PriorityId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "AgreementId") {
        Add-PatchOperation -Operations $ops -Path "/agreement" -Value (New-ReferenceObject -Id $Changes.AgreementId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "CompanyId") {
        Add-PatchOperation -Operations $ops -Path "/company" -Value (New-ReferenceObject -Id $Changes.CompanyId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "ContactId") {
        Add-PatchOperation -Operations $ops -Path "/contact" -Value (New-ReferenceObject -Id $Changes.ContactId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "OwnerId") {
        Add-PatchOperation -Operations $ops -Path "/owner" -Value (New-ReferenceObject -Id $Changes.OwnerId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "TeamId") {
        Add-PatchOperation -Operations $ops -Path "/team" -Value (New-ReferenceObject -Id $Changes.TeamId)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "Severity") {
        Add-PatchOperation -Operations $ops -Path "/severity" -Value ([string]$Changes.Severity)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "Impact") {
        Add-PatchOperation -Operations $ops -Path "/impact" -Value ([string]$Changes.Impact)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "RequiredDate") {
        Add-PatchOperation -Operations $ops -Path "/requiredDate" -Value ([string]$Changes.RequiredDate)
    }

    if (Test-HasProperty -Object $Changes -PropertyName "ClosedFlag") {
        Add-PatchOperation -Operations $ops -Path "/closedFlag" -Value (Get-BoolValue -Value $Changes.ClosedFlag -DefaultValue $false)
    }

    return @($ops.ToArray())
}

function ConvertTo-ConnectWiseTicketCreateBody {
    param(
        [object]$Ticket,
        [string]$ConnectWiseUrl,
        [hashtable]$Headers,
        [string[]]$AllowedBoards
    )

    if ($null -eq $Ticket) { Invoke-BadRequest "Missing Ticket object." }
    if ($Ticket -is [System.Collections.IDictionary]) { $Ticket = [PSCustomObject]$Ticket }

    $summary = Get-ValueOrNull $Ticket "Summary"
    if ([string]::IsNullOrWhiteSpace([string]$summary)) {
        Invoke-BadRequest "Ticket.Summary is required for CreateTicket."
    }

    # Resolve Board
    $boardRef = Resolve-BoardReference `
        -ConnectWiseUrl $ConnectWiseUrl `
        -Headers $Headers `
        -BoardId (Get-ValueOrNull $Ticket "BoardId") `
        -BoardName (Get-ValueOrNull $Ticket "BoardName") `
        -AllowedBoards $AllowedBoards

    if ($null -eq $boardRef) {
        Invoke-BadRequest "Ticket.BoardId or Ticket.BoardName is required for CreateTicket."
    }
    if (-not $boardRef.IsAllowed) {
        throw [System.Security.SecurityException]::new(("Target board '{0}' is not in the allowed boards list." -f $boardRef.Name))
    }

    # Resolve Company
    $companyId = Get-ValueOrNull $Ticket "CompanyId"
    $companyName = Get-ValueOrNull $Ticket "CompanyName"

    $resolvedCompanyId = $null
    if ($null -ne $companyId -and (-not [string]::IsNullOrWhiteSpace([string]$companyId))) {
        $parsedCid = 0
        if (-not [int]::TryParse([string]$companyId, [ref]$parsedCid) -or $parsedCid -lt 1) {
            Invoke-BadRequest "Ticket.CompanyId must be a positive integer."
        }
        $resolvedCompanyId = $parsedCid
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$companyName)) {
        $candidates = @(Find-ConnectWiseCompanies -ConnectWiseUrl $ConnectWiseUrl -CompanyName ([string]$companyName).Trim() -Headers $Headers)
        if ($candidates.Count -eq 0) {
            throw [System.Collections.Generic.KeyNotFoundException]::new(("No ConnectWise company matched the supplied CompanyName '{0}'." -f $companyName))
        }

        $exact = @($candidates | Where-Object { $_.name -and $_.name.ToString().Trim().Equals(([string]$companyName).Trim(), [System.StringComparison]::OrdinalIgnoreCase) })
        if ($exact.Count -eq 1) {
            $resolvedCompanyId = [int]$exact[0].id
        }
        elseif ($exact.Count -gt 1 -or $candidates.Count -gt 1) {
            throw [System.InvalidOperationException]::new(("Multiple ConnectWise companies matched CompanyName '{0}'. Provide the exact company name or CompanyId." -f $companyName))
        }
        else {
            $resolvedCompanyId = [int]$candidates[0].id
        }
    }
    else {
        Invoke-BadRequest "Ticket.CompanyId or Ticket.CompanyName is required for CreateTicket."
    }

    $body = @{
        summary = [string]$summary
        board   = @{ id = [int]$boardRef.Id }
        company = @{ id = [int]$resolvedCompanyId }
    }

    # Resolve Status if provided
    $statusRef = Resolve-StatusReference `
        -ConnectWiseUrl $ConnectWiseUrl `
        -Headers $Headers `
        -BoardId ([int]$boardRef.Id) `
        -StatusId (Get-ValueOrNull $Ticket "StatusId") `
        -StatusName (Get-ValueOrNull $Ticket "StatusName")
    if ($null -ne $statusRef) {
        $body.status = @{ id = [int]$statusRef.id }
    }

    # Resolve Priority if provided
    $priorityRef = Resolve-PriorityReference `
        -ConnectWiseUrl $ConnectWiseUrl `
        -Headers $Headers `
        -PriorityId (Get-ValueOrNull $Ticket "PriorityId") `
        -PriorityName (Get-ValueOrNull $Ticket "PriorityName")
    if ($null -ne $priorityRef) {
        $body.priority = @{ id = [int]$priorityRef.id }
    }

    if (Test-HasProperty -Object $Ticket -PropertyName "InitialDescription") { $body.initialDescription = [string]$Ticket.InitialDescription }
    if (Test-HasProperty -Object $Ticket -PropertyName "ContactId") { $body.contact = New-ReferenceObject -Id $Ticket.ContactId }
    if (Test-HasProperty -Object $Ticket -PropertyName "TypeId") { $body.type = New-ReferenceObject -Id $Ticket.TypeId }
    if (Test-HasProperty -Object $Ticket -PropertyName "SubTypeId") { $body.subType = New-ReferenceObject -Id $Ticket.SubTypeId }
    if (Test-HasProperty -Object $Ticket -PropertyName "ItemId") { $body.item = New-ReferenceObject -Id $Ticket.ItemId }
    if (Test-HasProperty -Object $Ticket -PropertyName "AgreementId") { $body.agreement = New-ReferenceObject -Id $Ticket.AgreementId }
    if (Test-HasProperty -Object $Ticket -PropertyName "OwnerId") { $body.owner = New-ReferenceObject -Id $Ticket.OwnerId }
    if (Test-HasProperty -Object $Ticket -PropertyName "TeamId") { $body.team = New-ReferenceObject -Id $Ticket.TeamId }
    if (Test-HasProperty -Object $Ticket -PropertyName "Severity") { $body.severity = [string]$Ticket.Severity }
    if (Test-HasProperty -Object $Ticket -PropertyName "Impact") { $body.impact = [string]$Ticket.Impact }
    if (Test-HasProperty -Object $Ticket -PropertyName "RequiredDate") { $body.requiredDate = [string]$Ticket.RequiredDate }

    return [PSCustomObject]@{
        Body      = $body
        BoardName = $boardRef.Name
    }
}

# -------------------------------
# Audit helpers
# -------------------------------

function New-AuditNoteText {
    param(
        [string]$Action,
        [string]$TicketId,
        [string]$ApprovedBy,
        [string]$ApprovalReason,
        [object]$RequestedChanges,
        [string]$AdditionalNote
    )

    $changeJson = $null
    try { $changeJson = $RequestedChanges | ConvertTo-Json -Depth 20 }
    catch { $changeJson = "Unable to serialize requested changes." }

    $parts = @()
    $parts += "CloudFirst Copilot approved ticket write action."
    $parts += ("Action: {0}" -f $Action)
    if ($TicketId) { $parts += ("TicketId: {0}" -f $TicketId) }
    $parts += ("ApprovedBy: {0}" -f $ApprovedBy)
    if (-not [string]::IsNullOrWhiteSpace($ApprovalReason)) { $parts += ("ApprovalReason: {0}" -f $ApprovalReason) }
    $parts += ("CorrelationId: {0}" -f $CorrelationId)
    $parts += ("TimestampUtc: {0}" -f ((Get-Date).ToUniversalTime().ToString("o")))
    $parts += "RequestedChanges:"
    $parts += $changeJson
    if (-not [string]::IsNullOrWhiteSpace($AdditionalNote)) {
        $parts += "AdditionalNote:"
        $parts += $AdditionalNote
    }

    return ($parts -join "`n")
}

# -------------------------------
# Main
# -------------------------------

try {
    Write-Info "Starting ConnectWise ticket write function."

    $body = Get-RequestBodyObject -Request $Request
    if ($body -is [System.Collections.IDictionary]) {
        $body = [PSCustomObject]$body
    }

    $Action = Get-InputValue -Request $Request -Body $body -Names @("Action", "action")
    if ([string]::IsNullOrWhiteSpace($Action)) { $Action = "UpdateTicket" }

    $TicketId = Get-InputValue -Request $Request -Body $body -Names @("TicketId", "ticketId", "id")
    $RequestSecurityKey = Get-InputValue -Request $Request -Body $body -Names @("SecurityKey", "securityKey")

    $DryRun = Get-BoolValue -Value (Get-InputValue -Request $Request -Body $body -Names @("DryRun", "dryRun", "WhatIf", "whatIf", "ValidateOnly", "validateOnly")) -DefaultValue $false

    $Approved = $false
    if ($body -and (Test-HasProperty -Object $body -PropertyName "Approved")) {
        $Approved = Get-BoolValue -Value $body.Approved -DefaultValue $false
    }

    $ApprovedBy = Get-InputValue -Request $Request -Body $body -Names @("ApprovedBy", "approvedBy")
    $ApprovalReason = Get-InputValue -Request $Request -Body $body -Names @("ApprovalReason", "approvalReason")

    $ConfiguredSecurityKey = $env:SecurityKey
    if (-not [string]::IsNullOrWhiteSpace($ConfiguredSecurityKey)) {
        $headerSecurityKey = if ($Request.Headers) { [string]$Request.Headers.SecurityKey } else { $null }
        if (-not [string]::Equals($ConfiguredSecurityKey, $RequestSecurityKey, [System.StringComparison]::Ordinal) -and
            -not [string]::Equals($ConfiguredSecurityKey, $headerSecurityKey, [System.StringComparison]::Ordinal)) {
            Send-JsonResponse -StatusCode ([HttpStatusCode]::Unauthorized) -Body @{
                success       = $false
                correlationId = $CorrelationId
                action        = $Action
                errorCode     = "InvalidSecurityKey"
                error         = "Invalid security key."
            }
            return
        }
    }

    $requireApproval = $true
    if ($env:ConnectWisePsa_RequireApproval -eq "0" -or $DryRun) { $requireApproval = $false }

    if ($requireApproval -and (-not $Approved -or [string]::IsNullOrWhiteSpace($ApprovedBy))) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = $Action
            errorCode     = "ApprovalRequired"
            error         = "Approval required. Send Approved=true and ApprovedBy before this function performs writes."
        }
        return
    }

    $requiredSettings = @(
        "ConnectWisePsa_ApiBaseUrl",
        "ConnectWisePsa_ApiCompanyId",
        "ConnectWisePsa_ApiPublicKey",
        "ConnectWisePsa_ApiPrivateKey",
        "ConnectWisePsa_ApiClientId"
    )

    $missingSettings = @()
    foreach ($setting in $requiredSettings) {
        $value = Get-Item -Path Env:$setting -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Value
        if ([string]::IsNullOrWhiteSpace([string]$value)) { $missingSettings += $setting }
    }

    if ($missingSettings.Count -gt 0) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::InternalServerError) -Body @{
            success         = $false
            correlationId   = $CorrelationId
            action          = $Action
            errorCode       = "MissingApplicationSettings"
            error           = "Missing required application settings."
            missingSettings = $missingSettings
        }
        return
    }

    $ConnectWiseUrl = $env:ConnectWisePsa_ApiBaseUrl.TrimEnd("/")
    $headers = New-ConnectWiseHeaders `
        -CompanyId $env:ConnectWisePsa_ApiCompanyId `
        -PublicKey $env:ConnectWisePsa_ApiPublicKey `
        -PrivateKey $env:ConnectWisePsa_ApiPrivateKey `
        -ClientId $env:ConnectWisePsa_ApiClientId

    $allowedBoards = Get-AllowedBoards
    $auditNoteResult = $null

    switch ($Action.ToLower()) {
        "createticket" {
            $ticketCreateResult = ConvertTo-ConnectWiseTicketCreateBody `
                -Ticket $body.Ticket `
                -ConnectWiseUrl $ConnectWiseUrl `
                -Headers $headers `
                -AllowedBoards $allowedBoards

            $ticketCreateBody = $ticketCreateResult.Body
            $targetBoardName = $ticketCreateResult.BoardName

            if ($DryRun) {
                Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
                    success         = $true
                    dryRun          = $true
                    correlationId   = $CorrelationId
                    action          = "CreateTicket"
                    approvedBy      = $ApprovedBy
                    boardName       = $targetBoardName
                    simulatedTicket = $ticketCreateBody
                    message         = "Dry run validation succeeded. All guardrails passed. No ticket was created."
                    retrievedAtUtc  = (Get-Date).ToUniversalTime().ToString("o")
                }
                return
            }

            Write-Info "Creating ConnectWise ticket."

            $createdTicket = New-ConnectWiseTicket `
                -ConnectWiseUrl $ConnectWiseUrl `
                -Headers $headers `
                -TicketBody $ticketCreateBody

            $newTicketId = $createdTicket.id

            $auditNoteText = New-AuditNoteText `
                -Action "CreateTicket" `
                -TicketId $newTicketId `
                -ApprovedBy $ApprovedBy `
                -ApprovalReason $ApprovalReason `
                -RequestedChanges $body.Ticket `
                -AdditionalNote $body.AuditNote

            try {
                $auditNoteResult = Add-ConnectWiseTicketNote `
                    -ConnectWiseUrl $ConnectWiseUrl `
                    -TicketId $newTicketId `
                    -Headers $headers `
                    -Text $auditNoteText `
                    -NoteType "InternalAnalysis"
            }
            catch {
                Write-ErrorLog ("Audit note failed after ticket creation: {0}" -f $_.Exception.Message)
            }

            Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
                success         = $true
                correlationId   = $CorrelationId
                action          = "CreateTicket"
                ticketId        = $newTicketId
                approvedBy      = $ApprovedBy
                boardName       = $targetBoardName
                createdTicket   = $createdTicket
                auditNoteAdded  = ($null -ne $auditNoteResult)
                retrievedAtUtc  = (Get-Date).ToUniversalTime().ToString("o")
            }
            return
        }

        "addnote" {
            if (-not $TicketId) { Invoke-BadRequest "Missing TicketId for AddNote." }
            if ([string]::IsNullOrWhiteSpace([string]$body.NoteText)) { Invoke-BadRequest "Missing NoteText for AddNote." }

            $ticketBeforeNote = Get-ConnectWiseTicket -ConnectWiseUrl $ConnectWiseUrl -TicketId $TicketId -Headers $headers
            $boardName = $ticketBeforeNote.board.name

            if (-not (Test-BoardAllowed -BoardName $boardName -AllowedBoards $allowedBoards)) {
                Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                    success       = $false
                    correlationId = $CorrelationId
                    action        = "AddNote"
                    ticketId      = $TicketId
                    boardName     = $boardName
                    errorCode     = "BoardNotAllowed"
                    error         = "Ticket board not eligible for this write action."
                    allowedBoards = $allowedBoards
                }
                return
            }

            $noteType = "InternalAnalysis"
            if (Test-HasProperty -Object $body -PropertyName "NoteType") { $noteType = [string]$body.NoteType }

            if ($DryRun) {
                Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
                    success        = $true
                    dryRun         = $true
                    correlationId  = $CorrelationId
                    action         = "AddNote"
                    ticketId       = $TicketId
                    boardName      = $boardName
                    approvedBy     = $ApprovedBy
                    noteType       = $noteType
                    simulatedNote  = [string]$body.NoteText
                    message        = "Dry run validation succeeded. All guardrails passed. No note was added."
                    retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
                }
                return
            }

            $note = Add-ConnectWiseTicketNote `
                -ConnectWiseUrl $ConnectWiseUrl `
                -TicketId $TicketId `
                -Headers $headers `
                -Text ([string]$body.NoteText) `
                -NoteType $noteType

            Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
                success        = $true
                correlationId  = $CorrelationId
                action         = "AddNote"
                ticketId       = $TicketId
                boardName      = $boardName
                approvedBy     = $ApprovedBy
                note           = $note
                retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
            }
            return
        }

        "updateticket" {
            if (-not $TicketId) { Invoke-BadRequest "Missing TicketId for UpdateTicket." }

            $changes = $body.Changes
            if ($null -eq $changes) { Invoke-BadRequest "Missing Changes object for UpdateTicket." }
            if ($changes -is [System.Collections.IDictionary]) { $changes = [PSCustomObject]$changes }

            $ticketBefore = Get-ConnectWiseTicket -ConnectWiseUrl $ConnectWiseUrl -TicketId $TicketId -Headers $headers
            $beforeBoardName = $ticketBefore.board.name
            $beforeBoardId = [int]$ticketBefore.board.id

            if (-not (Test-BoardAllowed -BoardName $beforeBoardName -AllowedBoards $allowedBoards)) {
                Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                    success       = $false
                    correlationId = $CorrelationId
                    action        = "UpdateTicket"
                    ticketId      = $TicketId
                    boardName     = $beforeBoardName
                    errorCode     = "BoardNotAllowed"
                    error         = "Ticket board not eligible for this write action."
                    allowedBoards = $allowedBoards
                }
                return
            }

            # Check and resolve destination board if BoardId or BoardName is supplied
            $targetBoardRef = Resolve-BoardReference `
                -ConnectWiseUrl $ConnectWiseUrl `
                -Headers $headers `
                -BoardId (Get-ValueOrNull $changes "BoardId") `
                -BoardName (Get-ValueOrNull $changes "BoardName") `
                -AllowedBoards $allowedBoards

            if ($null -ne $targetBoardRef) {
                if (-not $targetBoardRef.IsAllowed) {
                    Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                        success       = $false
                        correlationId = $CorrelationId
                        action        = "UpdateTicket"
                        ticketId      = $TicketId
                        boardId       = $targetBoardRef.Id
                        boardName     = $targetBoardRef.Name
                        errorCode     = "DestinationBoardNotAllowed"
                        error         = ("Destination board '{0}' is not in the allowed boards list." -f $targetBoardRef.Name)
                        allowedBoards = $allowedBoards
                    }
                    return
                }
                $changes.BoardId = $targetBoardRef.Id
            }

            $effectiveBoardId = if ($null -ne $targetBoardRef) { $targetBoardRef.Id } else { $beforeBoardId }

            # Resolve StatusName if provided
            $targetStatusRef = Resolve-StatusReference `
                -ConnectWiseUrl $ConnectWiseUrl `
                -Headers $headers `
                -BoardId $effectiveBoardId `
                -StatusId (Get-ValueOrNull $changes "StatusId") `
                -StatusName (Get-ValueOrNull $changes "StatusName")

            if ($null -ne $targetStatusRef) {
                $changes.StatusId = [int]$targetStatusRef.id
            }

            # Resolve PriorityName if provided
            $targetPriorityRef = Resolve-PriorityReference `
                -ConnectWiseUrl $ConnectWiseUrl `
                -Headers $headers `
                -PriorityId (Get-ValueOrNull $changes "PriorityId") `
                -PriorityName (Get-ValueOrNull $changes "PriorityName")

            if ($null -ne $targetPriorityRef) {
                $changes.PriorityId = [int]$targetPriorityRef.id
            }

            # Closure protection: unless explicitly allowed, block any change that would close the ticket
            if ($env:ConnectWisePsa_AllowCloseTickets -ne "1") {
                if (Test-HasProperty -Object $changes -PropertyName "ClosedFlag") {
                    $requestedClosedFlag = Get-BoolValue -Value (Get-ValueOrNull -Object $changes -PropertyName "ClosedFlag") -DefaultValue $false
                    if ($requestedClosedFlag) {
                        Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                            success       = $false
                            correlationId = $CorrelationId
                            action        = "UpdateTicket"
                            ticketId      = $TicketId
                            errorCode     = "ClosureBlocked"
                            error         = "Ticket closure via ClosedFlag is blocked. Set ConnectWisePsa_AllowCloseTickets=1 only if closure writes are intentionally allowed."
                        }
                        return
                    }
                }

                if (Test-HasProperty -Object $changes -PropertyName "StatusId") {
                    $requestedStatusId = 0
                    if ([int]::TryParse([string](Get-ValueOrNull -Object $changes -PropertyName "StatusId"), [ref]$requestedStatusId)) {
                        $boardStatuses = @(Get-ConnectWiseBoardStatuses -ConnectWiseUrl $ConnectWiseUrl -BoardId $effectiveBoardId -Headers $headers | ForEach-Object { $_ })
                        $closingStatus = @($boardStatuses) | Where-Object { $_.id -eq $requestedStatusId -and $_.closedFlag -eq $true } | Select-Object -First 1

                        if ($closingStatus) {
                            Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                                success       = $false
                                correlationId = $CorrelationId
                                action        = "UpdateTicket"
                                ticketId      = $TicketId
                                statusId      = $requestedStatusId
                                statusName    = $closingStatus.name
                                errorCode     = "ClosureBlocked"
                                error         = "StatusId resolves to a closing status and ticket closure is blocked. Set ConnectWisePsa_AllowCloseTickets=1 only if closure writes are intentionally allowed."
                            }
                            return
                        }
                    }
                }
            }

            $patchOperations = New-TicketPatchOperations -Changes $changes

            $maxPatchOps = 25
            if ($env:ConnectWisePsa_MaxPatchOperations) {
                $parsedMax = 0
                if ([int]::TryParse([string]$env:ConnectWisePsa_MaxPatchOperations, [ref]$parsedMax) -and $parsedMax -gt 0) {
                    $maxPatchOps = $parsedMax
                }
            }

            if ($patchOperations.Count -lt 1) { Invoke-BadRequest "No supported update fields were supplied in Changes." }
            if ($patchOperations.Count -gt $maxPatchOps) { Invoke-BadRequest ("Too many patch operations. Max allowed: {0}" -f $maxPatchOps) }

            if ($DryRun) {
                Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
                    success         = $true
                    dryRun          = $true
                    correlationId   = $CorrelationId
                    action          = "UpdateTicket"
                    ticketId        = $TicketId
                    approvedBy      = $ApprovedBy
                    boardName       = $beforeBoardName
                    patchOperations = $patchOperations
                    before          = @{
                        summary       = $ticketBefore.summary
                        boardName     = $ticketBefore.board.name
                        statusName    = $ticketBefore.status.name
                        typeName      = $ticketBefore.type.name
                        subTypeName   = $ticketBefore.subType.name
                        itemName      = $ticketBefore.item.name
                        priorityName  = $ticketBefore.priority.name
                        agreementName = $ticketBefore.agreement.name
                        owner         = $ticketBefore.owner
                        teamName      = $ticketBefore.team.name
                    }
                    message         = "Dry run validation succeeded. All guardrails passed. No changes were applied."
                    retrievedAtUtc  = (Get-Date).ToUniversalTime().ToString("o")
                }
                return
            }

            Write-Info ("Updating ConnectWise ticket {0} with {1} patch operation(s)." -f $TicketId, $patchOperations.Count)

            $updateResult = Update-ConnectWiseTicket `
                -ConnectWiseUrl $ConnectWiseUrl `
                -TicketId $TicketId `
                -Headers $headers `
                -PatchOperations $patchOperations

            $ticketAfter = Get-ConnectWiseTicket -ConnectWiseUrl $ConnectWiseUrl -TicketId $TicketId -Headers $headers

            $auditNoteText = New-AuditNoteText `
                -Action "UpdateTicket" `
                -TicketId $TicketId `
                -ApprovedBy $ApprovedBy `
                -ApprovalReason $ApprovalReason `
                -RequestedChanges $changes `
                -AdditionalNote $body.AuditNote

            try {
                $auditNoteResult = Add-ConnectWiseTicketNote `
                    -ConnectWiseUrl $ConnectWiseUrl `
                    -TicketId $TicketId `
                    -Headers $headers `
                    -Text $auditNoteText `
                    -NoteType "InternalAnalysis"
            }
            catch {
                Write-ErrorLog ("Audit note failed after ticket update: {0}" -f $_.Exception.Message)
            }

            Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
                success          = $true
                correlationId    = $CorrelationId
                action           = "UpdateTicket"
                ticketId         = $TicketId
                approvedBy       = $ApprovedBy
                boardName        = $beforeBoardName
                patchOperations  = $patchOperations
                updateResult     = $updateResult
                auditNoteAdded   = ($null -ne $auditNoteResult)
                before           = @{
                    summary       = $ticketBefore.summary
                    boardName     = $ticketBefore.board.name
                    statusName    = $ticketBefore.status.name
                    typeName      = $ticketBefore.type.name
                    subTypeName   = $ticketBefore.subType.name
                    itemName      = $ticketBefore.item.name
                    priorityName  = $ticketBefore.priority.name
                    agreementName = $ticketBefore.agreement.name
                    owner         = $ticketBefore.owner
                    teamName      = $ticketBefore.team.name
                }
                after            = @{
                    summary       = $ticketAfter.summary
                    boardName     = $ticketAfter.board.name
                    statusName    = $ticketAfter.status.name
                    typeName      = $ticketAfter.type.name
                    subTypeName   = $ticketAfter.subType.name
                    itemName      = $ticketAfter.item.name
                    priorityName  = $ticketAfter.priority.name
                    agreementName = $ticketAfter.agreement.name
                    owner         = $ticketAfter.owner
                    teamName      = $ticketAfter.team.name
                }
                retrievedAtUtc   = (Get-Date).ToUniversalTime().ToString("o")
            }
            return
        }

        default {
            Send-JsonResponse -StatusCode ([HttpStatusCode]::BadRequest) -Body @{
                success       = $false
                correlationId = $CorrelationId
                action        = $Action
                errorCode     = "UnknownAction"
                error         = "Unknown Action. Supported actions: UpdateTicket, AddNote, CreateTicket."
            }
            return
        }
    }
}
catch {
    $errorMessage = $_.Exception.Message
    Write-ErrorLog $errorMessage

    $responseStatusCode = $null
    if ($_.Exception.Response) {
        try { $responseStatusCode = [int]$_.Exception.Response.StatusCode } catch { $responseStatusCode = $null }
    }

    if ($_.Exception -is [System.Security.SecurityException]) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = $Action
            errorCode     = "BoardNotAllowed"
            error         = $errorMessage
            allowedBoards = $allowedBoards
        }
        return
    }

    if ($responseStatusCode -eq 404 -or $_.Exception -is [System.Collections.Generic.KeyNotFoundException]) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::NotFound) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = $Action
            errorCode     = "ResourceNotFound"
            error         = "Resource not found in ConnectWise. $errorMessage"
        }
        return
    }

    if ($_.Exception -is [System.InvalidOperationException]) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::Conflict) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = $Action
            errorCode     = "AmbiguousReference"
            error         = $errorMessage
        }
        return
    }

    if ($_.Exception -is [System.ArgumentException]) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::BadRequest) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = $Action
            errorCode     = "InvalidRequestArgument"
            error         = $errorMessage
        }
        return
    }

    Send-JsonResponse -StatusCode ([HttpStatusCode]::InternalServerError) -Body @{
        success       = $false
        correlationId = $CorrelationId
        action        = $Action
        errorCode     = "UnhandledFunctionError"
        error         = "An unexpected error occurred while processing the ticket write action."
    }
}
