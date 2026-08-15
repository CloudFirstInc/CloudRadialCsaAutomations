<#
.SYNOPSIS
    Creates or updates ConnectWise Manage tickets for the CloudFirst Copilot Dispatch Agent.

.DESCRIPTION
    This Azure Function is the write-side companion to the CloudFirst read-only ticket context function.
    It updates approved ConnectWise Manage service-ticket fields, creates tickets when explicitly requested,
    and adds audit notes. It is designed for Copilot Studio actions where AI recommends changes and a
    human dispatcher/manager approves before execution.

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
    ConnectWisePsa_AllowCloseTickets     - Default 0. Set to 1 to allow Closed/Resolved/Cancelled status changes.
    ConnectWisePsa_AllowedServiceBoards  - Comma-separated board names allowed for updates
    ConnectWisePsa_AllowedProjectBoards  - Comma-separated project board names allowed for updates
    ConnectWisePsa_MaxPatchOperations    - Default 25

.INPUTS
    JSON body examples:

    Update ticket:
    {
        "Action": "UpdateTicket",
        "TicketId": 123456,
        "SecurityKey": "optional",
        "Approved": true,
        "ApprovedBy": "Dispatcher Name",
        "ApprovalReason": "Dispatcher accepted Copilot recommendation",
        "Changes": {
            "Summary": "Updated summary",
            "StatusId": 12,
            "TypeId": 34,
            "SubTypeId": 56,
            "ItemId": 78,
            "PriorityId": 2,
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
        "NoteType": "InternalAnalysis"
    }

    Create ticket:
    {
        "Action": "CreateTicket",
        "Approved": true,
        "ApprovedBy": "Dispatcher Name",
        "Ticket": {
            "Summary": "New service request",
            "BoardId": 1,
            "CompanyId": 123,
            "ContactId": 456,
            "StatusId": 10,
            "TypeId": 11,
            "SubTypeId": 12,
            "ItemId": 13,
            "PriorityId": 2,
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
    return ($null -ne $Object.PSObject.Properties[$PropertyName])
}

function Throw-BadRequest {
    param([string]$Message)

    throw [System.ArgumentException]::new($Message)
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
                return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body $bodyJson
            }
            else {
                return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers
            }
        }
        catch {
            $statusCode = $null
            $retryAfter = $null

            if ($_.Exception.Response) {
                try {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                    $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                }
                catch {
                    $statusCode = $null
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
        catch { Throw-BadRequest "Request body is not valid JSON." }
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
    return @{ id = [int]$Id }
}

function New-TicketPatchOperations {
    param([object]$Changes)

    if ($null -eq $Changes) { Throw-BadRequest "Missing Changes object." }

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

    return @($ops.ToArray())
}

function ConvertTo-ConnectWiseTicketCreateBody {
    param([object]$Ticket)

    if ($null -eq $Ticket) { Throw-BadRequest "Missing Ticket object." }

    if (-not (Test-HasProperty -Object $Ticket -PropertyName "Summary") -or [string]::IsNullOrWhiteSpace([string]$Ticket.Summary)) {
        Throw-BadRequest "Ticket.Summary is required for CreateTicket."
    }

    if (-not (Test-HasProperty -Object $Ticket -PropertyName "BoardId")) {
        Throw-BadRequest "Ticket.BoardId is required for CreateTicket."
    }

    if (-not (Test-HasProperty -Object $Ticket -PropertyName "CompanyId")) {
        Throw-BadRequest "Ticket.CompanyId is required for CreateTicket."
    }

    $body = @{
        summary = [string]$Ticket.Summary
        board   = New-ReferenceObject -Id $Ticket.BoardId
        company = New-ReferenceObject -Id $Ticket.CompanyId
    }

    if (Test-HasProperty -Object $Ticket -PropertyName "InitialDescription") { $body.initialDescription = [string]$Ticket.InitialDescription }
    if (Test-HasProperty -Object $Ticket -PropertyName "StatusId") { $body.status = New-ReferenceObject -Id $Ticket.StatusId }
    if (Test-HasProperty -Object $Ticket -PropertyName "ContactId") { $body.contact = New-ReferenceObject -Id $Ticket.ContactId }
    if (Test-HasProperty -Object $Ticket -PropertyName "TypeId") { $body.type = New-ReferenceObject -Id $Ticket.TypeId }
    if (Test-HasProperty -Object $Ticket -PropertyName "SubTypeId") { $body.subType = New-ReferenceObject -Id $Ticket.SubTypeId }
    if (Test-HasProperty -Object $Ticket -PropertyName "ItemId") { $body.item = New-ReferenceObject -Id $Ticket.ItemId }
    if (Test-HasProperty -Object $Ticket -PropertyName "PriorityId") { $body.priority = New-ReferenceObject -Id $Ticket.PriorityId }
    if (Test-HasProperty -Object $Ticket -PropertyName "AgreementId") { $body.agreement = New-ReferenceObject -Id $Ticket.AgreementId }
    if (Test-HasProperty -Object $Ticket -PropertyName "OwnerId") { $body.owner = New-ReferenceObject -Id $Ticket.OwnerId }
    if (Test-HasProperty -Object $Ticket -PropertyName "TeamId") { $body.team = New-ReferenceObject -Id $Ticket.TeamId }
    if (Test-HasProperty -Object $Ticket -PropertyName "Severity") { $body.severity = [string]$Ticket.Severity }
    if (Test-HasProperty -Object $Ticket -PropertyName "Impact") { $body.impact = [string]$Ticket.Impact }
    if (Test-HasProperty -Object $Ticket -PropertyName "RequiredDate") { $body.requiredDate = [string]$Ticket.RequiredDate }

    return $body
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

    $Action = Get-InputValue -Request $Request -Body $body -Names @("Action", "action")
    if ([string]::IsNullOrWhiteSpace($Action)) { $Action = "UpdateTicket" }

    $TicketId = Get-InputValue -Request $Request -Body $body -Names @("TicketId", "ticketId", "id")
    $RequestSecurityKey = Get-InputValue -Request $Request -Body $body -Names @("SecurityKey", "securityKey")

    $Approved = $false
    if ($body -and (Test-HasProperty -Object $body -PropertyName "Approved")) {
        $Approved = Get-BoolValue -Value $body.Approved -DefaultValue $false
    }

    $ApprovedBy = Get-InputValue -Request $Request -Body $body -Names @("ApprovedBy", "approvedBy")
    $ApprovalReason = Get-InputValue -Request $Request -Body $body -Names @("ApprovalReason", "approvalReason")

    $requireApproval = $true
    if ($env:ConnectWisePsa_RequireApproval -eq "0") { $requireApproval = $false }

    if ($requireApproval -and (-not $Approved -or [string]::IsNullOrWhiteSpace($ApprovedBy))) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = $Action
            error         = "Approval required. Send Approved=true and ApprovedBy before this function performs writes."
        }
        return
    }

    $SecurityKey = $env:SecurityKey
    if ($SecurityKey -and $SecurityKey -ne $RequestSecurityKey -and $SecurityKey -ne $Request.Headers.SecurityKey) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::Unauthorized) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = $Action
            error         = "Invalid security key."
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
            $ticketCreateBody = ConvertTo-ConnectWiseTicketCreateBody -Ticket $body.Ticket
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
                createdTicket   = $createdTicket
                auditNoteAdded  = ($null -ne $auditNoteResult)
                retrievedAtUtc  = (Get-Date).ToUniversalTime().ToString("o")
            }
            return
        }

        "addnote" {
            if (-not $TicketId) { Throw-BadRequest "Missing TicketId for AddNote." }
            if ([string]::IsNullOrWhiteSpace([string]$body.NoteText)) { Throw-BadRequest "Missing NoteText for AddNote." }

            $ticketBeforeNote = Get-ConnectWiseTicket -ConnectWiseUrl $ConnectWiseUrl -TicketId $TicketId -Headers $headers
            $boardName = $ticketBeforeNote.board.name

            if (-not (Test-BoardAllowed -BoardName $boardName -AllowedBoards $allowedBoards)) {
                Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                    success       = $false
                    correlationId = $CorrelationId
                    action        = "AddNote"
                    ticketId      = $TicketId
                    boardName     = $boardName
                    error         = "Ticket board not eligible for this write action."
                    allowedBoards = $allowedBoards
                }
                return
            }

            $noteType = "InternalAnalysis"
            if (Test-HasProperty -Object $body -PropertyName "NoteType") { $noteType = [string]$body.NoteType }

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

        default {
            if (-not $TicketId) { Throw-BadRequest "Missing TicketId for UpdateTicket." }

            $changes = $body.Changes
            $patchOperations = New-TicketPatchOperations -Changes $changes

            $maxPatchOps = 25
            if ($env:ConnectWisePsa_MaxPatchOperations) { $maxPatchOps = [int]$env:ConnectWisePsa_MaxPatchOperations }

            if ($patchOperations.Count -lt 1) { Throw-BadRequest "No supported update fields were supplied in Changes." }
            if ($patchOperations.Count -gt $maxPatchOps) { throw ("Too many patch operations. Max allowed: {0}" -f $maxPatchOps) }

            $ticketBefore = Get-ConnectWiseTicket -ConnectWiseUrl $ConnectWiseUrl -TicketId $TicketId -Headers $headers
            $beforeBoardName = $ticketBefore.board.name

            if (-not (Test-BoardAllowed -BoardName $beforeBoardName -AllowedBoards $allowedBoards)) {
                Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                    success       = $false
                    correlationId = $CorrelationId
                    action        = "UpdateTicket"
                    ticketId      = $TicketId
                    boardName     = $beforeBoardName
                    error         = "Ticket board not eligible for this write action."
                    allowedBoards = $allowedBoards
                }
                return
            }

            # Closure protection. The function cannot know status names from IDs without a board-status lookup,
            # so this blocks explicit closed flags and requires a separate opt-in for status changes intended to close tickets.
            if ($env:ConnectWisePsa_AllowCloseTickets -ne "1") {
                if (Test-HasProperty -Object $changes -PropertyName "ClosedFlag") {
                    Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                        success       = $false
                        correlationId = $CorrelationId
                        action        = "UpdateTicket"
                        ticketId      = $TicketId
                        error         = "ClosedFlag updates are blocked. Set ConnectWisePsa_AllowCloseTickets=1 only if closure writes are intentionally allowed."
                    }
                    return
                }
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
    }
}
catch {
    $errorMessage = $_.Exception.Message
    Write-ErrorLog $errorMessage

    if ($_.Exception -is [System.ArgumentException]) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::BadRequest) -Body @{
            success       = $false
            correlationId = $CorrelationId
            error         = $errorMessage
        }
        return
    }

    Send-JsonResponse -StatusCode ([HttpStatusCode]::InternalServerError) -Body @{
        success       = $false
        correlationId = $CorrelationId
        error         = $errorMessage
    }
}
