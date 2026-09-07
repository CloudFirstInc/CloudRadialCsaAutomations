<#
.SYNOPSIS
    Retrieves ConnectWise Manage ticket data for the CloudFirst Copilot Dispatch Agent.

.DESCRIPTION
    This Azure Function retrieves ticket details, type/subtype/item classification,
    agreement, company, contact, priority, status, resources, and notes from
    ConnectWise Manage.

    The response is normalized for use by a Copilot Studio dispatch agent that
    analyzes ticket categorization, priority, agreement selection, routing,
    escalation, completeness, and recommended next actions.

.REQUIRED ENVIRONMENT VARIABLES
    ConnectWisePsa_ApiBaseUrl      - Base URL of the ConnectWise Manage API
    ConnectWisePsa_ApiCompanyId    - ConnectWise company ID
    ConnectWisePsa_ApiPublicKey    - ConnectWise public key
    ConnectWisePsa_ApiPrivateKey   - ConnectWise private key
    ConnectWisePsa_ApiClientId     - ConnectWise client ID
    SecurityKey                    - Optional shared secret to secure this function

.OPTIONAL ENVIRONMENT VARIABLES
    ConnectWisePsa_DefaultNotesPageSize - Default number of notes to retrieve
    ConnectWisePsa_MaxNotesLimit        - Maximum allowed notes
    DebugLogging                        - Set to 1 for verbose diagnostic logs

.INPUTS
    JSON body:
    {
        "TicketId": "123456",
        "SecurityKey": "optional",
        "IncludeNotes": true,
        "IncludeCompany": true,
        "IncludeContact": true,
        "IncludeAgreement": true,
        "MaxNotes": 100
    }

.OUTPUTS
    JSON object containing raw ConnectWise records and normalized Copilot context.

.NOTES
    This function is read-only. It does not modify ConnectWise tickets.
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

    $json = $Body | ConvertTo-Json -Depth 50

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

    if ($null -eq $Object) {
        return $null
    }

    $prop = $Object.PSObject.Properties[$PropertyName]

    if ($null -eq $prop) {
        return $null
    }

    return $prop.Value
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
                $bodyJson = $Body | ConvertTo-Json -Depth 20
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
                $sleepSeconds = [Math]::Pow(2, $attempt) + 1
            }

            Write-Info ("Retrying in {0} seconds" -f $sleepSeconds)
            Start-Sleep -Seconds $sleepSeconds
        }
    }
}

# -------------------------------
# Utility: URL encode condition values
# -------------------------------

function ConvertTo-CwUrlEncoded {
    param([string]$Value)

    if ($null -eq $Value) {
        return $null
    }

    return [System.Web.HttpUtility]::UrlEncode($Value)
}

# -------------------------------
# Input parsing
# -------------------------------

function Get-RequestBodyObject {
    param($Request)

    if ($null -eq $Request.Body) {
        return $null
    }

    if ($Request.Body -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Request.Body)) {
            return $null
        }

        try {
            return $Request.Body | ConvertFrom-Json
        }
        catch {
            throw "Request body is not valid JSON."
        }
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
            # If body is a hashtable/ordered dictionary, access by key
            if ($Body -is [System.Collections.IDictionary]) {
                if ($Body.Contains($name)) {
                    $val = $Body[$name]
                    if (-not [string]::IsNullOrWhiteSpace([string]$val)) { return [string]$val }
                }

                # try case-insensitive key match
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

                # try case-insensitive property name
                $prop = $Body.PSObject.Properties | Where-Object { $_.Name -eq $name -or $_.Name.ToLower() -eq $name.ToLower() } | Select-Object -First 1
                if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) { return [string]$prop.Value }
            }
        }

        if ($Request.Query) {
            $queryValue = $Request.Query[$name]
            if (-not [string]::IsNullOrWhiteSpace([string]$queryValue)) {
                return [string]$queryValue
            }

            foreach ($key in $Request.Query.Keys) {
                if ($key -eq $name -or $key.ToString().ToLower() -eq $name.ToLower()) {
                    $qv = $Request.Query[$key]
                    if (-not [string]::IsNullOrWhiteSpace([string]$qv)) { return [string]$qv }
                }
            }
        }

        if ($Request.Headers) {
            $headerValue = $Request.Headers[$name]
            if (-not [string]::IsNullOrWhiteSpace([string]$headerValue)) {
                return [string]$headerValue
            }

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

function Get-BoolInput {
    param(
        $Body,
        [string]$Name,
        [bool]$DefaultValue
    )

    if ($null -eq $Body) {
        return $DefaultValue
    }

    $prop = $Body.PSObject.Properties[$Name]
    if ($null -eq $prop) {
        return $DefaultValue
    }

    if ($prop.Value -is [bool]) {
        return $prop.Value
    }

    $text = [string]$prop.Value

    if ($text -match "^(true|1|yes|y)$") {
        return $true
    }

    if ($text -match "^(false|0|no|n)$") {
        return $false
    }

    return $DefaultValue
}

# -------------------------------
# ConnectWise retrieval functions
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

function Get-ConnectWiseTicketNotes {
    param(
        [string]$ConnectWiseUrl,
        [string]$TicketId,
        [hashtable]$Headers,
        [int]$MaxNotes
    )

    $pageSize = $MaxNotes
    if ($pageSize -lt 1) {
        $pageSize = 100
    }

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/tickets/{1}/notes?pageSize={2}&orderBy=dateCreated desc" -f $ConnectWiseUrl.TrimEnd("/"), $TicketId, $pageSize

    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
}

function Get-ConnectWiseCompany {
    param(
        [string]$ConnectWiseUrl,
        [int]$CompanyId,
        [hashtable]$Headers
    )

    if (-not $CompanyId) {
        return $null
    }

    $apiUrl = "{0}/v4_6_release/apis/3.0/company/companies/{1}" -f $ConnectWiseUrl.TrimEnd("/"), $CompanyId
    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
}

function Get-ConnectWiseContact {
    param(
        [string]$ConnectWiseUrl,
        [int]$ContactId,
        [hashtable]$Headers
    )

    if (-not $ContactId) {
        return $null
    }

    $apiUrl = "{0}/v4_6_release/apis/3.0/company/contacts/{1}" -f $ConnectWiseUrl.TrimEnd("/"), $ContactId
    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
}

function Get-ConnectWiseAgreement {
    param(
        [string]$ConnectWiseUrl,
        [int]$AgreementId,
        [hashtable]$Headers
    )

    if (-not $AgreementId) {
        return $null
    }

    $apiUrl = "{0}/v4_6_release/apis/3.0/finance/agreements/{1}" -f $ConnectWiseUrl.TrimEnd("/"), $AgreementId
    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
}

# -------------------------------
# Normalization for Copilot Dispatch Agent
# -------------------------------

function ConvertTo-PlainNote {
    param([object]$Note)

    return @{
        id              = Get-ValueOrNull -Object $Note -PropertyName "id"
        text            = Get-ValueOrNull -Object $Note -PropertyName "text"
        detailDescription = Get-ValueOrNull -Object $Note -PropertyName "detailDescription"
        internalAnalysisFlag = Get-ValueOrNull -Object $Note -PropertyName "internalAnalysisFlag"
        resolutionFlag  = Get-ValueOrNull -Object $Note -PropertyName "resolutionFlag"
        issueFlag       = Get-ValueOrNull -Object $Note -PropertyName "issueFlag"
        discussionFlag  = Get-ValueOrNull -Object $Note -PropertyName "discussionFlag"
        member          = Get-ValueOrNull -Object $Note -PropertyName "member"
        dateCreated     = Get-ValueOrNull -Object $Note -PropertyName "dateCreated"
        createdBy       = Get-ValueOrNull -Object $Note -PropertyName "createdBy"
    }
}

function New-CopilotDispatchContext {
    param(
        [object]$Ticket,
        [object[]]$Notes,
        [object]$Company,
        [object]$Contact,
        [object]$Agreement
    )

    $ticketTextParts = @()

    if ($Ticket.summary) {
        $ticketTextParts += ("Summary: {0}" -f $Ticket.summary)
    }

    if ($Ticket.initialDescription) {
        $ticketTextParts += ("Initial Description: {0}" -f $Ticket.initialDescription)
    }

    if ($Ticket.board.name) {
        $ticketTextParts += ("Board: {0}" -f $Ticket.board.name)
    }

    if ($Ticket.type.name) {
        $ticketTextParts += ("Type: {0}" -f $Ticket.type.name)
    }

    if ($Ticket.subType.name) {
        $ticketTextParts += ("SubType: {0}" -f $Ticket.subType.name)
    }

    if ($Ticket.item.name) {
        $ticketTextParts += ("Item: {0}" -f $Ticket.item.name)
    }

    if ($Ticket.priority.name) {
        $ticketTextParts += ("Priority: {0}" -f $Ticket.priority.name)
    }

    if ($Ticket.status.name) {
        $ticketTextParts += ("Status: {0}" -f $Ticket.status.name)
    }

    if ($Ticket.company.name) {
        $ticketTextParts += ("Company: {0}" -f $Ticket.company.name)
    }

    if ($Ticket.contactName) {
        $ticketTextParts += ("Contact: {0}" -f $Ticket.contactName)
    }

    if ($Ticket.agreement.name) {
        $ticketTextParts += ("Agreement: {0}" -f $Ticket.agreement.name)
    }

    $noteText = @()
    foreach ($note in $Notes) {
            $text = $note.text
        if ([string]::IsNullOrWhiteSpace([string]$text)) {
            $text = $note.detailDescription
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$text)) {
            $noteText += ("Note [{0}] by {1}: {2}" -f $note.dateCreated, $note.createdBy, $text)
        }
    }

    $allContextText = @()
    $allContextText += $ticketTextParts
    $allContextText += $noteText

    $missingFields = @()

    if (-not $Ticket.type.name) {
        $missingFields += "Type"
    }

    if (-not $Ticket.subType.name) {
        $missingFields += "Subtype"
    }

    if (-not $Ticket.item.name) {
        $missingFields += "Item"
    }

    if (-not $Ticket.priority.name) {
        $missingFields += "Priority"
    }

    if (-not $Ticket.agreement.name) {
        $missingFields += "Agreement"
    }

    if (-not $Ticket.contactName) {
        $missingFields += "Contact"
    }

    if (-not $Ticket.initialDescription -and -not $Ticket.summary) {
        $missingFields += "Problem description"
    }

    return @{
        purpose = "Provide ConnectWise ticket context to the CloudFirst Copilot Dispatch Agent for categorization, prioritization, routing, agreement validation, escalation review, and next-step recommendations."

        ticketIdentity = @{
            id              = $Ticket.id
            ticketNumber    = $Ticket.id
            summary         = $Ticket.summary
            recordType      = "ConnectWise Manage Service Ticket"
            source          = $Ticket.source.name
            dateEntered     = $Ticket.dateEntered
            lastUpdated     = $Ticket._info.lastUpdated
        }

        clientContext = @{
            companyId       = $Ticket.company.id
            companyName     = $Ticket.company.name
            contactId       = $Ticket.contact.id
            contactName     = $Ticket.contactName
            contactEmail    = $Ticket.contactEmailAddress
            contactPhone    = $Ticket.contactPhoneNumber
        }

        serviceContext = @{
            boardId         = $Ticket.board.id
            boardName       = $Ticket.board.name
            statusId        = $Ticket.status.id
            statusName      = $Ticket.status.name
            typeId          = $Ticket.type.id
            typeName        = $Ticket.type.name
            subtypeId       = $Ticket.subType.id
            subtypeName     = $Ticket.subType.name
            itemId          = $Ticket.item.id
            itemName        = $Ticket.item.name
            priorityId      = $Ticket.priority.id
            priorityName    = $Ticket.priority.name
            severity        = $Ticket.severity
            impact          = $Ticket.impact
        }

        agreementContext = @{
            agreementId     = $Ticket.agreement.id
            agreementName   = $Ticket.agreement.name
            agreementType   = $Agreement.type.name
            agreementStatus = $Agreement.status
        }

        assignmentContext = @{
            owner           = $Ticket.owner
            team            = $Ticket.team.name
            resources       = $Ticket.resources
        }

        slaContext = @{
            respondedFlag   = $Ticket.respondedFlag
            respondedDate   = $Ticket.respondedDate
            resolvedFlag    = $Ticket.resolvedFlag
            resolvedDate    = $Ticket.resolvedDate
            closedFlag      = $Ticket.closedFlag
            closedDate      = $Ticket.closedDate
        }

        qualityChecks = @{
            missingFields   = $missingFields
            hasNotes        = ($Notes.Count -gt 0)
            noteCount       = $Notes.Count
            hasAgreement    = (-not [string]::IsNullOrWhiteSpace([string]$Ticket.agreement.name))
            hasCategory     = (-not [string]::IsNullOrWhiteSpace([string]$Ticket.type.name))
            hasSubcategory  = (-not [string]::IsNullOrWhiteSpace([string]$Ticket.subType.name))
            hasItem         = (-not [string]::IsNullOrWhiteSpace([string]$Ticket.item.name))
        }

        agentInstructions = @{
            analyzeFor = @(
                "Correct service board",
                "Correct ticket type",
                "Correct subtype",
                "Correct item",
                "Correct agreement",
                "Correct priority based on impact and urgency",
                "Whether more information is needed before assignment",
                "Whether ticket should be escalated",
                "Recommended technician or team",
                "Relevant knowledge base or SOP suggestions",
                "Recommended next dispatcher action"
            )
            doNotDo = @(
                "Do not modify the ticket",
                "Do not close the ticket",
                "Do not assign the ticket without human dispatcher approval",
                "Do not make final HR or performance judgments"
            )
        }

        analysisText = ($allContextText -join "`n")
    }
}

# -------------------------------
# Main
# -------------------------------

try {
    Write-Info "Starting ConnectWise ticket dispatch context retrieval."

    $body = Get-RequestBodyObject -Request $Request

    Write-Host "BODY TYPE:"
    Write-Host ($Request.Body.GetType().FullName)

    Write-Host "BODY CONTENT:"
    Write-Host ($Request.Body | ConvertTo-Json -Depth 10)

    Write-Host "QUERY:"
    Write-Host ($Request.Query | ConvertTo-Json -Depth 10)

    $TicketId = Get-InputValue -Request $Request -Body $body -Names @("TicketId", "ticketId", "id")
    $RequestSecurityKey = Get-InputValue -Request $Request -Body $body -Names @("SecurityKey", "securityKey")

    $IncludeNotes = Get-BoolInput -Body $body -Name "IncludeNotes" -DefaultValue $true
    $IncludeCompany = Get-BoolInput -Body $body -Name "IncludeCompany" -DefaultValue $false
    $IncludeContact = Get-BoolInput -Body $body -Name "IncludeContact" -DefaultValue $false
    $IncludeAgreement = Get-BoolInput -Body $body -Name "IncludeAgreement" -DefaultValue $true

    $defaultNotesPageSize = 100
    if ($env:ConnectWisePsa_DefaultNotesPageSize) {
        $defaultNotesPageSize = [int]$env:ConnectWisePsa_DefaultNotesPageSize
    }

    $maxNotesLimit = 250
    if ($env:ConnectWisePsa_MaxNotesLimit) {
        $maxNotesLimit = [int]$env:ConnectWisePsa_MaxNotesLimit
    }

    $MaxNotes = $defaultNotesPageSize
    if ($body -and $body.PSObject.Properties["MaxNotes"]) {
        $MaxNotes = [int]$body.MaxNotes
    }

    if ($MaxNotes -gt $maxNotesLimit) {
        $MaxNotes = $maxNotesLimit
    }

    if ($MaxNotes -lt 1) {
        $MaxNotes = 1
    }

    $SecurityKey = $env:SecurityKey

    if ($SecurityKey -and $SecurityKey -ne $RequestSecurityKey -and $SecurityKey -ne $Request.Headers.SecurityKey) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::Unauthorized) -Body @{
            success       = $false
            correlationId = $CorrelationId
            error         = "Invalid security key."
        }
        return
    }

    if (-not $TicketId) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::BadRequest) -Body @{
            success       = $false
            correlationId = $CorrelationId
            error         = "Missing TicketId. Provide TicketId in the JSON body, query string, or header."
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
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            $missingSettings += $setting
        }
    }

    if ($missingSettings.Count -gt 0) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::InternalServerError) -Body @{
            success         = $false
            correlationId   = $CorrelationId
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

    Write-Info ("Retrieving ticket {0}" -f $TicketId)

    $ticket = Get-ConnectWiseTicket `
        -ConnectWiseUrl $ConnectWiseUrl `
        -TicketId $TicketId `
        -Headers $headers

    # -------------------------------
    # Board filtering: allow only specific service or project boards
    # Configure via environment variables:
    #  - ConnectWisePsa_AllowedServiceBoards (comma-separated)
    #  - ConnectWisePsa_AllowedProjectBoards (comma-separated)
    # If neither is set, defaults to Service Desk, Projects, Engineering.
    $boardName = $null
    try { $boardName = $ticket.board.name } catch { $boardName = $null }

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
        $allowedServiceDefaults = @(
            'AutomateBoard',
            'Managed Service Alerts',
            'Managed Services Board',
            'Service Desk',
            'Network',
            'Security & Monitoring',
            'Test Board'
        )

        $allowedProjectDefaults = @(
            'Engineering',
            'Engin',
            'Change Management',
            'CWRMM'
        )

        $allowedBoards = $allowedServiceDefaults + $allowedProjectDefaults
    }

    $boardNormalized = ''
    if ($boardName) { $boardNormalized = $boardName.ToString().Trim().ToLower() }

    $isAllowed = $false
    foreach ($b in $allowedBoards) {
        if ($b.ToString().Trim().ToLower() -eq $boardNormalized) {
            $isAllowed = $true
            break
        }
    }

    if (-not $isAllowed) {
        Write-Info ("Ticket {0} is on board '{1}' which is not in the allowed boards list." -f $TicketId, $boardName)

        # Structured blocked-ticket log for metrics/alerts
        $blockedLog = @{
            event = 'BLOCKED_TICKET'
            correlationId = $CorrelationId
            ticketId = $TicketId
            boardName = $boardName
            allowedBoards = $allowedBoards
            timestampUtc = (Get-Date).ToUniversalTime().ToString('o')
        }

        try {
            Write-Info ("BLOCKED_TICKET: {0}" -f ($blockedLog | ConvertTo-Json -Depth 5))
        }
        catch {
            Write-Info ("BLOCKED_TICKET: ticketId={0} board={1}" -f $TicketId, $boardName)
        }

        # Optional webhook for metrics: set ConnectWisePsa_BlockedTicketWebhook to a URL
        if (-not [string]::IsNullOrWhiteSpace($env:ConnectWisePsa_BlockedTicketWebhook)) {
            try {
                $webhookUrl = $env:ConnectWisePsa_BlockedTicketWebhook
                $headersForWebhook = @{ 'Content-Type' = 'application/json' }
                $bodyJson = $blockedLog | ConvertTo-Json -Depth 10
                Invoke-RestMethod -Uri $webhookUrl -Method Post -Headers $headersForWebhook -Body $bodyJson -ErrorAction Stop
                Write-Info "Blocked-ticket webhook POST succeeded."
            }
            catch {
                Write-ErrorLog ("Failed to POST blocked-ticket webhook: {0}" -f $_.Exception.Message)
            }
        }

        # Return success response but include a connector-style authorization message
        $exampleConversationId = "a72ec4d4-2c98-4174-a0ca-dc36c978c42a"
        $exampleTimeUtc = "2026-08-15T05:16:04.583Z"
        $authMessage = "Error Message: The connector 'ConnectWiseDispatchAgent' returned an authorization error (HTTP 403). The connection may need to be reauthorized.`nError Code: ConnectorAuthorizationError`nConversation Id: {0}`nTime (UTC): {1}" -f $exampleConversationId, $exampleTimeUtc

        Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
            success       = $true
            correlationId = $CorrelationId
            ticketId      = $TicketId
            boardName     = $boardName
            message       = $authMessage
            errorCode     = "ConnectorAuthorizationError"
            conversationId = $exampleConversationId
            timeUtc       = $exampleTimeUtc
            allowedBoards = $allowedBoards
        }

        return
    }

    $notes = @()
    if ($IncludeNotes) {
        Write-Info ("Retrieving up to {0} notes for ticket {1}" -f $MaxNotes, $TicketId)

        $rawNotes = Get-ConnectWiseTicketNotes `
            -ConnectWiseUrl $ConnectWiseUrl `
            -TicketId $TicketId `
            -Headers $headers `
            -MaxNotes $MaxNotes

        foreach ($note in $rawNotes) {
            $notes += ConvertTo-PlainNote -Note $note
        }
    }

    $company = $null
    if ($IncludeCompany -and $ticket.company.id) {
        Write-Info ("Retrieving company {0}" -f $ticket.company.id)

        $company = Get-ConnectWiseCompany `
            -ConnectWiseUrl $ConnectWiseUrl `
            -CompanyId ([int]$ticket.company.id) `
            -Headers $headers
    }

    $contact = $null
    if ($IncludeContact -and $ticket.contact.id) {
        Write-Info ("Retrieving contact {0}" -f $ticket.contact.id)

        $contact = Get-ConnectWiseContact `
            -ConnectWiseUrl $ConnectWiseUrl `
            -ContactId ([int]$ticket.contact.id) `
            -Headers $headers
    }

    $agreement = $null
    if ($IncludeAgreement -and $ticket.agreement.id) {
        Write-Info ("Retrieving agreement {0}" -f $ticket.agreement.id)

        $agreement = Get-ConnectWiseAgreement `
            -ConnectWiseUrl $ConnectWiseUrl `
            -AgreementId ([int]$ticket.agreement.id) `
            -Headers $headers
    }

    $copilotContext = New-CopilotDispatchContext `
        -Ticket $ticket `
        -Notes $notes `
        -Company $company `
        -Contact $contact `
        -Agreement $agreement

    $response = @{
        success       = $true
        correlationId = $CorrelationId
        ticketId      = $TicketId
        retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")

        requestOptions = @{
            includeNotes     = $IncludeNotes
            includeCompany   = $IncludeCompany
            includeContact   = $IncludeContact
            includeAgreement = $IncludeAgreement
            maxNotes         = $MaxNotes
        }

        copilotContext = $copilotContext

        raw = @{
            ticket    = $ticket
            notes     = $notes
            company   = $company
            contact   = $contact
            agreement = $agreement
        }
    }

    Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body $response
}
catch {
    $errorMessage = $_.Exception.Message
    Write-ErrorLog $errorMessage

    Send-JsonResponse -StatusCode ([HttpStatusCode]::InternalServerError) -Body @{
        success       = $false
        correlationId = $CorrelationId
        error         = $errorMessage
    }
}
