<#
#dcg
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
    Ticket context (default action):
    {
        "TicketId": "123456",
        "SecurityKey": "optional",
        "IncludeNotes": true,
        "IncludeCompany": true,
        "IncludeContact": true,
        "IncludeAgreement": true,
        "MaxNotes": 100
    }

    List new tickets (on-demand pull):
    {
        "Action": "ListNewTickets",
        "SecurityKey": "optional",
        "BoardNames": ["Service Desk", "Managed Services Board"],   // optional; defaults to all allowed boards
        "BoardIds": [33, 34],                                        // optional alternative to BoardNames
        "SinceMinutes": 60,                                          // or SinceHours; default 60, max 1440
        "PageSize": 100,                                             // default 100, max 200
        "ExcludeClosed": true                                        // default true
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

    $value = $null
    $found = $false

    if ($Body -is [System.Collections.IDictionary]) {
        foreach ($key in $Body.Keys) {
            if ($key.ToString().Equals($Name, [System.StringComparison]::OrdinalIgnoreCase)) {
                $value = $Body[$key]
                $found = $true
                break
            }
        }
    }
    else {
        $prop = $Body.PSObject.Properties |
            Where-Object {
                $_.Name.Equals($Name, [System.StringComparison]::OrdinalIgnoreCase)
            } |
            Select-Object -First 1

        if ($prop) {
            $value = $prop.Value
            $found = $true
        }
    }

    if (-not $found -or $null -eq $value) {
        return $DefaultValue
    }

    if ($value -is [bool]) {
        return $value
    }

    $text = [string]$value
    if ($text -match "^(true|1|yes|y)$") { return $true }
    if ($text -match "^(false|0|no|n)$") { return $false }
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
# Board guardrails + ticket listing
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

function Get-ConnectWiseBoards {
    param(
        [string]$ConnectWiseUrl,
        [hashtable]$Headers
    )

    $apiUrl = "{0}/v4_6_release/apis/3.0/service/boards?pageSize=1000" -f $ConnectWiseUrl.TrimEnd("/")
    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
}

function Get-ConnectWiseNewTickets {
    param(
        [string]$ConnectWiseUrl,
        [hashtable]$Headers,
        [int[]]$BoardIds,
        [int]$SinceMinutes,
        [int]$PageSize,
        [bool]$ExcludeClosed
    )

    $since = (Get-Date).ToUniversalTime().AddMinutes(-1 * $SinceMinutes).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $boardCondition = ($BoardIds | ForEach-Object { "board/id=$($_)" }) -join " or "
    $conditions = "($boardCondition) and dateEntered>=[$since]"
    if ($ExcludeClosed) {
        $conditions += " and closedFlag=false"
    }

    $encodedConditions = ConvertTo-CwUrlEncoded -Value $conditions
    $apiUrl = "{0}/v4_6_release/apis/3.0/service/tickets?conditions={1}&orderBy=dateEntered desc&pageSize={2}" -f $ConnectWiseUrl.TrimEnd("/"), $encodedConditions, $PageSize

    return Invoke-ConnectWiseApi -Uri $apiUrl -Method Get -Headers $Headers
}

function ConvertTo-TicketListItem {
    param([object]$Ticket)

    return @{
        id           = $Ticket.id
        summary      = $Ticket.summary
        boardId      = $Ticket.board.id
        boardName    = $Ticket.board.name
        statusId     = $Ticket.status.id
        statusName   = $Ticket.status.name
        priorityId   = $Ticket.priority.id
        priorityName = $Ticket.priority.name
        typeName     = $Ticket.type.name
        companyId    = $Ticket.company.id
        companyName  = $Ticket.company.name
        contactName  = $Ticket.contactName
        dateEntered  = if ($Ticket.dateEntered) { $Ticket.dateEntered } elseif ($Ticket._info.dateEntered) { $Ticket._info.dateEntered } else { $null }
        lastUpdated  = $Ticket._info.lastUpdated
        closedFlag   = $Ticket.closedFlag
        owner        = $Ticket.owner.identifier
    }
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
            dateEntered     = if ($Ticket.dateEntered) { $Ticket.dateEntered } elseif ($Ticket._info.dateEntered) { $Ticket._info.dateEntered } else { $null }
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
    if ($body -is [System.Collections.IDictionary]) {
        # Azure PowerShell HTTP triggers commonly deliver the body as an OrderedHashtable.
        # Normalize to PSCustomObject so PSObject.Properties checks work uniformly downstream.
        $body = [PSCustomObject]$body
    }

    if ($null -ne $Request.Body) {
        Write-DebugLog ("BODY TYPE: {0}" -f $Request.Body.GetType().FullName)
        Write-DebugLog ("BODY CONTENT: {0}" -f ($Request.Body | ConvertTo-Json -Depth 10 -Compress))
    }
    else {
        Write-DebugLog "BODY: null"
    }

    Write-DebugLog ("QUERY: {0}" -f ($Request.Query | ConvertTo-Json -Depth 10 -Compress))

    # Resolve action directly from the parsed body first (handles IDictionary bodies explicitly),
    # then fall back to the generalized input reader (query/headers).
    $Action = $null

    if ($body -is [System.Collections.IDictionary]) {
        foreach ($key in $body.Keys) {
            if ($null -ne $key -and $key.ToString().Equals("Action", [System.StringComparison]::OrdinalIgnoreCase)) {
                $Action = [string]$body[$key]
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($Action)) {
        $Action = Get-InputValue -Request $Request -Body $body -Names @("Action", "action")
    }

    if ([string]::IsNullOrWhiteSpace($Action)) {
        $Action = "GetTicketContext"
    }

    $Action = $Action.Trim()

    Write-Info ("Resolved action: '{0}'" -f $Action)

    $TicketId = Get-InputValue -Request $Request -Body $body -Names @("TicketId", "ticketId", "id")
    $RequestSecurityKey = Get-InputValue -Request $Request -Body $body -Names @("SecurityKey", "securityKey")

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
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            $missingSettings += $setting
        }
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

    if ($Action.Equals("ListNewTickets", [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-Info "Executing ListNewTickets action."

        $sinceMinutes = 60
        $rawSinceMinutes = Get-InputValue -Request $Request -Body $body -Names @("SinceMinutes", "sinceMinutes")
        $rawSinceHours = Get-InputValue -Request $Request -Body $body -Names @("SinceHours", "sinceHours")

        $parsedMinutes = 0
        if ($rawSinceMinutes -and [int]::TryParse([string]$rawSinceMinutes, [ref]$parsedMinutes) -and $parsedMinutes -gt 0) {
            $sinceMinutes = $parsedMinutes
        }
        elseif ($rawSinceHours) {
            $parsedHours = 0.0
            if ([double]::TryParse([string]$rawSinceHours, [ref]$parsedHours) -and $parsedHours -gt 0) {
                $sinceMinutes = [int][Math]::Ceiling($parsedHours * 60)
            }
        }

        if ($sinceMinutes -gt 1440) { $sinceMinutes = 1440 }
        if ($sinceMinutes -lt 1) { $sinceMinutes = 1 }

        $pageSize = 100
        $rawPageSize = Get-InputValue -Request $Request -Body $body -Names @("PageSize", "pageSize")
        $parsedPageSize = 0
        if ($rawPageSize -and [int]::TryParse([string]$rawPageSize, [ref]$parsedPageSize) -and $parsedPageSize -gt 0) {
            $pageSize = $parsedPageSize
        }
        if ($pageSize -gt 200) { $pageSize = 200 }

        $excludeClosed = $true
        if ($body -and $body.PSObject.Properties["ExcludeClosed"]) {
            $excludeClosed = Get-BoolInput -Body $body -Name "ExcludeClosed" -DefaultValue $true
        }

        # Resolve target board ids: explicit BoardIds, else BoardNames, else all allowed boards.
        $targetBoardIds = @()
        $allBoards = $null

        if ($body -and $body.PSObject.Properties["BoardIds"] -and $body.BoardIds) {
            $targetBoardIds = @($body.BoardIds | ForEach-Object { [int]$_ })

            # Guardrail: explicit ids must map to allowed board names.
            $allBoards = Get-ConnectWiseBoards -ConnectWiseUrl $ConnectWiseUrl -Headers $headers
            $allowedNormalized = $allowedBoards | ForEach-Object { $_.ToString().Trim().ToLower() }
            foreach ($bid in $targetBoardIds) {
                $matched = @($allBoards) | Where-Object { $_.id -eq $bid } | Select-Object -First 1
                if (-not $matched -or ($allowedNormalized -notcontains $matched.name.ToString().Trim().ToLower())) {
                    Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                        success       = $false
                        correlationId = $CorrelationId
                        action        = "ListNewTickets"
                        boardId       = $bid
                        error         = "Requested board id is not in the allowed boards list."
                        allowedBoards = $allowedBoards
                    }
                    return
                }
            }
        }
        else {
            $requestedNames = @()
            if ($body -and $body.PSObject.Properties["BoardNames"] -and $body.BoardNames) {
                $requestedNames = @($body.BoardNames | ForEach-Object { [string]$_ })
            }
            elseif ($body -and $body.PSObject.Properties["BoardName"] -and $body.BoardName) {
                $requestedNames = @([string]$body.BoardName)
            }

            if ($requestedNames.Count -gt 0) {
                # Guardrail: requested names must be in the allowed list.
                $allowedNormalized = $allowedBoards | ForEach-Object { $_.ToString().Trim().ToLower() }
                foreach ($rn in $requestedNames) {
                    if ($allowedNormalized -notcontains $rn.Trim().ToLower()) {
                        Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
                            success       = $false
                            correlationId = $CorrelationId
                            action        = "ListNewTickets"
                            boardName     = $rn
                            error         = "Requested board is not in the allowed boards list."
                            allowedBoards = $allowedBoards
                        }
                        return
                    }
                }
                $namesToResolve = $requestedNames
            }
            else {
                $namesToResolve = $allowedBoards
            }

            if ($null -eq $allBoards) {
                $allBoards = Get-ConnectWiseBoards -ConnectWiseUrl $ConnectWiseUrl -Headers $headers
            }

            foreach ($name in $namesToResolve) {
                $matched = @($allBoards) | Where-Object { $_.name.ToString().Trim().ToLower() -eq $name.Trim().ToLower() } | Select-Object -First 1
                if ($matched) {
                    $targetBoardIds += [int]$matched.id
                }
                else {
                    Write-Info ("Allowed board name '{0}' did not match any ConnectWise board; skipping." -f $name)
                }
            }
        }

        $targetBoardIds = @($targetBoardIds | Select-Object -Unique)

        if ($targetBoardIds.Count -eq 0) {
            Send-JsonResponse -StatusCode ([HttpStatusCode]::BadRequest) -Body @{
                success       = $false
                correlationId = $CorrelationId
                action        = "ListNewTickets"
                error         = "No resolvable boards. Provide BoardIds/BoardNames, or configure allowed board names that exist in ConnectWise."
                allowedBoards = $allowedBoards
            }
            return
        }

        Write-Info ("Listing tickets entered in the last {0} minute(s) on board ids: {1}" -f $sinceMinutes, ($targetBoardIds -join ","))

        $rawTickets = Get-ConnectWiseNewTickets `
            -ConnectWiseUrl $ConnectWiseUrl `
            -Headers $headers `
            -BoardIds $targetBoardIds `
            -SinceMinutes $sinceMinutes `
            -PageSize $pageSize `
            -ExcludeClosed $excludeClosed

        $tickets = @()
        foreach ($t in @($rawTickets)) {
            $tickets += ConvertTo-TicketListItem -Ticket $t
        }

        Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
            success        = $true
            correlationId  = $CorrelationId
            action         = "ListNewTickets"
            sinceMinutes   = $sinceMinutes
            sinceUtc       = (Get-Date).ToUniversalTime().AddMinutes(-1 * $sinceMinutes).ToString("o")
            boardIds       = $targetBoardIds
            count          = $tickets.Count
            tickets        = $tickets
            retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        }
        return
    }

    # Default action: GetTicketContext (single ticket retrieval)

    $IncludeNotes = Get-BoolInput -Body $body -Name "IncludeNotes" -DefaultValue $true
    $IncludeCompany = Get-BoolInput -Body $body -Name "IncludeCompany" -DefaultValue $false
    $IncludeContact = Get-BoolInput -Body $body -Name "IncludeContact" -DefaultValue $false
    $IncludeAgreement = Get-BoolInput -Body $body -Name "IncludeAgreement" -DefaultValue $true

    $defaultNotesPageSize = 100
    $parsedDefaultNotes = 0
    if ($env:ConnectWisePsa_DefaultNotesPageSize -and
        [int]::TryParse([string]$env:ConnectWisePsa_DefaultNotesPageSize, [ref]$parsedDefaultNotes) -and
        $parsedDefaultNotes -gt 0) {
        $defaultNotesPageSize = $parsedDefaultNotes
    }

    $maxNotesLimit = 250
    $parsedMaxNotes = 0
    if ($env:ConnectWisePsa_MaxNotesLimit -and
        [int]::TryParse([string]$env:ConnectWisePsa_MaxNotesLimit, [ref]$parsedMaxNotes) -and
        $parsedMaxNotes -gt 0) {
        $maxNotesLimit = $parsedMaxNotes
    }

    $MaxNotes = $defaultNotesPageSize
    if ($body -and $body.PSObject.Properties["MaxNotes"]) {
        $parsedBodyMax = 0
        if ([int]::TryParse([string]$body.MaxNotes, [ref]$parsedBodyMax)) {
            $MaxNotes = $parsedBodyMax
        }
    }

    if ($MaxNotes -gt $maxNotesLimit) {
        $MaxNotes = $maxNotesLimit
    }

    if ($MaxNotes -lt 1) {
        $MaxNotes = 1
    }

    if (-not $TicketId) {
        Send-JsonResponse -StatusCode ([HttpStatusCode]::BadRequest) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = $Action
            errorCode     = "MissingTicketId"
            error         = "Missing TicketId. Provide TicketId in the JSON body, query string, or header."
        }
        return
    }

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

        # Return a genuine 403 so the agent sees an accurate authorization failure.
        Send-JsonResponse -StatusCode ([HttpStatusCode]::Forbidden) -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = "GetTicketContext"
            ticketId      = $TicketId
            boardName     = $boardName
            errorCode     = "BoardNotAllowed"
            error         = "The ticket is on a ConnectWise board that this dispatcher integration is not authorized to access."
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
        action        = "GetTicketContext"
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
        action        = $Action
        error         = $errorMessage
    }
}
