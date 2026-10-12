<#
.SYNOPSIS
    Retrieves ConnectWise Manage service tickets for a Client company.

.DESCRIPTION
    This read-only Azure Function accepts a company name, resolves the matching
    ConnectWise company, validates the company type Client, and retrieves that
    company's service tickets.

    TicketMode controls the result set:
      Open   - tickets where closedFlag=false (default)
      Closed - tickets where closedFlag=true
      All    - all company service tickets

    Tickets are always restricted to the allowed boards (same rules as
    ConnectWise-GetTicketDispatchContext: ConnectWisePsa_AllowedServiceBoards and
    ConnectWisePsa_AllowedProjectBoards, comma-separated, with a built-in default
    list). Optional BoardName narrows results to one allowed board; a board outside
    the list returns 403 BoardNotAllowed. Results are paged,
    normalized for agent use, and also returned as complete raw ticket records.

.REQUIRED ENVIRONMENT VARIABLES
    ConnectWisePsa_ApiBaseUrl
    ConnectWisePsa_ApiCompanyId
    ConnectWisePsa_ApiPublicKey
    ConnectWisePsa_ApiPrivateKey
    ConnectWisePsa_ApiClientId
    SecurityKey                           - Optional shared secret

.OPTIONAL ENVIRONMENT VARIABLES
    ConnectWisePsa_TicketsPageSize        - Default 100; maximum 1000
    ConnectWisePsa_TicketsMaxPages        - Default 25; maximum 1000
    DebugLogging                          - Set to 1 for diagnostic logs

.INPUTS
    JSON body:
    {
        "CompanyName": "City of Rye",
        "TicketMode": "Open",
        "BoardName": "Service Desk",
        "SecurityKey": "optional"
    }

    Query-string example:
    ?CompanyName=City%20of%20Rye&TicketMode=Open

.OUTPUTS
    JSON containing the selected company, retrieval criteria, ticket counts,
    normalized ticket summaries, and complete raw ticket records.

.NOTES
    This function is read-only. The default is Open to keep agent context focused.
#>

using namespace System.Net
param($Request, $TriggerMetadata)

$CorrelationId = [guid]::NewGuid().ToString()
$DebugEnabled = ($env:DebugLogging -eq "1")

function Write-Info { param([string]$Message) Write-Host ("[{0}] [INFO] {1}" -f $CorrelationId,$Message) }
function Write-DebugLog { param([string]$Message) if ($DebugEnabled) { Write-Host ("[{0}] [DEBUG] {1}" -f $CorrelationId,$Message) } }
function Write-ErrorLog { param([string]$Message) Write-Host ("[{0}] [ERROR] {1}" -f $CorrelationId,$Message) }

function Send-JsonResponse {
    param([System.Net.HttpStatusCode]$StatusCode,[object]$Body)
    $json=$Body | ConvertTo-Json -Depth 100
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode=$StatusCode; Body=$json; ContentType="application/json"
    })
}

function Get-ValueOrNull {
    param([object]$Object,[string]$PropertyName)
    if ($null -eq $Object) { return $null }
    $property=$Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function New-ConnectWiseHeaders {
    param([string]$CompanyId,[string]$PublicKey,[string]$PrivateKey,[string]$ClientId)
    $authString="{0}+{1}:{2}" -f $CompanyId,$PublicKey,$PrivateKey
    $encodedAuth=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($authString))
    return @{
        Authorization="Basic {0}" -f $encodedAuth
        "Content-Type"="application/json"
        Accept="application/json"
        ClientId=$ClientId
    }
}

function Invoke-ConnectWiseApi {
    param([string]$Uri,[string]$Method="Get",[hashtable]$Headers,[object]$Body=$null,[int]$MaxRetries=5)
    $attempt=0
    while ($true) {
        try {
            $attempt++
            Write-DebugLog ("CW API {0} {1}, attempt {2}" -f $Method,$Uri,$attempt)
            if ($null -ne $Body) {
                return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body ($Body | ConvertTo-Json -Depth 20) -ErrorAction Stop
            }
            return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -ErrorAction Stop
        }
        catch {
            $statusCode=$null; $retryAfter=$null
            if ($_.Exception.Response) {
                try { $statusCode=[int]$_.Exception.Response.StatusCode } catch {}
                try { $retryAfter=$_.Exception.Response.Headers["Retry-After"] } catch {}
            }
            Write-ErrorLog ("CW API call failed. StatusCode={0}. Message={1}" -f $statusCode,$_.Exception.Message)
            if (($statusCode -notin @(408,429,500,502,503,504)) -or $attempt -ge $MaxRetries) { throw }
            $sleepSeconds=0; $parsedRetryAfter=0
            if ($retryAfter -and [int]::TryParse([string]$retryAfter,[ref]$parsedRetryAfter)) { $sleepSeconds=$parsedRetryAfter }
            if ($sleepSeconds -lt 1) { $sleepSeconds=[int][Math]::Pow(2,$attempt) + (Get-Random -Minimum 1 -Maximum 4) }
            Write-Info ("Retrying ConnectWise request in {0} seconds." -f $sleepSeconds)
            Start-Sleep -Seconds $sleepSeconds
        }
    }
}

function ConvertTo-CwUrlEncoded { param([string]$Value) if ($null -eq $Value) { return $null }; return [Uri]::EscapeDataString($Value) }
function ConvertTo-CwConditionString { param([string]$Value) if ($null -eq $Value) { return $null }; return $Value.Replace("'","''") }

function Get-RequestBodyObject {
    param($Request)
    if ($null -eq $Request.Body) { return $null }
    if ($Request.Body -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Request.Body)) { return $null }
        try { return $Request.Body | ConvertFrom-Json } catch { throw "Request body is not valid JSON." }
    }
    return $Request.Body
}

function Get-InputValue {
    param($Request,$Body,[string[]]$Names)
    foreach ($name in $Names) {
        if ($Body) {
            if ($Body -is [Collections.IDictionary]) {
                foreach ($key in $Body.Keys) {
                    if ($null -ne $key -and $key.ToString().Equals($name,[StringComparison]::OrdinalIgnoreCase)) {
                        $value=$Body[$key]
                        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
                    }
                }
            }
            else {
                $property=$Body.PSObject.Properties | Where-Object { $_.Name.Equals($name,[StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
                if ($property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return [string]$property.Value }
            }
        }
        foreach ($source in @($Request.Query,$Request.Headers)) {
            if ($source) {
                foreach ($key in $source.Keys) {
                    if ($null -ne $key -and $key.ToString().Equals($name,[StringComparison]::OrdinalIgnoreCase)) {
                        $value=$source[$key]
                        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
                    }
                }
            }
        }
    }
    return $null
}

function Find-ConnectWiseCompanies {
    param([string]$ConnectWiseUrl,[string]$CompanyName,[hashtable]$Headers)
    $safeName=ConvertTo-CwConditionString $CompanyName
    $conditions=ConvertTo-CwUrlEncoded ("name contains '{0}'" -f $safeName)
    $uri=("{0}/v4_6_release/apis/3.0/company/companies?conditions={1}&orderBy=name&pageSize=1000" -f $ConnectWiseUrl.TrimEnd('/'),$conditions)
    return @(Invoke-ConnectWiseApi -Uri $uri -Headers $Headers | ForEach-Object { $_ })
}

function Get-ConnectWiseCompany {
    param([string]$ConnectWiseUrl,[int]$CompanyId,[hashtable]$Headers)
    $uri=("{0}/v4_6_release/apis/3.0/company/companies/{1}" -f $ConnectWiseUrl.TrimEnd('/'),$CompanyId)
    return Invoke-ConnectWiseApi -Uri $uri -Headers $Headers
}

function Get-CompanyTypeNames {
    param([object]$Company)
    $names=@()
    foreach ($type in @(Get-ValueOrNull $Company "types")) {
        if ($null -eq $type) { continue }
        if ($type -is [string]) { $names += $type.Trim() }
        elseif ($type.PSObject.Properties["name"]) { $names += ([string]$type.name).Trim() }
    }
    $single=Get-ValueOrNull $Company "type"
    if ($single -is [string]) { $names += $single.Trim() }
    elseif ($single -and $single.PSObject.Properties["name"]) { $names += ([string]$single.name).Trim() }
    return @($names | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
}

function Test-CompanyIsClient {
    param([object]$Company)
    foreach ($name in @(Get-CompanyTypeNames $Company)) {
        if ($name.Equals("Client",[StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Get-NormalizedTicketMode {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "Open" }
    switch ($Value.Trim().ToLowerInvariant()) {
        "open"   { return "Open" }
        "closed" { return "Closed" }
        "all"    { return "All" }
        default  { throw "TicketMode must be Open, Closed, or All." }
    }
}

function Get-AllowedBoards {
    $allowedBoards=@()
    foreach ($value in @($env:ConnectWisePsa_AllowedServiceBoards,$env:ConnectWisePsa_AllowedProjectBoards)) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $allowedBoards += ($value -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        }
    }
    if ($allowedBoards.Count -eq 0) {
        $allowedBoards=@('AutomateBoard','Managed Service Alerts','Managed Services Board','Service Desk','Network','Security & Monitoring','Test Board','Engineering','Engin','Change Management','CWRMM')
    }
    return $allowedBoards
}

function Get-ConnectWiseBoards {
    param([string]$ConnectWiseUrl,[hashtable]$Headers)
    $uri="{0}/v4_6_release/apis/3.0/service/boards?pageSize=1000" -f $ConnectWiseUrl.TrimEnd('/')
    return @(Invoke-ConnectWiseApi -Uri $uri -Method Get -Headers $Headers | ForEach-Object { $_ })
}

function Get-ConnectWiseClientTickets {
    param(
        [string]$ConnectWiseUrl,
        [int]$CompanyId,
        [string]$TicketMode,
        [int[]]$BoardIds,
        [hashtable]$Headers,
        [int]$PageSize=100,
        [int]$MaxPages=25
    )
    if ($PageSize -lt 1) { $PageSize=100 }; if ($PageSize -gt 1000) { $PageSize=1000 }
    if ($MaxPages -lt 1) { $MaxPages=25 }; if ($MaxPages -gt 1000) { $MaxPages=1000 }

    $conditionParts=@("company/id=$CompanyId")
    if ($TicketMode -eq "Open") { $conditionParts += "closedFlag=false" }
    elseif ($TicketMode -eq "Closed") { $conditionParts += "closedFlag=true" }
    $conditionParts += "(" + (($BoardIds | ForEach-Object { "board/id=$_" }) -join " or ") + ")"
    $conditions=$conditionParts -join " and "
    $encodedConditions=ConvertTo-CwUrlEncoded $conditions
    $tickets=@(); $page=1; $pagesRetrieved=0; $truncated=$false

    while ($page -le $MaxPages) {
        $uri=("{0}/v4_6_release/apis/3.0/service/tickets?conditions={1}&orderBy=dateEntered desc&page={2}&pageSize={3}" -f $ConnectWiseUrl.TrimEnd('/'),$encodedConditions,$page,$PageSize)
        Write-DebugLog ("Retrieving ticket page {0} for company ID {1}." -f $page,$CompanyId)
        $response=Invoke-ConnectWiseApi -Uri $uri -Headers $Headers
        $items=@($response | ForEach-Object { $_ })
        if ($items.Count -eq 0) { break }
        $pagesRetrieved++
        $tickets += $items
        if ($items.Count -lt $PageSize) { break }
        if ($page -eq $MaxPages) { $truncated=$true; break }
        $page++
    }
    return [ordered]@{ tickets=@($tickets); conditions=$conditions; pagesRetrieved=$pagesRetrieved; truncated=$truncated }
}

function ConvertTo-TicketSummary {
    param([object]$Ticket)
    $company=Get-ValueOrNull $Ticket "company"
    $board=Get-ValueOrNull $Ticket "board"
    $status=Get-ValueOrNull $Ticket "status"
    $priority=Get-ValueOrNull $Ticket "priority"
    $contact=Get-ValueOrNull $Ticket "contact"
    $site=Get-ValueOrNull $Ticket "site"
    $team=Get-ValueOrNull $Ticket "team"
    $owner=Get-ValueOrNull $Ticket "owner"
    $type=Get-ValueOrNull $Ticket "type"
    $subtype=Get-ValueOrNull $Ticket "subType"
    if ($null -eq $subtype) { $subtype=Get-ValueOrNull $Ticket "subtype" }
    $item=Get-ValueOrNull $Ticket "item"
    $source=Get-ValueOrNull $Ticket "source"
    $agreement=Get-ValueOrNull $Ticket "agreement"
    $info=Get-ValueOrNull $Ticket "_info"
    return [ordered]@{
        id=Get-ValueOrNull $Ticket "id"
        summary=Get-ValueOrNull $Ticket "summary"
        recordType=Get-ValueOrNull $Ticket "recordType"
        closedFlag=Get-ValueOrNull $Ticket "closedFlag"
        companyId=Get-ValueOrNull $company "id"
        companyName=Get-ValueOrNull $company "name"
        boardId=Get-ValueOrNull $board "id"
        boardName=Get-ValueOrNull $board "name"
        statusId=Get-ValueOrNull $status "id"
        statusName=Get-ValueOrNull $status "name"
        priorityId=Get-ValueOrNull $priority "id"
        priorityName=Get-ValueOrNull $priority "name"
        contactId=Get-ValueOrNull $contact "id"
        contactName=Get-ValueOrNull $contact "name"
        siteId=Get-ValueOrNull $site "id"
        siteName=Get-ValueOrNull $site "name"
        teamId=Get-ValueOrNull $team "id"
        teamName=Get-ValueOrNull $team "name"
        ownerId=Get-ValueOrNull $owner "id"
        ownerIdentifier=Get-ValueOrNull $owner "identifier"
        ownerName=Get-ValueOrNull $owner "name"
        resources=Get-ValueOrNull $Ticket "resources"
        typeId=Get-ValueOrNull $type "id"
        typeName=Get-ValueOrNull $type "name"
        subtypeId=Get-ValueOrNull $subtype "id"
        subtypeName=Get-ValueOrNull $subtype "name"
        itemId=Get-ValueOrNull $item "id"
        itemName=Get-ValueOrNull $item "name"
        sourceId=Get-ValueOrNull $source "id"
        sourceName=Get-ValueOrNull $source "name"
        agreementId=Get-ValueOrNull $agreement "id"
        agreementName=Get-ValueOrNull $agreement "name"
        initialDescription=Get-ValueOrNull $Ticket "initialDescription"
        internalAnalysis=Get-ValueOrNull $Ticket "internalAnalysis"
        resolution=Get-ValueOrNull $Ticket "resolution"
        impact=Get-ValueOrNull $Ticket "impact"
        severity=Get-ValueOrNull $Ticket "severity"
        requiredDate=Get-ValueOrNull $Ticket "requiredDate"
        dueDate=Get-ValueOrNull $Ticket "dueDate"
        estimatedStartDate=Get-ValueOrNull $Ticket "estimatedStartDate"
        dateResolved=Get-ValueOrNull $Ticket "dateResolved"
        dateClosed=Get-ValueOrNull $Ticket "dateClosed"
        budgetHours=Get-ValueOrNull $Ticket "budgetHours"
        actualHours=Get-ValueOrNull $Ticket "actualHours"
        isInSla=Get-ValueOrNull $Ticket "isInSla"
        dateEntered=Get-ValueOrNull $info "dateEntered"
        enteredBy=Get-ValueOrNull $info "enteredBy"
        lastUpdated=Get-ValueOrNull $info "lastUpdated"
        updatedBy=Get-ValueOrNull $info "updatedBy"
    }
}

try {
    Write-Info "Starting ConnectWise client ticket retrieval."
    try { $body=Get-RequestBodyObject $Request }
    catch {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";errorCode="InvalidRequestBody";error="Request body is not valid JSON."}
        return
    }
    if ($body -is [Collections.IDictionary]) { $body=[PSCustomObject]$body }

    $CompanyName=Get-InputValue $Request $body @("CompanyName","companyName","Company","company")
    $TicketModeInput=Get-InputValue $Request $body @("TicketMode","ticketMode","Mode","mode")
    $BoardName=Get-InputValue $Request $body @("BoardName","boardName","Board","board")
    $RequestSecurityKey=Get-InputValue $Request $body @("SecurityKey","securityKey")

    if ([string]::IsNullOrWhiteSpace($CompanyName)) {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";errorCode="MissingCompanyName";error="Provide CompanyName in the JSON body, query string, or header."}; return
    }
    $CompanyName=$CompanyName.Trim()
    try { $TicketMode=Get-NormalizedTicketMode $TicketModeInput }
    catch {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";errorCode="InvalidTicketMode";error="TicketMode must be Open, Closed, or All."}
        return
    }
    if (-not [string]::IsNullOrWhiteSpace($BoardName)) { $BoardName=$BoardName.Trim() }

    $configuredKey=$env:SecurityKey
    if (-not [string]::IsNullOrWhiteSpace($configuredKey) -and -not [string]::Equals($configuredKey,$RequestSecurityKey,[StringComparison]::Ordinal)) {
        Send-JsonResponse ([HttpStatusCode]::Unauthorized) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";errorCode="InvalidSecurityKey";error="Invalid security key."}; return
    }

    $required=@("ConnectWisePsa_ApiBaseUrl","ConnectWisePsa_ApiCompanyId","ConnectWisePsa_ApiPublicKey","ConnectWisePsa_ApiPrivateKey","ConnectWisePsa_ApiClientId")
    $missing=@()
    foreach ($setting in $required) {
        $value=Get-Item -Path "Env:$setting" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Value
        if ([string]::IsNullOrWhiteSpace([string]$value)) { $missing += $setting }
    }
    if ($missing.Count -gt 0) {
        Send-JsonResponse ([HttpStatusCode]::InternalServerError) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";errorCode="MissingApplicationSettings";error="Missing required application settings.";missingSettings=$missing}; return
    }

    $cwUrl=$env:ConnectWisePsa_ApiBaseUrl.TrimEnd('/')
    $headers=New-ConnectWiseHeaders $env:ConnectWisePsa_ApiCompanyId $env:ConnectWisePsa_ApiPublicKey $env:ConnectWisePsa_ApiPrivateKey $env:ConnectWisePsa_ApiClientId
    $candidates=@(Find-ConnectWiseCompanies $cwUrl $CompanyName $headers)
    if ($candidates.Count -eq 0) {
        Send-JsonResponse ([HttpStatusCode]::NotFound) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";companyInput=$CompanyName;errorCode="CompanyNotFound";error="No ConnectWise company matched the supplied company name."}; return
    }

    $exact=@($candidates | Where-Object { $_.name -and $_.name.ToString().Trim().Equals($CompanyName,[StringComparison]::OrdinalIgnoreCase) })
    if ($exact.Count -eq 1) { $selectedSummary=$exact[0] }
    elseif ($exact.Count -gt 1 -or $candidates.Count -gt 1) {
        $matchedCompanies=if ($exact.Count -gt 1) { $exact } else { $candidates }
        Send-JsonResponse ([HttpStatusCode]::Conflict) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";companyInput=$CompanyName;errorCode="AmbiguousCompanyName";error="Multiple ConnectWise companies matched. Provide the exact company name.";matches=@($matchedCompanies | ForEach-Object {@{id=$_.id;identifier=$_.identifier;name=$_.name}})}; return
    }
    else { $selectedSummary=$candidates[0] }

    $company=Get-ConnectWiseCompany $cwUrl ([int]$selectedSummary.id) $headers
    $types=@(Get-CompanyTypeNames $company)
    if (-not (Test-CompanyIsClient $company)) {
        Send-JsonResponse ([HttpStatusCode]::Forbidden) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";companyInput=$CompanyName;company=@{id=$company.id;identifier=$company.identifier;name=$company.name;companyTypes=$types};errorCode="CompanyIsNotClient";error="The matching ConnectWise company does not have the company type Client."}; return
    }

    $pageSize=100; $parsedPageSize=0
    if ($env:ConnectWisePsa_TicketsPageSize -and [int]::TryParse([string]$env:ConnectWisePsa_TicketsPageSize,[ref]$parsedPageSize) -and $parsedPageSize -gt 0) { $pageSize=[Math]::Min($parsedPageSize,1000) }
    $maxPages=25; $parsedMaxPages=0
    if ($env:ConnectWisePsa_TicketsMaxPages -and [int]::TryParse([string]$env:ConnectWisePsa_TicketsMaxPages,[ref]$parsedMaxPages) -and $parsedMaxPages -gt 0) { $maxPages=[Math]::Min($parsedMaxPages,1000) }

    $allowedBoards=@(Get-AllowedBoards)
    $allowedNormalized=@($allowedBoards | ForEach-Object { $_.ToString().Trim().ToLowerInvariant() })
    if (-not [string]::IsNullOrWhiteSpace($BoardName) -and $allowedNormalized -notcontains $BoardName.ToLowerInvariant()) {
        Send-JsonResponse ([HttpStatusCode]::Forbidden) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";boardName=$BoardName;errorCode="BoardNotAllowed";error="Requested board is not in the allowed boards list.";allowedBoards=$allowedBoards}; return
    }
    $namesToResolve=if ([string]::IsNullOrWhiteSpace($BoardName)) { $allowedNormalized } else { @($BoardName.ToLowerInvariant()) }
    $allBoards=Get-ConnectWiseBoards $cwUrl $headers
    $boardIds=@($allBoards | Where-Object { $_.name -and $namesToResolve -contains $_.name.ToString().Trim().ToLowerInvariant() } | ForEach-Object { [int]$_.id } | Select-Object -Unique)
    if ($boardIds.Count -eq 0) {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";errorCode="NoResolvableBoards";error="No allowed boards exist in ConnectWise.";allowedBoards=$allowedBoards}; return
    }

    $retrieval=Get-ConnectWiseClientTickets -ConnectWiseUrl $cwUrl -CompanyId ([int]$company.id) -TicketMode $TicketMode -BoardIds $boardIds -Headers $headers -PageSize $pageSize -MaxPages $maxPages
    $tickets=@($retrieval.tickets)
    $normalized=@($tickets | ForEach-Object { ConvertTo-TicketSummary $_ })
    $openCount=@($tickets | Where-Object { (Get-ValueOrNull $_ "closedFlag") -ne $true }).Count
    $closedCount=@($tickets | Where-Object { (Get-ValueOrNull $_ "closedFlag") -eq $true }).Count

    Write-Info ("Retrieved {0} {1} ticket(s) for company {2}." -f $tickets.Count,$TicketMode,$company.name)
    $response=[ordered]@{
        success=$true
        correlationId=$CorrelationId
        action="GetClientTickets"
        companyInput=$CompanyName
        retrievedAtUtc=(Get-Date).ToUniversalTime().ToString('o')
        company=[ordered]@{id=$company.id;identifier=$company.identifier;name=$company.name;companyTypes=$types;isClient=$true;status=Get-ValueOrNull $company "status"}
        criteria=[ordered]@{ticketMode=$TicketMode;boardName=$BoardName;boardIds=$boardIds;allowedBoards=$allowedBoards;conditions=$retrieval.conditions;pageSize=$pageSize;maxPages=$maxPages}
        ticketSummary=[ordered]@{returnedTicketCount=$tickets.Count;openTicketCount=$openCount;closedTicketCount=$closedCount;pagesRetrieved=$retrieval.pagesRetrieved;truncated=$retrieval.truncated}
        tickets=$normalized
        raw=[ordered]@{company=$company;tickets=$tickets}
    }
    Send-JsonResponse ([HttpStatusCode]::OK) $response
}
catch {
    Write-ErrorLog $_.Exception.Message
    Send-JsonResponse ([HttpStatusCode]::InternalServerError) @{success=$false;correlationId=$CorrelationId;action="GetClientTickets";errorCode="UnhandledFunctionError";error="An unexpected error occurred while retrieving client tickets."}
}
