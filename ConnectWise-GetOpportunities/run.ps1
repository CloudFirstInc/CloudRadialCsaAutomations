<#
.SYNOPSIS
    Retrieves open ConnectWise Manage sales opportunities for a Client company.

.DESCRIPTION
    Resolves a ConnectWise company from CompanyId (preferred) or CompanyName,
    validates that it is a Client, and returns up to 100 open opportunities with
    normalized pipeline details and the raw opportunity records. If more than
    100 open opportunities exist, the response is flagged as truncated.
    This function is read-only.

.REQUIRED ENVIRONMENT VARIABLES
    ConnectWisePsa_ApiBaseUrl
    ConnectWisePsa_ApiCompanyId
    ConnectWisePsa_ApiPublicKey
    ConnectWisePsa_ApiPrivateKey
    ConnectWisePsa_ApiClientId
    SecurityKey                           - Optional shared secret

.OPTIONAL ENVIRONMENT VARIABLES
    DebugLogging                          - Set to 1 for diagnostic logs

.INPUTS
    Provide CompanyId or CompanyName. When both are supplied, CompanyId is used.
    JSON body: { "CompanyId": 20385, "SecurityKey": "optional" }
    JSON body: { "CompanyName": "City of Rye", "SecurityKey": "optional" }
    Query string: ?CompanyId=20385 or ?CompanyName=City%20of%20Rye

.OUTPUTS
    Company details, retrieval completeness, opportunity counts, normalized
    opportunity pipeline summaries, and complete raw open opportunity records.
#>

using namespace System.Net
param($Request, $TriggerMetadata)

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

function Send-JsonResponse {
    param(
        [System.Net.HttpStatusCode]$StatusCode,
        [object]$Body
    )

    $json = ConvertTo-Json -InputObject $Body -Depth 100
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = $StatusCode
        Body        = $json
        ContentType = "application/json"
    })
}

function Get-ValueOrNull {
    param(
        [object]$Object,
        [string]$PropertyName
    )

    if ($null -eq $Object) {
        return $null
    }

    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) {
            if ([string]::Equals([string]$key, $PropertyName, [StringComparison]::OrdinalIgnoreCase)) {
                return $Object[$key]
            }
        }
        return $null
    }

    $property = $Object.PSObject.Properties |
        Where-Object { $_.Name.Equals($PropertyName, [StringComparison]::OrdinalIgnoreCase) } |
        Select-Object -First 1

    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-InputValue {
    param(
        $Request,
        $Body,
        [string[]]$Names
    )

    foreach ($name in $Names) {
        foreach ($source in @($Body, $Request.Query, $Request.Headers)) {
            if ($null -eq $source) {
                continue
            }

            if ($source -is [System.Collections.IDictionary]) {
                foreach ($key in $source.Keys) {
                    if ([string]::Equals([string]$key, $name, [StringComparison]::OrdinalIgnoreCase)) {
                        $value = $source[$key]
                        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                            return [string]$value
                        }
                    }
                }
            }
            else {
                $value = Get-ValueOrNull -Object $source -PropertyName $name
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    return [string]$value
                }
            }
        }
    }

    return $null
}

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
            return ConvertFrom-Json -InputObject $Request.Body -ErrorAction Stop
        }
        catch {
            throw [System.ArgumentException]::new("Request body is not valid JSON.")
        }
    }

    return $Request.Body
}

function ConvertTo-FlatArray {
    param([object]$Value)

    $items = [System.Collections.Generic.List[object]]::new()
    $pending = [System.Collections.Stack]::new()
    $pending.Push($Value)

    while ($pending.Count -gt 0) {
        $current = $pending.Pop()
        if ($null -eq $current) {
            continue
        }
        if ($current -is [System.Array]) {
            for ($index = $current.Length - 1; $index -ge 0; $index--) {
                $pending.Push($current[$index])
            }
        }
        else {
            $items.Add($current)
        }
    }

    return ,$items.ToArray()
}

function New-ConnectWiseHeaders {
    param(
        [string]$CompanyId,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId
    )

    $authString = "{0}+{1}:{2}" -f $CompanyId, $PublicKey, $PrivateKey
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($authString))

    return @{
        Authorization = "Basic {0}" -f $encodedAuth
        "Content-Type" = "application/json"
        Accept         = "application/json"
        ClientId       = $ClientId
    }
}

function Invoke-ConnectWiseApi {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [int]$MaxRetries = 5
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-DebugLog ("GET {0}, attempt {1}" -f $Uri, $attempt)
            return Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers -TimeoutSec 100 -ErrorAction Stop
        }
        catch {
            $statusCode = $null
            $retryAfter = $null
            if ($_.Exception.Response) {
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
                try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch { }
            }

            Write-ErrorLog ("ConnectWise GET failed; status={0}; {1}" -f $statusCode, $_.Exception.Message)
            $retryable = $statusCode -in @(408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -ge $MaxRetries) {
                throw
            }

            $delay = 0
            if ($retryAfter -and [int]::TryParse([string]$retryAfter, [ref]$delay)) {
                $delay = [Math]::Min($delay, 60)
            }
            if ($delay -lt 1) {
                $delay = [Math]::Min(60, [int][Math]::Pow(2, $attempt) + (Get-Random -Minimum 1 -Maximum 4))
            }
            Write-Info ("Retrying ConnectWise request in {0} seconds." -f $delay)
            Start-Sleep -Seconds $delay
        }
    }
}

function ConvertTo-CwConditionString {
    param([string]$Value)
    return $Value.Replace("'", "''")
}

function ConvertTo-CwUrlEncoded {
    param([string]$Value)
    return [Uri]::EscapeDataString($Value)
}

function Find-ConnectWiseCompanies {
    param(
        [string]$ConnectWiseUrl,
        [string]$CompanyName,
        [hashtable]$Headers
    )

    $safeName = ConvertTo-CwConditionString $CompanyName
    $conditions = ConvertTo-CwUrlEncoded ("name contains '{0}'" -f $safeName)
    $uri = "{0}/v4_6_release/apis/3.0/company/companies?conditions={1}&orderBy=name&page=1&pageSize=1000" -f $ConnectWiseUrl, $conditions
    return ,(ConvertTo-FlatArray (Invoke-ConnectWiseApi -Uri $uri -Headers $Headers))
}

function Get-ConnectWiseCompany {
    param(
        [string]$ConnectWiseUrl,
        [int]$CompanyId,
        [hashtable]$Headers
    )

    $uri = "{0}/v4_6_release/apis/3.0/company/companies/{1}" -f $ConnectWiseUrl, $CompanyId
    return Invoke-ConnectWiseApi -Uri $uri -Headers $Headers
}

function Get-CompanyTypeNames {
    param([object]$Company)

    $names = @()
    foreach ($type in (ConvertTo-FlatArray (Get-ValueOrNull $Company "types"))) {
        if ($type -is [string]) {
            $names += $type.Trim()
        }
        else {
            $name = Get-ValueOrNull $type "name"
            if (-not [string]::IsNullOrWhiteSpace([string]$name)) {
                $names += ([string]$name).Trim()
            }
        }
    }

    $singleType = Get-ValueOrNull $Company "type"
    if ($singleType -is [string]) {
        $names += $singleType.Trim()
    }
    elseif ($null -ne $singleType) {
        $singleName = Get-ValueOrNull $singleType "name"
        if (-not [string]::IsNullOrWhiteSpace([string]$singleName)) {
            $names += ([string]$singleName).Trim()
        }
    }

    return @($names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Test-CompanyIsClient {
    param([object]$Company)

    foreach ($typeName in (Get-CompanyTypeNames $Company)) {
        if ($typeName.Equals("Client", [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Test-OpportunityOpen {
    param([object]$Opportunity)

    $closedFlag = Get-ValueOrNull $Opportunity "closedFlag"
    if ($closedFlag -is [bool]) {
        return (-not $closedFlag)
    }
    if ($null -eq $closedFlag) {
        return $false
    }

    $parsed = $false
    if ([bool]::TryParse([string]$closedFlag, [ref]$parsed)) {
        return (-not $parsed)
    }
    if ([string]$closedFlag -eq "0") {
        return $true
    }
    if ([string]$closedFlag -eq "1") {
        return $false
    }

    return $false
}

function Get-ConnectWiseCompanyOpportunities {
    param(
        [string]$ConnectWiseUrl,
        [int]$CompanyId,
        [hashtable]$Headers,
        [int]$Limit = 100
    )

    # Request one extra record to detect whether more open opportunities exist.
    $condition = ConvertTo-CwUrlEncoded ("company/id={0} and closedFlag=false" -f $CompanyId)
    $uri = "{0}/v4_6_release/apis/3.0/sales/opportunities?conditions={1}&orderBy=id%20desc&page=1&pageSize={2}" -f $ConnectWiseUrl, $condition, ($Limit + 1)
    Write-DebugLog ("Retrieving up to {0} open opportunities for company ID {1}." -f $Limit, $CompanyId)

    $items = ConvertTo-FlatArray (Invoke-ConnectWiseApi -Uri $uri -Headers $Headers)
    $truncated = $items.Count -gt $Limit

    return [ordered]@{
        opportunities = @($items | Select-Object -First $Limit)
        truncated = $truncated
    }
}

function ConvertTo-OpportunitySummary {
    param([object]$Opportunity)

    $company = Get-ValueOrNull $Opportunity "company"
    $contact = Get-ValueOrNull $Opportunity "contact"
    $site = Get-ValueOrNull $Opportunity "site"
    $status = Get-ValueOrNull $Opportunity "status"
    $type = Get-ValueOrNull $Opportunity "type"
    $source = Get-ValueOrNull $Opportunity "source"
    $rating = Get-ValueOrNull $Opportunity "rating"
    $primarySalesRep = Get-ValueOrNull $Opportunity "primarySalesRep"
    $secondarySalesRep = Get-ValueOrNull $Opportunity "secondarySalesRep"
    $territory = Get-ValueOrNull $Opportunity "territory"
    $currency = Get-ValueOrNull $Opportunity "currency"
    $competitor = Get-ValueOrNull $Opportunity "competitor"
    $info = Get-ValueOrNull $Opportunity "_info"

    return [ordered]@{
        id = Get-ValueOrNull $Opportunity "id"
        name = Get-ValueOrNull $Opportunity "name"
        closedFlag = Get-ValueOrNull $Opportunity "closedFlag"
        companyId = Get-ValueOrNull $company "id"
        companyName = Get-ValueOrNull $company "name"
        contactId = Get-ValueOrNull $contact "id"
        contactName = Get-ValueOrNull $contact "name"
        siteId = Get-ValueOrNull $site "id"
        siteName = Get-ValueOrNull $site "name"
        statusId = Get-ValueOrNull $status "id"
        statusName = Get-ValueOrNull $status "name"
        typeId = Get-ValueOrNull $type "id"
        typeName = Get-ValueOrNull $type "name"
        sourceId = Get-ValueOrNull $source "id"
        sourceName = Get-ValueOrNull $source "name"
        ratingId = Get-ValueOrNull $rating "id"
        ratingName = Get-ValueOrNull $rating "name"
        primarySalesRepId = Get-ValueOrNull $primarySalesRep "id"
        primarySalesRepName = Get-ValueOrNull $primarySalesRep "name"
        secondarySalesRepId = Get-ValueOrNull $secondarySalesRep "id"
        secondarySalesRepName = Get-ValueOrNull $secondarySalesRep "name"
        territoryId = Get-ValueOrNull $territory "id"
        territoryName = Get-ValueOrNull $territory "name"
        currencyId = Get-ValueOrNull $currency "id"
        currencyName = Get-ValueOrNull $currency "name"
        opportunityValue = Get-ValueOrNull $Opportunity "opportunityValue"
        probability = Get-ValueOrNull $Opportunity "probability"
        expectedCloseDate = Get-ValueOrNull $Opportunity "expectedCloseDate"
        actualCloseDate = Get-ValueOrNull $Opportunity "actualCloseDate"
        closedDate = Get-ValueOrNull $Opportunity "closedDate"
        nextStep = Get-ValueOrNull $Opportunity "nextStep"
        description = Get-ValueOrNull $Opportunity "description"
        notes = Get-ValueOrNull $Opportunity "notes"
        customerPO = Get-ValueOrNull $Opportunity "customerPO"
        externalReference = Get-ValueOrNull $Opportunity "externalReference"
        competitorId = Get-ValueOrNull $competitor "id"
        competitorName = Get-ValueOrNull $competitor "name"
        forecastCategory = Get-ValueOrNull $Opportunity "forecastCategory"
        dateEntered = Get-ValueOrNull $info "dateEntered"
        enteredBy = Get-ValueOrNull $info "enteredBy"
        lastUpdated = Get-ValueOrNull $info "lastUpdated"
        updatedBy = Get-ValueOrNull $info "updatedBy"
    }
}

function Get-PositiveIntSetting {
    param(
        [string]$Name,
        [int]$DefaultValue,
        [int]$Maximum
    )

    $parsed = 0
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $DefaultValue
    }
    if (-not [int]::TryParse($value, [ref]$parsed) -or $parsed -lt 1) {
        throw [System.InvalidOperationException]::new("$Name must be a positive integer.")
    }
    return [Math]::Min($parsed, $Maximum)
}

try {
    Write-Info "Starting ConnectWise open-opportunity retrieval."
    try {
        $body = Get-RequestBodyObject $Request
    }
    catch {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{
            success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
            errorCode = "InvalidRequestBody"; error = "Request body is not valid JSON."
        }
        return
    }

    $companyName = Get-InputValue $Request $body @("CompanyName", "companyName", "Company", "company")
    $companyIdInput = Get-InputValue $Request $body @("CompanyId", "companyId")
    $requestSecurityKey = Get-InputValue $Request $body @("SecurityKey", "securityKey")
    $hasCompanyId = -not [string]::IsNullOrWhiteSpace([string]$companyIdInput)
    $hasCompanyName = -not [string]::IsNullOrWhiteSpace([string]$companyName)
    if (-not $hasCompanyId -and -not $hasCompanyName) {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{
            success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
            errorCode = "MissingCompanyIdentifier"; error = "Provide CompanyId or CompanyName in the JSON body or query string."
        }
        return
    }
    $requestedCompanyId = 0
    if ($hasCompanyId -and (-not [int]::TryParse(([string]$companyIdInput).Trim(), [ref]$requestedCompanyId) -or $requestedCompanyId -lt 1)) {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{
            success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
            errorCode = "InvalidCompanyId"; error = "CompanyId must be a positive integer."
        }
        return
    }
    $companyName = if ($hasCompanyName) { ([string]$companyName).Trim() } else { $null }
    $companyInput = if ($hasCompanyId) { [string]$requestedCompanyId } else { $companyName }

    $configuredKey = $env:SecurityKey
    if (-not [string]::IsNullOrWhiteSpace($configuredKey) -and
        -not [string]::Equals($configuredKey, $requestSecurityKey, [StringComparison]::Ordinal)) {
        Send-JsonResponse ([HttpStatusCode]::Unauthorized) @{
            success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
            errorCode = "InvalidSecurityKey"; error = "Invalid security key."
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
    $missingSettings = @(
        $requiredSettings | Where-Object {
            [string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($_))
        }
    )
    if ($missingSettings.Count -gt 0) {
        Send-JsonResponse ([HttpStatusCode]::InternalServerError) @{
            success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
            errorCode = "MissingApplicationSettings"; error = "Missing required application settings."
            missingSettings = $missingSettings
        }
        return
    }

    $connectWiseUrl = $env:ConnectWisePsa_ApiBaseUrl.TrimEnd("/")
    $headers = New-ConnectWiseHeaders `
        -CompanyId $env:ConnectWisePsa_ApiCompanyId `
        -PublicKey $env:ConnectWisePsa_ApiPublicKey `
        -PrivateKey $env:ConnectWisePsa_ApiPrivateKey `
        -ClientId $env:ConnectWisePsa_ApiClientId

    $companyId = 0
    if ($hasCompanyId) {
        $companyId = $requestedCompanyId
    }
    else {
    Write-Info ("Searching ConnectWise for company '{0}'." -f $companyName)
    $candidates = Find-ConnectWiseCompanies $connectWiseUrl $companyName $headers
    if ($candidates.Count -eq 0) {
        Send-JsonResponse ([HttpStatusCode]::NotFound) @{
            success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
            companyInput = $companyName; errorCode = "CompanyNotFound"
            error = "No ConnectWise company matched the supplied name."
        }
        return
    }

    $exactMatches = @(
        $candidates | Where-Object {
            $null -ne $_.name -and $_.name.ToString().Trim().Equals($companyName, [StringComparison]::OrdinalIgnoreCase)
        }
    )
    if ($exactMatches.Count -eq 1) {
        $selectedCompanySummary = $exactMatches[0]
    }
    elseif ($exactMatches.Count -gt 1 -or $candidates.Count -gt 1) {
        $ambiguousMatches = if ($exactMatches.Count -gt 1) { $exactMatches } else { $candidates }
        Send-JsonResponse ([HttpStatusCode]::Conflict) @{
            success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
            companyInput = $companyName; errorCode = "AmbiguousCompanyName"
            error = "Multiple ConnectWise companies matched. Provide the exact company name."
            matches = @($ambiguousMatches | ForEach-Object {
                @{ id = $_.id; identifier = $_.identifier; name = $_.name }
            })
        }
        return
    }
    else {
        $selectedCompanySummary = $candidates[0]
    }

    if (-not [int]::TryParse([string]$selectedCompanySummary.id, [ref]$companyId) -or $companyId -lt 1) {
        throw [System.InvalidOperationException]::new("ConnectWise returned an invalid company ID.")
    }
    }

    try {
        $company = Get-ConnectWiseCompany $connectWiseUrl $companyId $headers
    }
    catch {
        if ($hasCompanyId -and $_.Exception.Message -match "404|Not Found") {
            Send-JsonResponse ([HttpStatusCode]::NotFound) @{
                success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
                companyInput = $companyInput; errorCode = "CompanyNotFound"
                error = "No ConnectWise company exists with the supplied CompanyId."
            }
            return
        }
        throw
    }
    $companyTypes = @(Get-CompanyTypeNames $company)
    if (-not (Test-CompanyIsClient $company)) {
        Send-JsonResponse ([HttpStatusCode]::Forbidden) @{
            success = $false; correlationId = $CorrelationId; action = "GetClientOpportunities"
            companyInput = $companyInput
            company = @{ id = $company.id; identifier = $company.identifier; name = $company.name; companyTypes = $companyTypes }
            errorCode = "CompanyIsNotClient"; error = "The matching ConnectWise company is not a Client."
        }
        return
    }

    $maxRecords = 100
    $retrieval = Get-ConnectWiseCompanyOpportunities `
        -ConnectWiseUrl $connectWiseUrl `
        -CompanyId $companyId `
        -Headers $headers `
        -Limit $maxRecords

    $allOpportunities = ConvertTo-FlatArray $retrieval.opportunities
    $openOpportunities = @(
        $allOpportunities | Where-Object { Test-OpportunityOpen $_ }
    )
    $summaries = @(
        $openOpportunities | ForEach-Object { ConvertTo-OpportunitySummary $_ }
    )

    Write-Info ("Retrieved {0} open opportunity record(s) for {1}." -f $openOpportunities.Count, $company.name)
    $response = [ordered]@{
        success = $true
        correlationId = $CorrelationId
        action = "GetClientOpportunities"
        companyInput = $companyInput
        retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        company = [ordered]@{
            id = $company.id
            identifier = $company.identifier
            name = $company.name
            companyTypes = $companyTypes
            isClient = $true
            status = Get-ValueOrNull $company "status"
        }
        retrieval = [ordered]@{
            scope = "Open opportunities where closedFlag is false, newest first (id descending), limited to the top 100."
            maxRecords = $maxRecords
            truncated = $retrieval.truncated
            complete = (-not $retrieval.truncated)
        }
        opportunitySummary = [ordered]@{
            retrievedOpportunityCount = $allOpportunities.Count
            openOpportunityCount = $openOpportunities.Count
        }
        opportunities = $summaries
        raw = [ordered]@{
            company = $company
            openOpportunities = $openOpportunities
        }
    }

    Send-JsonResponse ([HttpStatusCode]::OK) $response
}
catch {
    Write-ErrorLog $_.Exception.Message
    Send-JsonResponse ([HttpStatusCode]::InternalServerError) @{
        success = $false
        correlationId = $CorrelationId
        action = "GetClientOpportunities"
        errorCode = "UnhandledFunctionError"
        error = "An unexpected error occurred while retrieving client opportunities."
    }
}
