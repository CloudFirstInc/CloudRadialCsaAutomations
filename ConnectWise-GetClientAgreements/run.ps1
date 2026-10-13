<#
.SYNOPSIS
    Retrieves all active ConnectWise Manage agreements for a Client company.

.DESCRIPTION
    This Azure Function accepts a company identifier (preferred) or company
    name string, resolves the company in ConnectWise Manage, validates that
    the selected company has the company type "Client", retrieves the company's
    agreements, and returns all agreements that are active as of the current
    UTC date.

    Each agreement includes its complete ConnectWise agreement record,
    including the internalNotes property returned by the agreement endpoint.

    The function is read-only and does not modify company or agreement records.

.REQUIRED ENVIRONMENT VARIABLES
    ConnectWisePsa_ApiBaseUrl      - Base URL of the ConnectWise Manage API
    ConnectWisePsa_ApiCompanyId    - ConnectWise API company identifier
    ConnectWisePsa_ApiPublicKey    - ConnectWise public key
    ConnectWisePsa_ApiPrivateKey   - ConnectWise private key
    ConnectWisePsa_ApiClientId     - ConnectWise client ID
    SecurityKey                    - Optional shared secret

.OPTIONAL ENVIRONMENT VARIABLES
    ConnectWisePsa_AgreementsPageSize - Agreement retrieval page size
                                        Default: 100
                                        Maximum: 1000
    DebugLogging                      - Set to 1 for verbose diagnostic logs

.INPUTS
    Provide CompanyId (preferred) or CompanyName. If both are supplied, CompanyId
    is used and the name search is skipped.

    JSON body:
    {
        "CompanyId": 20385,
        "CompanyName": "City of Rye",
        "SecurityKey": "optional"
    }

    The company identifier can also be supplied through the query string:

    ?CompanyId=20385
    ?CompanyName=City%20of%20Rye

.OUTPUTS
    JSON response containing:

    - Matching ConnectWise company
    - Company type validation
    - Active agreement count
    - Normalized agreement summaries
    - Complete raw agreement records
    - Agreement internal notes

.NOTES
    This function is read-only.

    An agreement is treated as active when:

    - cancelledFlag is not true
    - startDate is blank or on/before the current UTC date
    - noEndingDateFlag is true, or endDate is blank or on/after the current UTC date
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

    $json = $Body | ConvertTo-Json -Depth 100

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

    $property = $Object.PSObject.Properties[$PropertyName]

    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

# -------------------------------
# Utility: Build ConnectWise headers
# -------------------------------

function New-ConnectWiseHeaders {
    param(
        [string]$CompanyId,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId
    )

    $authString = "{0}+{1}:{2}" -f $CompanyId, $PublicKey, $PrivateKey

    $encodedAuth = [Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes($authString)
    )

    return @{
        "Authorization" = "Basic {0}" -f $encodedAuth
        "Content-Type"  = "application/json"
        "Accept"        = "application/json"
        "ClientId"      = $ClientId
    }
}

# -------------------------------
# Utility: Invoke ConnectWise API
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

            Write-DebugLog (
                "CW API {0} {1}, attempt {2}" -f
                $Method,
                $Uri,
                $attempt
            )

            if ($null -ne $Body) {
                $bodyJson = $Body | ConvertTo-Json -Depth 20

                return Invoke-RestMethod `
                    -Uri $Uri `
                    -Method $Method `
                    -Headers $Headers `
                    -Body $bodyJson `
                    -ErrorAction Stop
            }

            return Invoke-RestMethod `
                -Uri $Uri `
                -Method $Method `
                -Headers $Headers `
                -ErrorAction Stop
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
                    $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                }
                catch {
                    $retryAfter = $null
                }
            }

            $message = $_.Exception.Message

            Write-ErrorLog (
                "CW API call failed. StatusCode={0}. Message={1}" -f
                $statusCode,
                $message
            )

            $retryable = (
                $statusCode -in @(408, 429, 500, 502, 503, 504)
            )

            if (-not $retryable -or $attempt -ge $MaxRetries) {
                throw
            }

            $sleepSeconds = 0

            if ($retryAfter) {
                $parsedRetryAfter = 0

                if (
                    [int]::TryParse(
                        [string]$retryAfter,
                        [ref]$parsedRetryAfter
                    )
                ) {
                    $sleepSeconds = $parsedRetryAfter
                }
            }

            if ($sleepSeconds -lt 1) {
                $sleepSeconds = [Math]::Pow(2, $attempt + (Get-Random -Minimum 1 -Maximum 4)
                )
            }

            Write-Info (
                "Retrying ConnectWise request in {0} seconds." -f
                $sleepSeconds
            )

            Start-Sleep -Seconds $sleepSeconds
        }
    }
}

# -------------------------------
# Utility: URL encoding
# -------------------------------

function ConvertTo-CwUrlEncoded {
    param([string]$Value)

    if ($null -eq $Value) {
        return $null
    }

    return [System.Uri]::EscapeDataString($Value)
}

function ConvertTo-CwConditionString {
    param([string]$Value)

    if ($null -eq $Value) {
        return $null
    }

    # Escape apostrophes in ConnectWise condition string literals.
    return $Value.Replace("'", "''")
}

# -------------------------------
# Utility: Input parsing
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
            if ($Body -is [System.Collections.IDictionary]) {
                foreach ($key in $Body.Keys) {
                    if (
                        $null -ne $key -and
                        $key.ToString().Equals(
                            $name,
                            [System.StringComparison]::OrdinalIgnoreCase
                        )
                    ) {
                        $value = $Body[$key]

                        if (
                            -not [string]::IsNullOrWhiteSpace(
                                [string]$value
                            )
                        ) {
                            return [string]$value
                        }
                    }
                }
            }
            else {
                $property = $Body.PSObject.Properties |
                    Where-Object {
                        $_.Name.Equals(
                            $name,
                            [System.StringComparison]::OrdinalIgnoreCase
                        )
                    } |
                    Select-Object -First 1

                if (
                    $property -and
                    -not [string]::IsNullOrWhiteSpace(
                        [string]$property.Value
                    )
                ) {
                    return [string]$property.Value
                }
            }
        }

        if ($Request.Query) {
            foreach ($key in $Request.Query.Keys) {
                if (
                    $null -ne $key -and
                    $key.ToString().Equals(
                        $name,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                ) {
                    $value = $Request.Query[$key]

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$value
                        )
                    ) {
                        return [string]$value
                    }
                }
            }
        }

        if ($Request.Headers) {
            foreach ($key in $Request.Headers.Keys) {
                if (
                    $null -ne $key -and
                    $key.ToString().Equals(
                        $name,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                ) {
                    $value = $Request.Headers[$key]

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$value
                        )
                    ) {
                        return [string]$value
                    }
                }
            }
        }
    }

    return $null
}

# -------------------------------
# Utility: Date conversion
# -------------------------------

function ConvertTo-UtcDateOrNull {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value

    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsedDate = [datetime]::MinValue

    if (
        [datetime]::TryParse(
            $text,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal,
            [ref]$parsedDate
        )
    ) {
        return $parsedDate.ToUniversalTime()
    }

    return $null
}

# -------------------------------
# ConnectWise: Company retrieval
# -------------------------------

function Find-ConnectWiseCompanies {
    param(
        [string]$ConnectWiseUrl,
        [string]$CompanyName,
        [hashtable]$Headers
    )

    $safeCompanyName = ConvertTo-CwConditionString -Value $CompanyName

    # Retrieve candidate companies containing the supplied name.
    # Exact matching is performed locally to avoid choosing the wrong client.
    $conditions = "name contains '$safeCompanyName'"
    $encodedConditions = ConvertTo-CwUrlEncoded -Value $conditions

    $apiUrl = (
        "{0}/v4_6_release/apis/3.0/company/companies" +
        "?conditions={1}&orderBy=name&pageSize=1000"
    ) -f $ConnectWiseUrl.TrimEnd("/"), $encodedConditions

    $response = Invoke-ConnectWiseApi `
        -Uri $apiUrl `
        -Method Get `
        -Headers $Headers

    # Enumerate the REST array so each company is a single object, not a nested Object[].
    return @(
        $response | ForEach-Object { $_ }
    )
}

function Get-ConnectWiseCompany {
    param(
        [string]$ConnectWiseUrl,
        [int]$CompanyId,
        [hashtable]$Headers
    )

    $apiUrl = (
        "{0}/v4_6_release/apis/3.0/company/companies/{1}"
    ) -f $ConnectWiseUrl.TrimEnd("/"), $CompanyId

    return Invoke-ConnectWiseApi `
        -Uri $apiUrl `
        -Method Get `
        -Headers $Headers
}

function Get-CompanyTypeNames {
    param([object]$Company)

    $typeNames = @()

    # Most ConnectWise company records expose company types through "types".
    $types = Get-ValueOrNull -Object $Company -PropertyName "types"

    foreach ($type in @($types)) {
        if ($null -eq $type) {
            continue
        }

        if ($type -is [string]) {
            if (-not [string]::IsNullOrWhiteSpace($type)) {
                $typeNames += $type.Trim()
            }
        }
        elseif ($type.PSObject.Properties["name"]) {
            if (
                -not [string]::IsNullOrWhiteSpace(
                    [string]$type.name
                )
            ) {
                $typeNames += ([string]$type.name).Trim()
            }
        }
    }

    # Compatibility fallback for records exposing singular type.
    $singleType = Get-ValueOrNull -Object $Company -PropertyName "type"

    if ($singleType) {
        if ($singleType -is [string]) {
            $typeNames += $singleType.Trim()
        }
        elseif ($singleType.PSObject.Properties["name"]) {
            $typeNames += ([string]$singleType.name).Trim()
        }
    }

    return @(
        $typeNames |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace([string]$_)
            } |
            Select-Object -Unique
    )
}

function Test-CompanyIsClient {
    param([object]$Company)

    $typeNames = Get-CompanyTypeNames -Company $Company

    foreach ($typeName in $typeNames) {
        if (
            $typeName.Equals(
                "Client",
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            return $true
        }
    }

    return $false
}

# -------------------------------
# ConnectWise: Agreement retrieval
# -------------------------------

function Get-ConnectWiseCompanyAgreements {
    param(
        [string]$ConnectWiseUrl,
        [int]$CompanyId,
        [hashtable]$Headers,
        [int]$PageSize = 100
    )

    if ($PageSize -lt 1) {
        $PageSize = 100
    }

    if ($PageSize -gt 1000) {
        $PageSize = 1000
    }

    $agreements = @()
    $page = 1

    while ($true) {
        $conditions = "company/id=$CompanyId"
        $encodedConditions = ConvertTo-CwUrlEncoded -Value $conditions

        $apiUrl = (
            "{0}/v4_6_release/apis/3.0/finance/agreements" +
            "?conditions={1}" +
            "&orderBy=id" +
            "&page={2}" +
            "&pageSize={3}"
        ) -f (
            $ConnectWiseUrl.TrimEnd("/"),
            $encodedConditions,
            $page,
            $PageSize
        )

        Write-DebugLog (
            "Retrieving agreements page {0} for company ID {1}." -f
            $page,
            $CompanyId
        )

        $pageResults = @(
            Invoke-ConnectWiseApi `
                -Uri $apiUrl `
                -Method Get `
                -Headers $Headers
        )

        if ($pageResults.Count -eq 0) {
            break
        }

        $agreements += $pageResults

        if ($pageResults.Count -lt $PageSize) {
            break
        }

        $page++
    }

    return @($agreements)
}

function Test-AgreementIsActive {
    param(
        [object]$Agreement,
        [datetime]$AsOfUtc
    )

    $cancelledFlag = Get-ValueOrNull `
        -Object $Agreement `
        -PropertyName "cancelledFlag"

    if ($cancelledFlag -eq $true) {
        return $false
    }

    $startDate = ConvertTo-UtcDateOrNull -Value (
        Get-ValueOrNull `
            -Object $Agreement `
            -PropertyName "startDate"
    )

    if ($null -ne $startDate -and $startDate.Date -gt $AsOfUtc.Date) {
        return $false
    }

    $noEndingDateFlag = Get-ValueOrNull `
        -Object $Agreement `
        -PropertyName "noEndingDateFlag"

    if ($noEndingDateFlag -eq $true) {
        return $true
    }

    $endDate = ConvertTo-UtcDateOrNull -Value (
        Get-ValueOrNull `
            -Object $Agreement `
            -PropertyName "endDate"
    )

    if ($null -ne $endDate -and $endDate.Date -lt $AsOfUtc.Date) {
        return $false
    }

    return $true
}

function ConvertTo-AgreementSummary {
    param([object]$Agreement)

    $company = Get-ValueOrNull `
        -Object $Agreement `
        -PropertyName "company"

    $type = Get-ValueOrNull `
        -Object $Agreement `
        -PropertyName "type"

    $contact = Get-ValueOrNull `
        -Object $Agreement `
        -PropertyName "contact"

    $billToCompany = Get-ValueOrNull `
        -Object $Agreement `
        -PropertyName "billToCompany"

    $internalNotes = Get-ValueOrNull `
        -Object $Agreement `
        -PropertyName "internalNotes"

    return [ordered]@{
        id                       = Get-ValueOrNull -Object $Agreement -PropertyName "id"
        name                     = Get-ValueOrNull -Object $Agreement -PropertyName "name"

        companyId                = Get-ValueOrNull -Object $company -PropertyName "id"
        companyName              = Get-ValueOrNull -Object $company -PropertyName "name"

        typeId                   = Get-ValueOrNull -Object $type -PropertyName "id"
        typeName                 = Get-ValueOrNull -Object $type -PropertyName "name"

        status                   = Get-ValueOrNull -Object $Agreement -PropertyName "status"
        startDate                = Get-ValueOrNull -Object $Agreement -PropertyName "startDate"
        endDate                  = Get-ValueOrNull -Object $Agreement -PropertyName "endDate"
        noEndingDateFlag         = Get-ValueOrNull -Object $Agreement -PropertyName "noEndingDateFlag"
        cancelledFlag            = Get-ValueOrNull -Object $Agreement -PropertyName "cancelledFlag"
        dateCancelled            = Get-ValueOrNull -Object $Agreement -PropertyName "dateCancelled"
        reasonCancelled          = Get-ValueOrNull -Object $Agreement -PropertyName "reasonCancelled"

        internalNotes            = $internalNotes
        hasInternalNotes         = (
            -not [string]::IsNullOrWhiteSpace([string]$internalNotes)
        )

        customerPO               = Get-ValueOrNull -Object $Agreement -PropertyName "customerPO"
        externalReference        = Get-ValueOrNull -Object $Agreement -PropertyName "externalReference"
        businessUnitId           = Get-ValueOrNull -Object $Agreement -PropertyName "businessUnitId"
        locationId               = Get-ValueOrNull -Object $Agreement -PropertyName "locationId"
        departmentId             = Get-ValueOrNull -Object $Agreement -PropertyName "departmentId"
        restrictLocationFlag     = Get-ValueOrNull -Object $Agreement -PropertyName "restrictLocationFlag"
        restrictDepartmentFlag   = Get-ValueOrNull -Object $Agreement -PropertyName "restrictDepartmentFlag"

        contactId                = Get-ValueOrNull -Object $contact -PropertyName "id"
        contactName              = Get-ValueOrNull -Object $contact -PropertyName "name"

        billToCompanyId          = Get-ValueOrNull -Object $billToCompany -PropertyName "id"
        billToCompanyName        = Get-ValueOrNull -Object $billToCompany -PropertyName "name"

        billAmount               = Get-ValueOrNull -Object $Agreement -PropertyName "billAmount"
        billingCycle             = Get-ValueOrNull -Object $Agreement -PropertyName "billCycleId"
        invoicingCycle           = Get-ValueOrNull -Object $Agreement -PropertyName "invoicingCycle"
        billOneTimeFlag          = Get-ValueOrNull -Object $Agreement -PropertyName "billOneTimeFlag"
        billStartDate            = Get-ValueOrNull -Object $Agreement -PropertyName "billStartDate"

        billTime                 = Get-ValueOrNull -Object $Agreement -PropertyName "billTime"
        billExpenses             = Get-ValueOrNull -Object $Agreement -PropertyName "billExpenses"
        billProducts             = Get-ValueOrNull -Object $Agreement -PropertyName "billProducts"

        applicationUnits         = Get-ValueOrNull -Object $Agreement -PropertyName "applicationUnits"
        applicationLimit         = Get-ValueOrNull -Object $Agreement -PropertyName "applicationLimit"
        applicationCycle         = Get-ValueOrNull -Object $Agreement -PropertyName "applicationCycle"
        applicationUnlimitedFlag = Get-ValueOrNull -Object $Agreement -PropertyName "applicationUnlimitedFlag"
        allowOverruns            = Get-ValueOrNull -Object $Agreement -PropertyName "allowOverruns"

        slaId                    = Get-ValueOrNull -Object $Agreement -PropertyName "slaId"
        parentAgreementId        = Get-ValueOrNull -Object $Agreement -PropertyName "parentAgreementId"
        lastUpdated              = Get-ValueOrNull `
            -Object (Get-ValueOrNull -Object $Agreement -PropertyName "_info") `
            -PropertyName "lastUpdated"
    }
}

# -------------------------------
# Main
# -------------------------------

try {
    Write-Info "Starting ConnectWise active agreement retrieval."

    $body = Get-RequestBodyObject -Request $Request

    if ($body -is [System.Collections.IDictionary]) {
        $body = [PSCustomObject]$body
    }

    if ($null -ne $Request.Body) {
        Write-DebugLog (
            "BODY TYPE: {0}" -f
            $Request.Body.GetType().FullName
        )

        Write-DebugLog "BODY CONTENT: [omitted]"
    }
    else {
        Write-DebugLog "BODY: null"
    }

    if ($Request.Query) {
        Write-DebugLog (
            "QUERY KEYS: {0}" -f
            (@($Request.Query.Keys) -join ", ")
        )
    }

    $CompanyName = Get-InputValue `
        -Request $Request `
        -Body $body `
        -Names @(
            "CompanyName",
            "companyName",
            "Company",
            "company"
        )

    $CompanyIdInput = Get-InputValue `
        -Request $Request `
        -Body $body `
        -Names @(
            "CompanyId",
            "companyId"
        )

    $RequestSecurityKey = Get-InputValue `
        -Request $Request `
        -Body $body `
        -Names @(
            "SecurityKey",
            "securityKey"
        )

    $hasCompanyId = -not [string]::IsNullOrWhiteSpace([string]$CompanyIdInput)
    $hasCompanyName = -not [string]::IsNullOrWhiteSpace([string]$CompanyName)

    if (-not $hasCompanyId -and -not $hasCompanyName) {
        Send-JsonResponse `
            -StatusCode ([HttpStatusCode]::BadRequest) `
            -Body @{
                success       = $false
                correlationId = $CorrelationId
                action        = "GetCompanyActiveAgreements"
                errorCode     = "MissingCompanyIdentifier"
                error         = "Provide CompanyId or CompanyName in the JSON body, query string, or header."
            }

        return
    }

    $requestedCompanyId = 0
    if ($hasCompanyId -and (-not [int]::TryParse(([string]$CompanyIdInput).Trim(), [ref]$requestedCompanyId) -or $requestedCompanyId -lt 1)) {
        Send-JsonResponse `
            -StatusCode ([HttpStatusCode]::BadRequest) `
            -Body @{
                success       = $false
                correlationId = $CorrelationId
                action        = "GetCompanyActiveAgreements"
                errorCode     = "InvalidCompanyId"
                error         = "CompanyId must be a positive integer."
            }

        return
    }

    $CompanyName = if ($hasCompanyName) { ([string]$CompanyName).Trim() } else { $null }
    $CompanyInput = if ($hasCompanyId) { [string]$requestedCompanyId } else { $CompanyName }

    # -------------------------------
    # Security-key validation
    # -------------------------------

    $ConfiguredSecurityKey = $env:SecurityKey

    if (
        -not [string]::IsNullOrWhiteSpace($ConfiguredSecurityKey) -and
        -not [string]::Equals(
            $ConfiguredSecurityKey,
            $RequestSecurityKey,
            [System.StringComparison]::Ordinal
        )
    ) {
        Send-JsonResponse `
            -StatusCode ([HttpStatusCode]::Unauthorized) `
            -Body @{
                success       = $false
                correlationId = $CorrelationId
                action        = "GetCompanyActiveAgreements"
                errorCode     = "InvalidSecurityKey"
                error         = "Invalid security key."
            }

        return
    }

    # -------------------------------
    # Environment validation
    # -------------------------------

    $requiredSettings = @(
        "ConnectWisePsa_ApiBaseUrl",
        "ConnectWisePsa_ApiCompanyId",
        "ConnectWisePsa_ApiPublicKey",
        "ConnectWisePsa_ApiPrivateKey",
        "ConnectWisePsa_ApiClientId"
    )

    $missingSettings = @()

    foreach ($setting in $requiredSettings) {
        $settingValue = Get-Item `
            -Path "Env:$setting" `
            -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Value

        if (
            [string]::IsNullOrWhiteSpace(
                [string]$settingValue
            )
        ) {
            $missingSettings += $setting
        }
    }

    if ($missingSettings.Count -gt 0) {
        Send-JsonResponse `
            -StatusCode ([HttpStatusCode]::InternalServerError) `
            -Body @{
                success         = $false
                correlationId   = $CorrelationId
                action          = "GetCompanyActiveAgreements"
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

    # -------------------------------
    # Find company
    # -------------------------------

    $companyId = 0
    if ($hasCompanyId) {
        $companyId = $requestedCompanyId
    }
    else {
        Write-Info (
            "Searching for ConnectWise company '{0}'." -f
            $CompanyName
        )

        $companyCandidates = @(
            Find-ConnectWiseCompanies `
                -ConnectWiseUrl $ConnectWiseUrl `
                -CompanyName $CompanyName `
                -Headers $headers
        )

        Write-DebugLog (
            "Company candidate count: {0}" -f
            $companyCandidates.Count
        )

        foreach ($candidate in $companyCandidates) {
            Write-DebugLog (
                "Company candidate: ID={0}; Name={1}; Type={2}" -f
                $candidate.id,
                $candidate.name,
                $candidate.GetType().FullName
            )
        }

        if ($companyCandidates.Count -eq 0) {
            Send-JsonResponse `
                -StatusCode ([HttpStatusCode]::NotFound) `
                -Body @{
                    success        = $false
                    correlationId  = $CorrelationId
                    action         = "GetCompanyActiveAgreements"
                    companyInput   = $CompanyInput
                    errorCode      = "CompanyNotFound"
                    error          = "No ConnectWise company matched the supplied company name."
                    retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
                }

            return
        }

        # Prefer an exact case-insensitive name match.
        $exactMatches = @(
            $companyCandidates |
                Where-Object {
                    $_.name -and
                    $_.name.ToString().Trim().Equals(
                        $CompanyName,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
        )

        if ($exactMatches.Count -eq 1) {
            $selectedCompanySummary = $exactMatches[0]
        }
        elseif ($exactMatches.Count -gt 1) {
            Send-JsonResponse `
                -StatusCode ([HttpStatusCode]::Conflict) `
                -Body @{
                    success        = $false
                    correlationId  = $CorrelationId
                    action         = "GetCompanyActiveAgreements"
                    companyInput   = $CompanyInput
                    errorCode      = "MultipleExactCompanyMatches"
                    error          = "Multiple ConnectWise companies have the exact supplied name."
                    matches        = @(
                        $exactMatches | ForEach-Object {
                            @{
                                id         = $_.id
                                identifier = $_.identifier
                                name       = $_.name
                            }
                        }
                    )
                    retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
                }

            return
        }
        elseif ($companyCandidates.Count -eq 1) {
            $selectedCompanySummary = $companyCandidates[0]
        }
        else {
            Send-JsonResponse `
                -StatusCode ([HttpStatusCode]::Conflict) `
                -Body @{
                    success        = $false
                    correlationId  = $CorrelationId
                    action         = "GetCompanyActiveAgreements"
                    companyInput   = $CompanyInput
                    errorCode      = "AmbiguousCompanyName"
                    error          = "The supplied company name matched multiple ConnectWise companies. Provide the exact company name or a CompanyId."
                    matches        = @(
                        $companyCandidates | ForEach-Object {
                            @{
                                id         = $_.id
                                identifier = $_.identifier
                                name       = $_.name
                            }
                        }
                    )
                    retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
                }

            return
        }

        if (-not [int]::TryParse([string]$selectedCompanySummary.id, [ref]$companyId) -or $companyId -lt 1) {
            throw [System.InvalidOperationException]::new("ConnectWise returned an invalid company ID.")
        }
    }

    try {
        $selectedCompany = Get-ConnectWiseCompany `
            -ConnectWiseUrl $ConnectWiseUrl `
            -CompanyId $companyId `
            -Headers $headers
    }
    catch {
        if ($hasCompanyId -and $_.Exception.Message -match "404|Not Found") {
            Send-JsonResponse `
                -StatusCode ([HttpStatusCode]::NotFound) `
                -Body @{
                    success        = $false
                    correlationId  = $CorrelationId
                    action         = "GetCompanyActiveAgreements"
                    companyInput   = $CompanyInput
                    errorCode      = "CompanyNotFound"
                    error          = "No ConnectWise company exists with the supplied CompanyId."
                    retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
                }

            return
        }

        throw
    }

    $companyTypeNames = @(
        Get-CompanyTypeNames -Company $selectedCompany
    )

    $isClient = Test-CompanyIsClient -Company $selectedCompany

    if (-not $isClient) {
        Send-JsonResponse `
            -StatusCode ([HttpStatusCode]::Forbidden) `
            -Body @{
                success       = $false
                correlationId = $CorrelationId
                action        = "GetCompanyActiveAgreements"
                companyInput  = $CompanyInput
                company       = @{
                    id          = $selectedCompany.id
                    identifier  = $selectedCompany.identifier
                    name        = $selectedCompany.name
                    companyTypes = $companyTypeNames
                }
                errorCode     = "CompanyIsNotClient"
                error         = "The matching ConnectWise company does not have the company type Client."
            }

        return
    }

    # -------------------------------
    # Retrieve and filter agreements
    # -------------------------------

    $pageSize = 100
    $parsedPageSize = 0

    if (
        $env:ConnectWisePsa_AgreementsPageSize -and
        [int]::TryParse(
            [string]$env:ConnectWisePsa_AgreementsPageSize,
            [ref]$parsedPageSize
        ) -and
        $parsedPageSize -gt 0
    ) {
        $pageSize = $parsedPageSize
    }

    if ($pageSize -gt 1000) {
        $pageSize = 1000
    }

    Write-Info (
        "Retrieving agreements for company {0}, ID {1}." -f
        $selectedCompany.name,
        $selectedCompany.id
    )

    $allAgreements = @(
        Get-ConnectWiseCompanyAgreements `
            -ConnectWiseUrl $ConnectWiseUrl `
            -CompanyId ([int]$selectedCompany.id) `
            -Headers $headers `
            -PageSize $pageSize
    )

    $asOfUtc = (Get-Date).ToUniversalTime()

    $activeAgreements = @(
        $allAgreements |
            Where-Object {
                Test-AgreementIsActive `
                    -Agreement $_ `
                    -AsOfUtc $asOfUtc
            }
    )

    $normalizedAgreements = @(
        $activeAgreements |
            ForEach-Object {
                ConvertTo-AgreementSummary -Agreement $_
            }
    )

    Write-Info (
        "Found {0} active agreement(s) from {1} total agreement record(s)." -f
        $activeAgreements.Count,
        $allAgreements.Count
    )

    # The request expects at least one active agreement.
    if ($activeAgreements.Count -eq 0) {
        Send-JsonResponse `
            -StatusCode ([HttpStatusCode]::NotFound) `
            -Body @{
                success       = $false
                correlationId = $CorrelationId
                action        = "GetCompanyActiveAgreements"
                companyInput  = $CompanyInput
                company       = @{
                    id           = $selectedCompany.id
                    identifier   = $selectedCompany.identifier
                    name         = $selectedCompany.name
                    companyTypes = $companyTypeNames
                    isClient     = $true
                }
                totalAgreementCount  = $allAgreements.Count
                activeAgreementCount = 0
                errorCode            = "NoActiveAgreements"
                error                = "The matching Client company does not have any active agreements as of the current UTC date."
                asOfUtc              = $asOfUtc.ToString("o")
            }

        return
    }

    # -------------------------------
    # Successful response
    # -------------------------------

    $response = [ordered]@{
        success       = $true
        correlationId = $CorrelationId
        action        = "GetCompanyActiveAgreements"
        companyInput  = $CompanyInput
        asOfUtc       = $asOfUtc.ToString("o")
        retrievedAtUtc = (Get-Date).ToUniversalTime().ToString("o")

        company = [ordered]@{
            id           = $selectedCompany.id
            identifier   = $selectedCompany.identifier
            name         = $selectedCompany.name
            companyTypes = $companyTypeNames
            isClient     = $true
            status       = Get-ValueOrNull `
                -Object $selectedCompany `
                -PropertyName "status"
        }

        agreementSummary = [ordered]@{
            totalAgreementCount  = $allAgreements.Count
            activeAgreementCount = $activeAgreements.Count
            activeDefinition     = @(
                "cancelledFlag is not true",
                "startDate is blank or on/before the current UTC date",
                "noEndingDateFlag is true, or endDate is blank or on/after the current UTC date"
            )
        }

        agreements = $normalizedAgreements

        raw = [ordered]@{
            company          = $selectedCompany
            activeAgreements = $activeAgreements
        }
    }

    Send-JsonResponse `
        -StatusCode ([HttpStatusCode]::OK) `
        -Body $response
}
catch {
    $errorMessage = $_.Exception.Message

    Write-ErrorLog $errorMessage

    Send-JsonResponse `
        -StatusCode ([HttpStatusCode]::InternalServerError) `
        -Body @{
            success       = $false
            correlationId = $CorrelationId
            action        = "GetCompanyActiveAgreements"
            errorCode     = "UnhandledFunctionError"
            error         = "An unexpected error occurred while retrieving agreements."
        }
}