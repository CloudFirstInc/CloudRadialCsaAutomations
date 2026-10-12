<#
.SYNOPSIS
    Retrieves ConnectWise Manage projects for a Client company.

.DESCRIPTION
    This read-only Azure Function accepts a company name, resolves the exact
    ConnectWise company, validates the company type Client, retrieves all
    projects belonging to the company, and returns open projects by default.

    Set IncludeClosed to true to return every project. Open-project filtering is
    performed locally. Terminal project status names can be configured with
    ConnectWisePsa_ClosedProjectStatuses as a comma-separated list.

.REQUIRED ENVIRONMENT VARIABLES
    ConnectWisePsa_ApiBaseUrl
    ConnectWisePsa_ApiCompanyId
    ConnectWisePsa_ApiPublicKey
    ConnectWisePsa_ApiPrivateKey
    ConnectWisePsa_ApiClientId
    SecurityKey                                - Optional shared secret

.OPTIONAL ENVIRONMENT VARIABLES
    ConnectWisePsa_ProjectsPageSize             - Default 100; maximum 1000
    ConnectWisePsa_ClosedProjectStatuses        - Comma-separated terminal names
                                                  Default: Closed,Completed,Complete,
                                                  Cancelled,Canceled
    DebugLogging                                - Set to 1 for diagnostic logs

.INPUTS
    JSON body:
    {
        "CompanyName": "City of Rye",
        "IncludeClosed": false,
        "SecurityKey": "optional"
    }

    Query-string example:
    ?CompanyName=City%20of%20Rye&IncludeClosed=false

.OUTPUTS
    JSON containing the selected company, project counts, normalized project
    summaries, and complete raw project records.
#>

using namespace System.Net
param($Request, $TriggerMetadata)

$CorrelationId = [guid]::NewGuid().ToString()
$DebugEnabled = ($env:DebugLogging -eq "1")

function Write-Info { param([string]$Message) Write-Host ("[{0}] [INFO] {1}" -f $CorrelationId, $Message) }
function Write-DebugLog { param([string]$Message) if ($DebugEnabled) { Write-Host ("[{0}] [DEBUG] {1}" -f $CorrelationId, $Message) } }
function Write-ErrorLog { param([string]$Message) Write-Host ("[{0}] [ERROR] {1}" -f $CorrelationId, $Message) }

function Send-JsonResponse {
    param([System.Net.HttpStatusCode]$StatusCode, [object]$Body)
    $json = $Body | ConvertTo-Json -Depth 100
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode=$StatusCode; Body=$json; ContentType="application/json"
    })
}

function Get-ValueOrNull {
    param([object]$Object,[string]$PropertyName)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function New-ConnectWiseHeaders {
    param([string]$CompanyId,[string]$PublicKey,[string]$PrivateKey,[string]$ClientId)
    $authString = "{0}+{1}:{2}" -f $CompanyId,$PublicKey,$PrivateKey
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($authString))
    return @{
        Authorization = "Basic {0}" -f $encodedAuth
        "Content-Type" = "application/json"
        Accept = "application/json"
        ClientId = $ClientId
    }
}

function Invoke-ConnectWiseApi {
    param([string]$Uri,[string]$Method="Get",[hashtable]$Headers,[object]$Body=$null,[int]$MaxRetries=5)
    $attempt = 0
    while ($true) {
        try {
            $attempt++
            Write-DebugLog ("CW API {0} {1}, attempt {2}" -f $Method,$Uri,$attempt)
            if ($null -ne $Body) {
                return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers `
                    -Body ($Body | ConvertTo-Json -Depth 20) -ErrorAction Stop
            }
            return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -ErrorAction Stop
        }
        catch {
            $statusCode = $null; $retryAfter = $null
            if ($_.Exception.Response) {
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch {}
            }
            Write-ErrorLog ("CW API call failed. StatusCode={0}. Message={1}" -f $statusCode,$_.Exception.Message)
            if (($statusCode -notin @(408,429,500,502,503,504)) -or $attempt -ge $MaxRetries) { throw }
            $sleepSeconds = 0; $parsedRetryAfter = 0
            if ($retryAfter -and [int]::TryParse([string]$retryAfter,[ref]$parsedRetryAfter)) {
                $sleepSeconds = $parsedRetryAfter
            }
            if ($sleepSeconds -lt 1) {
                $sleepSeconds = [int][Math]::Pow(2,$attempt) + (Get-Random -Minimum 1 -Maximum 4)
            }
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
                        $value=$Body[$key]; if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
                    }
                }
            } else {
                $property = $Body.PSObject.Properties | Where-Object { $_.Name.Equals($name,[StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
                if ($property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return [string]$property.Value }
            }
        }
        foreach ($source in @($Request.Query,$Request.Headers)) {
            if ($source) {
                foreach ($key in $source.Keys) {
                    if ($null -ne $key -and $key.ToString().Equals($name,[StringComparison]::OrdinalIgnoreCase)) {
                        $value=$source[$key]; if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
                    }
                }
            }
        }
    }
    return $null
}

function ConvertTo-Boolean {
    param([object]$Value,[bool]$Default=$false)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $Default }
    $parsed=$false
    if ([bool]::TryParse([string]$Value,[ref]$parsed)) { return $parsed }
    if ([string]$Value -in @("1","yes","y")) { return $true }
    if ([string]$Value -in @("0","no","n")) { return $false }
    throw "IncludeClosed must be true or false."
}

function Find-ConnectWiseCompanies {
    param([string]$ConnectWiseUrl,[string]$CompanyName,[hashtable]$Headers)
    $safeCompanyName=ConvertTo-CwConditionString $CompanyName
    $conditions=ConvertTo-CwUrlEncoded ("name contains '{0}'" -f $safeCompanyName)
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

function Get-ConnectWiseCompanyProjects {
    param([string]$ConnectWiseUrl,[int]$CompanyId,[hashtable]$Headers,[int]$PageSize=100)
    if ($PageSize -lt 1) { $PageSize=100 }; if ($PageSize -gt 1000) { $PageSize=1000 }
    $projects=@(); $page=1
    while ($true) {
        $conditions=ConvertTo-CwUrlEncoded ("company/id={0}" -f $CompanyId)
        $uri=("{0}/v4_6_release/apis/3.0/project/projects?conditions={1}&orderBy=id&page={2}&pageSize={3}" -f $ConnectWiseUrl.TrimEnd('/'),$conditions,$page,$PageSize)
        Write-DebugLog ("Retrieving projects page {0} for company ID {1}." -f $page,$CompanyId)
        $items=@(Invoke-ConnectWiseApi -Uri $uri -Headers $Headers | ForEach-Object { $_ })
        if ($items.Count -eq 0) { break }
        $projects += $items
        if ($items.Count -lt $PageSize) { break }
        $page++
    }
    return @($projects)
}

function Get-TerminalProjectStatuses {
    $configured=[string]$env:ConnectWisePsa_ClosedProjectStatuses
    if ([string]::IsNullOrWhiteSpace($configured)) { $configured="Closed,Completed,Complete,Cancelled,Canceled" }
    return @($configured.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Test-ProjectIsOpen {
    param([object]$Project,[string[]]$TerminalStatuses)
    if ((Get-ValueOrNull $Project "closedFlag") -eq $true) { return $false }
    $status=Get-ValueOrNull $Project "status"
    $statusName=[string](Get-ValueOrNull $status "name")
    if ([string]::IsNullOrWhiteSpace($statusName)) { return $true }
    foreach ($terminal in $TerminalStatuses) {
        if ($statusName.Equals($terminal,[StringComparison]::OrdinalIgnoreCase)) { return $false }
    }
    return $true
}

function ConvertTo-ProjectSummary {
    param([object]$Project,[bool]$IsOpen)
    $company=Get-ValueOrNull $Project "company"
    $status=Get-ValueOrNull $Project "status"
    $type=Get-ValueOrNull $Project "type"
    $manager=Get-ValueOrNull $Project "manager"
    $board=Get-ValueOrNull $Project "board"
    $department=Get-ValueOrNull $Project "department"
    $location=Get-ValueOrNull $Project "location"
    $contact=Get-ValueOrNull $Project "contact"
    $billingMethod=Get-ValueOrNull $Project "billingMethod"
    $description=Get-ValueOrNull $Project "description"
    $info=Get-ValueOrNull $Project "_info"
    return [ordered]@{
        id=Get-ValueOrNull $Project "id"
        name=Get-ValueOrNull $Project "name"
        isOpen=$IsOpen
        companyId=Get-ValueOrNull $company "id"
        companyName=Get-ValueOrNull $company "name"
        statusId=Get-ValueOrNull $status "id"
        statusName=Get-ValueOrNull $status "name"
        typeId=Get-ValueOrNull $type "id"
        typeName=Get-ValueOrNull $type "name"
        boardId=Get-ValueOrNull $board "id"
        boardName=Get-ValueOrNull $board "name"
        projectManagerId=Get-ValueOrNull $manager "id"
        projectManagerName=Get-ValueOrNull $manager "name"
        contactId=Get-ValueOrNull $contact "id"
        contactName=Get-ValueOrNull $contact "name"
        departmentId=Get-ValueOrNull $department "id"
        departmentName=Get-ValueOrNull $department "name"
        locationId=Get-ValueOrNull $location "id"
        locationName=Get-ValueOrNull $location "name"
        startDate=Get-ValueOrNull $Project "startDate"
        endDate=Get-ValueOrNull $Project "endDate"
        estimatedStart=Get-ValueOrNull $Project "estimatedStart"
        estimatedEnd=Get-ValueOrNull $Project "estimatedEnd"
        actualStart=Get-ValueOrNull $Project "actualStart"
        actualEnd=Get-ValueOrNull $Project "actualEnd"
        percentComplete=Get-ValueOrNull $Project "percentComplete"
        description=$description
        scope=$description
        customerPO=Get-ValueOrNull $Project "customerPO"
        externalReference=Get-ValueOrNull $Project "externalReference"
        billingMethodId=Get-ValueOrNull $billingMethod "id"
        billingMethodName=Get-ValueOrNull $billingMethod "name"
        billTime=Get-ValueOrNull $Project "billTime"
        billExpenses=Get-ValueOrNull $Project "billExpenses"
        billProducts=Get-ValueOrNull $Project "billProducts"
        budgetHours=Get-ValueOrNull $Project "budgetHours"
        estimatedRevenue=Get-ValueOrNull $Project "estimatedRevenue"
        lastUpdated=Get-ValueOrNull $info "lastUpdated"
    }
}

try {
    Write-Info "Starting ConnectWise project retrieval."
    $body=Get-RequestBodyObject $Request
    if ($body -is [Collections.IDictionary]) { $body=[PSCustomObject]$body }

    $CompanyName=Get-InputValue $Request $body @("CompanyName","companyName","Company","company")
    $RequestSecurityKey=Get-InputValue $Request $body @("SecurityKey","securityKey")
    $IncludeClosedValue=Get-InputValue $Request $body @("IncludeClosed","includeClosed")

    if ([string]::IsNullOrWhiteSpace($CompanyName)) {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{success=$false;correlationId=$CorrelationId;action="GetCompanyProjects";errorCode="MissingCompanyName";error="Provide CompanyName in the JSON body, query string, or header."}; return
    }
    $CompanyName=$CompanyName.Trim()
    try { $IncludeClosed=ConvertTo-Boolean $IncludeClosedValue $false }
    catch {
        Send-JsonResponse ([HttpStatusCode]::BadRequest) @{success=$false;correlationId=$CorrelationId;action="GetCompanyProjects";errorCode="InvalidIncludeClosed";error="IncludeClosed must be true or false."}; return
    }

    $configuredKey=$env:SecurityKey
    if (-not [string]::IsNullOrWhiteSpace($configuredKey) -and -not [string]::Equals($configuredKey,$RequestSecurityKey,[StringComparison]::Ordinal)) {
        Send-JsonResponse ([HttpStatusCode]::Unauthorized) @{success=$false;correlationId=$CorrelationId;action="GetCompanyProjects";errorCode="InvalidSecurityKey";error="Invalid security key."}; return
    }

    $required=@("ConnectWisePsa_ApiBaseUrl","ConnectWisePsa_ApiCompanyId","ConnectWisePsa_ApiPublicKey","ConnectWisePsa_ApiPrivateKey","ConnectWisePsa_ApiClientId")
    $missing=@()
    foreach ($setting in $required) {
        $value=Get-Item -Path "Env:$setting" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Value
        if ([string]::IsNullOrWhiteSpace([string]$value)) { $missing += $setting }
    }
    if ($missing.Count -gt 0) {
        Send-JsonResponse ([HttpStatusCode]::InternalServerError) @{success=$false;correlationId=$CorrelationId;action="GetCompanyProjects";errorCode="MissingApplicationSettings";error="Missing required application settings.";missingSettings=$missing}; return
    }

    $cwUrl=$env:ConnectWisePsa_ApiBaseUrl.TrimEnd('/')
    $headers=New-ConnectWiseHeaders $env:ConnectWisePsa_ApiCompanyId $env:ConnectWisePsa_ApiPublicKey $env:ConnectWisePsa_ApiPrivateKey $env:ConnectWisePsa_ApiClientId
    $candidates=@(Find-ConnectWiseCompanies $cwUrl $CompanyName $headers)
    if ($candidates.Count -eq 0) {
        Send-JsonResponse ([HttpStatusCode]::NotFound) @{success=$false;correlationId=$CorrelationId;action="GetCompanyProjects";companyInput=$CompanyName;errorCode="CompanyNotFound";error="No ConnectWise company matched the supplied company name."}; return
    }

    $exact=@($candidates | Where-Object { $_.name -and $_.name.ToString().Trim().Equals($CompanyName,[StringComparison]::OrdinalIgnoreCase) })
    if ($exact.Count -eq 1) { $selectedSummary=$exact[0] }
    elseif ($exact.Count -gt 1 -or $candidates.Count -gt 1) {
        $matchedCompanies=if ($exact.Count -gt 1) {$exact} else {$candidates}
        Send-JsonResponse ([HttpStatusCode]::Conflict) @{success=$false;correlationId=$CorrelationId;action="GetCompanyProjects";companyInput=$CompanyName;errorCode="AmbiguousCompanyName";error="Multiple ConnectWise companies matched. Provide the exact company name.";matches=@($matchedCompanies | ForEach-Object {@{id=$_.id;identifier=$_.identifier;name=$_.name}})}; return
    } else { $selectedSummary=$candidates[0] }

    $company=Get-ConnectWiseCompany $cwUrl ([int]$selectedSummary.id) $headers
    $types=@(Get-CompanyTypeNames $company)
    if (-not (Test-CompanyIsClient $company)) {
        Send-JsonResponse ([HttpStatusCode]::Forbidden) @{success=$false;correlationId=$CorrelationId;action="GetCompanyProjects";companyInput=$CompanyName;company=@{id=$company.id;identifier=$company.identifier;name=$company.name;companyTypes=$types};errorCode="CompanyIsNotClient";error="The matching ConnectWise company does not have the company type Client."}; return
    }

    $pageSize=100; $parsed=0
    if ($env:ConnectWisePsa_ProjectsPageSize -and [int]::TryParse([string]$env:ConnectWisePsa_ProjectsPageSize,[ref]$parsed) -and $parsed -gt 0) { $pageSize=$parsed }
    if ($pageSize -gt 1000) { $pageSize=1000 }

    $all=@(Get-ConnectWiseCompanyProjects $cwUrl ([int]$company.id) $headers $pageSize)
    $terminal=@(Get-TerminalProjectStatuses)
    $open=@($all | Where-Object { Test-ProjectIsOpen $_ $terminal })
    $returned=if ($IncludeClosed) {@($all)} else {@($open)}
    $normalized=@($returned | ForEach-Object { ConvertTo-ProjectSummary $_ (Test-ProjectIsOpen $_ $terminal) })

    $response=[ordered]@{
        success=$true
        correlationId=$CorrelationId
        action="GetCompanyProjects"
        companyInput=$CompanyName
        includeClosed=$IncludeClosed
        retrievedAtUtc=(Get-Date).ToUniversalTime().ToString('o')
        company=[ordered]@{id=$company.id;identifier=$company.identifier;name=$company.name;companyTypes=$types;isClient=$true;status=Get-ValueOrNull $company "status"}
        projectSummary=[ordered]@{totalProjectCount=$all.Count;openProjectCount=$open.Count;returnedProjectCount=$returned.Count;terminalStatusNames=$terminal;openDefinition="A project is open when closedFlag is not true and its status name is not in terminalStatusNames."}
        projects=$normalized
        raw=[ordered]@{company=$company;projects=$returned}
    }
    Send-JsonResponse ([HttpStatusCode]::OK) $response
}
catch {
    Write-ErrorLog $_.Exception.Message
    Send-JsonResponse ([HttpStatusCode]::InternalServerError) @{success=$false;correlationId=$CorrelationId;action="GetCompanyProjects";errorCode="UnhandledFunctionError";error="An unexpected error occurred while retrieving projects."}
}
