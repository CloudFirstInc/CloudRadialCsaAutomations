
# Ms365-GetUserDepartment/run.ps1
# CloudRadial API tenant lookup (by PSA Company Id) + CW dept UDF update
# Parser-safe: uses -f format strings and avoids colon-adjacent interpolation.

using namespace System.Net

param($Request, $TriggerMetadata)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------
# Correlation & Logging
# ---------------------------
$CorrelationId = [Guid]::NewGuid().ToString()
$IsDebug       = ([Environment]::GetEnvironmentVariable('DebugLogging') -as [int]) -eq 1

function LogInfo  { param([string]$msg) Write-Information ("[{0}] {1}" -f $CorrelationId, $msg) }
function LogError { param([string]$msg) Write-Error       -Message ("[{0}] {1}" -f $CorrelationId, $msg) -ErrorAction Continue }
function LogDebug { param([string]$msg) if ($IsDebug) { Write-Information ("[{0}][DEBUG] {1}" -f $CorrelationId, $msg) } }

# --- Simple helper to dump keys & values from NameValueCollection-like inputs ---
function Dump-Collection {
    param(
        [Parameter(Mandatory=$true)][string]$name,
        [Parameter()][object]$coll
    )
    try {
        if (-not $coll) { Write-Information ("[{0}] {1}: <null>" -f $CorrelationId, $name); return }
        $keys = @()
        if ($coll.PSObject.Properties['AllKeys']) { $keys = $coll.AllKeys }
        elseif ($coll.PSObject.Properties['Keys']) { $keys = $coll.Keys }
        elseif ($coll -is [System.Collections.IDictionary]) { $keys = @($coll.Keys) }
        else { try { foreach ($k in $coll) { $keys += $k } } catch { $keys = @() } }
        Write-Information ("[{0}] {1} keys: {2}" -f $CorrelationId, $name, (($keys | Where-Object { $_ }) -join ', '))
        if ($IsDebug) {
            foreach ($k in $keys) {
                try { $val = $coll[$k]; Write-Information ("[{0}][DEBUG] {1}[{2}] = '{3}'" -f $CorrelationId, $name, $k, (""+$val)) }
                catch { Write-Information ("[{0}][DEBUG] {1}[{2}] = <error: {3}>" -f $CorrelationId, $name, $k, $_.Exception.Message) }
            }
        }
    } catch { Write-Error ("[{0}] Failed to dump {1}: {2}" -f $CorrelationId, $name, $_.Exception.Message) }
}

# Diagnostics
Dump-Collection -name 'Query'   -coll $Request.Query
Dump-Collection -name 'Headers' -coll $Request.Headers

# ---------------------------
# HTTP JSON Response
# ---------------------------
function New-JsonResponse {
    param([int]$Code,[string]$Message,[hashtable]$Extra = @{ })
    $body = @{
        Message       = $Message
        ResultCode    = $Code
        ResultStatus  = if ($Code -ge 200 -and $Code -lt 300) { "Success" } else { "Failure" }
        CorrelationId = $CorrelationId
    } + $Extra
    try {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]$Code
            Body       = $body
            Headers    = @{ "Content-Type" = "application/json" }
        })
    } catch { Write-Output ($body | ConvertTo-Json -Depth 6) }
}

# ---------------------------
# Request Body Reader
# ---------------------------
function Get-RequestBodyObject {
    param([object]$Request)
    if (-not $Request) { return @{} }
    $raw = $Request.Body
    if ($null -eq $raw) { return @{} }
    if ($raw -is [System.IO.Stream]) {
        try {
            $reader = New-Object System.IO.StreamReader($raw)
            $text   = $reader.ReadToEnd()
            if (-not [string]::IsNullOrWhiteSpace($text) -and $text.Trim().StartsWith('{')) {
                return $text | ConvertFrom-Json -ErrorAction Stop
            }
            return @{}
        } catch { return @{} }
    }
    if ($raw -is [string]) {
        try { if ($raw.Trim().StartsWith('{')) { return $raw | ConvertFrom-Json -ErrorAction Stop } } catch { }
        return @{}
    }
    return $raw
}

# ---------------------------
# Robust integer parsing helper
# ---------------------------
function To-IntOrNull {
    param([object]$value)
    if ($null -eq $value) { return $null }
    $s = "" + $value
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $out = 0
    if ([int]::TryParse($s, [ref]$out)) { return $out }
    return $null
}

# Probe mode
if ($Request -and $Request.PSObject.Properties['Query'] -and $Request.Query['probe'] -eq '1') {
    New-JsonResponse -Code 200 -Message "Probe OK" -Extra @{ QueryKeys=$Request.Query.AllKeys; HeaderKeys=$Request.Headers.Keys }; return
}

# ---------------------------
# Microsoft Graph helpers
# ---------------------------
function Connect-GraphApp {
    param([string]$TenantId)
    try {
        $appId     = [Environment]::GetEnvironmentVariable('Ms365_AuthAppId')
        $appSecret = [Environment]::GetEnvironmentVariable('Ms365_AuthSecretId')
        if ([string]::IsNullOrWhiteSpace($appId))     { throw "Missing Microsoft Graph app setting: Ms365_AuthAppId" }
        if ([string]::IsNullOrWhiteSpace($appSecret)) { throw "Missing Microsoft Graph app setting: Ms365_AuthSecretId (client secret VALUE)" }
        if ([string]::IsNullOrWhiteSpace($TenantId))  { throw "Missing Microsoft Graph TenantId" }
        if ($appSecret -match '^[0-9a-fA-F-]{36}$') { LogError "Ms365_AuthSecretId appears to be a GUID (secret ID). Store the secret VALUE, not the ID." }
        LogInfo ("Connecting to Graph (TenantId={0}, AppId={1})" -f $TenantId, $appId)
        $secureSecret           = ConvertTo-SecureString $appSecret -AsPlainText -Force
        $clientSecretCredential = New-Object System.Management.Automation.PSCredential($appId, $secureSecret)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $clientSecretCredential -NoWelcome -ErrorAction Stop
        return $true
    } catch { LogError ("Graph connect failed: " + $_.Exception.Message); return $false }
}

function Get-UserDepartment {
    param([string]$UserPrincipalName,[string]$UserEmail)
    $user = $null
    if ($UserEmail) {
        try {
            $safeEmail = $UserEmail.Replace("'","''")
            $filter    = "mail eq '{0}' or userPrincipalName eq '{0}'" -f $safeEmail
            LogInfo ("Graph lookup by filter: {0}" -f $filter)
            $user      = Get-MgUser -Filter $filter -Property department,userPrincipalName -Top 1 -ErrorAction Stop
        } catch { LogError ("Get-MgUser by email filter failed: " + $_.Exception.Message) }
    }
    if (-not $user -and $UserPrincipalName) {
        try { LogInfo ("Graph lookup by Id/UPN: {0}" -f $UserPrincipalName); $user = Get-MgUser -UserId $UserPrincipalName -Property department,userPrincipalName -ErrorAction Stop }
        catch { LogError ("Get-MgUser by Id/UPN failed: " + $_.Exception.Message) }
    }
    if ($user) { LogInfo ("Graph user department = '{0}'" -f (""+$user.Department)); return $user.Department }
    LogInfo ("Graph user not found or no department."); return $null
}

# ---------------------------
# ConnectWise helpers
# ---------------------------
$CwServer     = 'api-na.myconnectwise.net'
$UseJsonPatch = ([Environment]::GetEnvironmentVariable('ConnectWise_UseJsonPatch') -as [int]) -eq 1

function Get-CwHeaders {
    param([string]$ContentType = 'application/json')
    $required = 'ConnectWisePsa_ApiCompanyId','ConnectWisePsa_ApiPublicKey','ConnectWisePsa_ApiPrivateKey','ConnectWisePsa_ApiClientId'
    foreach ($n in $required) { if ([string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($n))) { throw "Missing ConnectWise app setting: $n" } }
    $companyId = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiCompanyId')
    $pubKey    = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiPublicKey')
    $privKey   = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiPrivateKey')
    $clientId  = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiClientId')
    $authString  = "{0}+{1}:{2}" -f $companyId, $pubKey, $privKey
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
    return @{
        "Authorization" = "Basic $encodedAuth"
        "clientId"      = $clientId
        "Content-Type"  = $ContentType
        "Accept"        = "application/vnd.connectwise.com+json; version=2022.1"
    }
}

function Get-CwErrorBody {
    param($ex)
    if ($ex.PSObject.Properties['Response'] -and $ex.Response) {
        try { if ($ex.Response.PSObject.Properties['Content'] -and $ex.Response.Content) { return $ex.Response.Content.ReadAsStringAsync().Result } } catch {}
        try {
            $stream = $ex.Response.GetResponseStream()
            if ($stream) { $reader = New-Object System.IO.StreamReader($stream); return $reader.ReadToEnd() }
        } catch {}
    }
    return $null
}

function Get-CwTicket {
    param([int]$TicketId)
    $headers = Get-CwHeaders -ContentType 'application/json'
    $url     = 'https://{0}/v4_6_release/apis/3.0/service/tickets/{1}' -f $CwServer, $TicketId
    LogInfo ("CW GET ticket: id={0}, url={1}" -f $TicketId, $url)
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
        LogInfo ("CW GET ticket -> OK")
        return $resp
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW GET ticket failed: " + $_.Exception.Message + $suffix)
        return $null
    }
}

function Get-CwContactById {
    param([int]$ContactId)
    $headers = Get-CwHeaders -ContentType 'application/json'
    $url     = 'https://{0}/v4_6_release/apis/3.0/company/contacts/{1}' -f $CwServer, $ContactId
    LogInfo ("CW GET contact: id={0}, url={1}" -f $ContactId, $url)
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
        LogInfo ("CW GET contact -> OK")
        return $resp
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW GET contact failed: " + $_.Exception.Message + $suffix)
        return $null
    }
}

# ---------------------------
# CloudRadial v2 OData helpers  (UPDATED)
# ---------------------------

function Get-CloudRadialHeaders {
    # Use OData-friendly Accept; Basic Auth (public:private) is still required
    # CloudRadial shows Accept with OData minimal metadata in Swagger. [1](https://www.reddit.com/r/ConnectWise/comments/16293pl/best_practices_for_service_boards/)
    $publicKey  = [Environment]::GetEnvironmentVariable('CloudRadialCsa_ApiPublicKey')
    $privateKey = [Environment]::GetEnvironmentVariable('CloudRadialCsa_ApiPrivateKey')
    if ([string]::IsNullOrWhiteSpace($publicKey))  { throw "Missing app setting: CloudRadialCsa_ApiPublicKey" }
    if ([string]::IsNullOrWhiteSpace($privateKey)) { throw "Missing app setting: CloudRadialCsa_ApiPrivateKey" }
    $pair    = "{0}:{1}" -f $publicKey, $privateKey
    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
    return @{
        Authorization = "Basic $encoded"
        "Accept"       = "application/json;odata.metadata=minimal;odata.streaming=true"
        "Content-Type" = "application/json"
        "OData-Version" = "4.0"
    }
}

function Get-CloudRadialBaseUrl {
    # Keep your existing app setting; default stays US
    $base = [Environment]::GetEnvironmentVariable('CloudRadialCsa_BaseUrl')
    if ([string]::IsNullOrWhiteSpace($base)) { $base = "https://api.us.cloudradial.com" }
    return $base.TrimEnd('/')
}

# Find the company via PSA ID (CW ticket.company.id -> CloudRadial 'psaKey')
function Get-CloudRadialCompanyByPsaId {
    param([Parameter(Mandatory=$true)][int]$PsaCompanyId)
    $headers = Get-CloudRadialHeaders
    $baseUrl = Get-CloudRadialBaseUrl
    # OData v2: $top=1, $filter on numeric 'psaKey' (per Swagger sample). [1](https://www.reddit.com/r/ConnectWise/comments/16293pl/best_practices_for_service_boards/)
    $uri = '{0}/v2/odata/company?$top=1&$filter=psaKey eq {1}' -f $baseUrl, $PsaCompanyId
    LogInfo ("CloudRadial GET company by PSAId: {0} | {1}" -f $PsaCompanyId, $uri)
    try {
        $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        if ($resp -and $resp.PSObject.Properties['value'] -and $resp.value) { return ($resp.value | Select-Object -First 1) }
        return $null
    } catch {
        LogError ("CloudRadial company lookup (psaKey) failed: " + $_.Exception.Message)
        return $null
    }
}

# Fallback: find the company by PSA identifier (string 'psaIdentifier')
function Get-CloudRadialCompanyByIdentifier {
    param([Parameter(Mandatory=$true)][string]$Identifier)
    $headers = Get-CloudRadialHeaders
    $baseUrl = Get-CloudRadialBaseUrl
    # OData v2: $filter on string 'psaIdentifier'. [1](https://www.reddit.com/r/ConnectWise/comments/16293pl/best_practices_for_service_boards/)
    $safe = $Identifier.Replace("'","''")
    $uri  = '{0}/v2/odata/company?$top=1&$filter=psaIdentifier eq ''{1}''' -f $baseUrl, $safe
    LogInfo ("CloudRadial GET company by Identifier: {0} | {1}" -f $Identifier, $uri)
    try {
        $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        if ($resp -and $resp.PSObject.Properties['value'] -and $resp.value) { return ($resp.value | Select-Object -First 1) }
        return $null
    } catch {
        LogError ("CloudRadial company lookup (psaIdentifier) failed: " + $_.Exception.Message)
        return $null
    }
}

# Resolve TenantId using company tokens in v2:
# Tries embedded tokens via $expand, then a dedicated tokens feed (/v2/odata/companyToken)
function Get-CloudRadialTenantIdByCompanyId {
    param([int]$CompanyId)
    $headers = Get-CloudRadialHeaders
    $baseUrl = Get-CloudRadialBaseUrl

    # Try expand first (if the service exposes a navigation for tokens)
    try {
        $uriExpand = '{0}/v2/odata/company?$top=1&$filter=companyId eq {1}&$expand=companyTokens' -f $baseUrl, $CompanyId
        LogInfo ("CloudRadial GET company with tokens expand: {0}" -f $uriExpand)
        $resp = Invoke-RestMethod -Uri $uriExpand -Headers $headers -Method Get -ErrorAction Stop
        if ($resp -and $resp.value) {
            $company = $resp.value | Select-Object -First 1
            foreach ($propName in @('companyTokens','tokens','Tokens','CompanyTokens')) {
                if ($company.PSObject.Properties[$propName] -and $company.$propName) {
                    foreach ($t in @($company.$propName)) {
                        $n = ("" + ($t.tokenName ?? $t.name)).Trim()
                        $v = ("" + ($t.tokenValue ?? $t.value)).Trim()
                        if ($n -eq 'CompanyTenantId' -and $v) { return $v }
                    }
                }
            }
        }
    } catch {
        LogDebug ("CloudRadial expand tokens failed: " + $_.Exception.Message)
    }

    # Try a dedicated tokens feed (common OData pattern)
    try {
        # Filter on companyId and tokenName; select just the value
        $uriTokens = '{0}/v2/odata/companyToken?$top=1&$filter=companyId eq {1} and tokenName eq ''CompanyTenantId''' -f $baseUrl, $CompanyId
        LogInfo ("CloudRadial GET companyToken: {0}" -f $uriTokens)
        $resp2 = Invoke-RestMethod -Uri $uriTokens -Headers $headers -Method Get -ErrorAction Stop
        if ($resp2 -and $resp2.value) {
            $tok = $resp2.value | Select-Object -First 1
            foreach ($field in @('tokenValue','value','TokenValue','Value')) {
                if ($tok.PSObject.Properties[$field] -and $tok.$field) { return ("" + $tok.$field).Trim() }
            }
        }
    } catch {
        LogDebug ("CloudRadial companyToken lookup failed: " + $_.Exception.Message)
    }

    return $null
}

# ---------------------------
# Ticket-level UDF #54 updater (Client Department)
# ---------------------------
function Set-CwTicketDepartmentCustomField {
    param([int]$TicketId,[string]$DepartmentValue)
    $url    = 'https://{0}/v4_6_release/apis/3.0/service/tickets/{1}' -f $CwServer, $TicketId
    $ticket = Get-CwTicket -TicketId $TicketId
    if (-not $ticket) { return $false }

    $targetId = 54
    $existing = @()
    if ($ticket.customFields) { $existing = @($ticket.customFields) }

    $targetIndex = -1
    for ($i = 0; $i -lt $existing.Count; $i++) { if ( ($existing[$i].id -as [int]) -eq $targetId ) { $targetIndex = $i; break } }

    function ConvertTo-JsonArray {
        param([System.Collections.IList]$ops)
        if ($ops.Count -gt 1) { return ($ops | ConvertTo-Json -Depth 6) }
        $single = $ops[0] | ConvertTo-Json -Depth 6
        return ("[" + $single + "]")
    }

    try {
        if ($UseJsonPatch) {
            $ops = New-Object System.Collections.ArrayList
            if ($targetIndex -ge 0) { [void]$ops.Add(@{ op = "replace"; path = "/customFields/$targetIndex/value"; value = $DepartmentValue }) }
            else { [void]$ops.Add(@{ op = "add"; path = "/customFields/-"; value = @{ id = $targetId; value = $DepartmentValue } }) }
            $patch   = ConvertTo-JsonArray -ops $ops
            $headers = Get-CwHeaders -ContentType 'application/json'
            LogDebug ("CW JSON Patch body: {0}" -f $patch)
            try {
                $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $patch -ErrorAction Stop
                LogInfo ("CW PATCH (JSON Patch) customFields -> OK")
                return $true
            } catch {
                $errBody = Get-CwErrorBody -ex $_.Exception
                $suffix  = $errBody ? (" | Body: " + $errBody) : ""
                LogError ("CW PATCH (JSON Patch) failed: " + $_.Exception.Message + $suffix)
                LogInfo  ("Falling back to full-array object replace for customFields")
                $headers      = Get-CwHeaders -ContentType 'application/json'
                $customFields = @()
                if ($existing.Count -gt 0) { $customFields = @($existing) }
                if ($targetIndex -ge 0) { $customFields[$targetIndex].value = $DepartmentValue }
                else { $customFields += @{ id = $targetId; value = $DepartmentValue } }
                $body = @{ customFields = $customFields } | ConvertTo-Json -Depth 6
                LogDebug ("CW object replace body: {0}" -f $body)
                $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $body -ErrorAction Stop
                LogInfo ("CW PATCH (object replace) customFields -> OK")
                return $true
            }
        } else {
            $headers      = Get-CwHeaders -ContentType 'application/json'
            $customFields = @()
            if ($existing.Count -gt 0) { $customFields = @($existing) }
            if ($targetIndex -ge 0) { $customFields[$targetIndex].value = $DepartmentValue }
            else { $customFields += @{ id = $targetId; value = $DepartmentValue } }
            $body = @{ customFields = $customFields } | ConvertTo-Json -Depth 6
            LogDebug ("CW object replace body: {0}" -f $body)
            $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $body -ErrorAction Stop
            LogInfo ("CW PATCH (object replace) customFields -> OK")
            return $true
        }
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW PATCH customFields failed: " + $_.Exception.Message + $suffix)
        return $false
    }
}

# ---------------------------
# Add Note with fallback: Internal -> Discussion -> Resolution
# ---------------------------
function Add-CwTicketNote {
    param([int]$TicketId,[string]$Text,[bool]$InternalAnalysisFirst = $true)
    $headers = Get-CwHeaders -ContentType 'application/json; charset=utf-8'
    $url     = 'https://{0}/v4_6_release/apis/3.0/service/tickets/{1}/notes' -f $CwServer, $TicketId

    $memberId    = [Environment]::GetEnvironmentVariable('ConnectWisePsa_MemberId')
    $memberIdent = [Environment]::GetEnvironmentVariable('ConnectWisePsa_MemberIdentifier')
    function AttachMember { param([hashtable]$payload)
        if ($memberId -or $memberIdent) {
            $member = @{}
            if ($memberId)    { $member.id        = ($memberId -as [int]) }
            if ($memberIdent) { $member.identifier = $memberIdent }
            $payload.member = $member
        }
        return $payload
    }

    function New-NotePayload { param([bool]$internal,[bool]$discussion,[bool]$resolution,[string]$text)
        $maxLen = 2000
        if ($text.Length -gt $maxLen) { $text = $text.Substring(0,$maxLen) + "`n(truncated)" }
        $payload = @{
            ticketId              = $TicketId
            text                  = $text
            internalAnalysisFlag  = $internal
            internalFlag          = $internal
            detailDescriptionFlag = $discussion
            resolutionFlag        = $resolution
            externalFlag          = $false
            customerUpdatedFlag   = $false
        }
        $payload = AttachMember -payload $payload
        return ($payload | ConvertTo-Json -Depth 5)
    }

    $body1 = New-NotePayload -internal $InternalAnalysisFirst -discussion $false -resolution $false -text $Text
    LogInfo  ("CW Add Note: TicketId={0}, internalAnalysisFlag={1}" -f $TicketId, $InternalAnalysisFirst)
    LogDebug ("CW Note body (truncated): {0}" -f $body1.Substring(0, [Math]::Min(600, $body1.Length)))
    try { Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body1 -ErrorAction Stop; LogInfo ("CW Add Note -> OK"); return $true }
    catch { $errBody = Get-CwErrorBody -ex $_.Exception; $suffix  = $errBody ? (" | Body: " + $errBody) : ""; LogError ("CW Add Note failed: " + $_.Exception.Message + $suffix) }

    $body2 = New-NotePayload -internal $false -discussion $true -resolution $false -text $Text
    LogInfo  ("CW Add Note (Discussion): TicketId={0}" -f $TicketId)
    LogDebug ("CW Note body (truncated): {0}" -f $body2.Substring(0, [Math]::Min(600, $body2.Length)))
    try { Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body2 -ErrorAction Stop; LogInfo ("CW Add Note (Discussion) -> OK"); return $true }
    catch { $errBody = Get-CwErrorBody -ex $_.Exception; $suffix  = $errBody ? (" | Body: " + $errBody) : ""; LogError ("CW Add Note (Discussion) failed: " + $_.Exception.Message + $suffix) }

    $body3 = New-NotePayload -internal $false -discussion $false -resolution $true -text $Text
    LogInfo  ("CW Add Note (Resolution): TicketId={0}" -f $TicketId)
    LogDebug ("CW Note body (truncated): {0}" -f $body3.Substring(0, [Math]::Min(600, $body3.Length)))
    try { Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body3 -ErrorAction Stop; LogInfo ("CW Add Note (Resolution) -> OK"); return $true }
    catch { $errBody = Get-CwErrorBody -ex $_.Exception; $suffix  = $errBody ? (" | Body: " + $errBody) : ""; LogError ("CW Add Note (Resolution) failed: " + $_.Exception.Message + $suffix); return $false }
}

# ---------------------------
# SecurityKey validation
# ---------------------------
$secKey = [Environment]::GetEnvironmentVariable('SecurityKey')
$reqKey = $null
if ($Request -and $Request.PSObject.Properties['Headers']) { $reqKey = $Request.Headers['SecurityKey'] }
if ($secKey -and (-not $reqKey) -and $Request -and $Request.PSObject.Properties['Query']) { $reqKey = $Request.Query['securityKey'] }
if ($secKey) { if (-not $reqKey -or $reqKey -ne $secKey) { LogError "Invalid or missing SecurityKey"; New-JsonResponse -Code 401 -Message "Invalid or missing SecurityKey"; return } }

# ---------------------------
# Inputs
# ---------------------------
$body = Get-RequestBodyObject -Request $Request
try {
    if ($IsDebug -and $body) {
        $bodyPreview = $body | ConvertTo-Json -Depth 6
        if ($bodyPreview.Length -gt 600) { $bodyPreview = $bodyPreview.Substring(0,600) + "...(truncated)" }
        LogDebug ("Parsed body preview: " + $bodyPreview)
    }
} catch {}

function Get-Prop { param([object]$obj,[string]$name) if ($null -eq $obj) { return $null } $p = $obj.PSObject.Properties[$name]; if ($null -ne $p) { return $p.Value }; return $null }

# ---------------------------
# TicketId extraction
# ---------------------------
[int]$TicketId = 0
$TicketId = To-IntOrNull (Get-Prop $body 'TicketId')
if (-not $TicketId) { $ticketObj = Get-Prop $body 'Ticket'; if ($ticketObj) { $TicketId = To-IntOrNull (Get-Prop $ticketObj 'TicketId') } }
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Query'])   { $TicketId = To-IntOrNull $Request.Query['ticketId'] }
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Headers']) { $TicketId = To-IntOrNull $Request.Headers['TicketId'] }
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Query'])   { $TicketId = To-IntOrNull $Request.Query['id'] }
if (-not $TicketId) {
    $TicketId = To-IntOrNull (Get-Prop $body 'ID')
    if (-not $TicketId) {
        $Entity = Get-Prop $body 'Entity'
        if ($Entity -is [string] -and $Entity.Trim().StartsWith('{')) { try { $Entity = $Entity | ConvertFrom-Json -ErrorAction Stop } catch { } }
        if ($Entity -and $Entity.PSObject.Properties['id']) { $TicketId = To-IntOrNull $Entity.id }
    }
}
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Query']) {
    $routeVal = $Request.Query['route']
    if ($routeVal) {
        $m = [Regex]::Match(("" + $routeVal), '\d+$')
        if ($m.Success) { $TicketId = To-IntOrNull $m.Value; if ($IsDebug) { LogDebug ("Derived TicketId from route='{0}' -> {1}" -f $routeVal, $TicketId) } }
    }
}
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Headers']) {
    $urlHeaders = @('x-original-url','x-waws-unencoded-url')
    foreach ($h in $urlHeaders) {
        $u = $Request.Headers[$h]
        if ($u) {
            # Handle both raw '&' and HTML '&amp;'
            $mId    = [Regex]::Match(("" + $u), '(\?|&|&amp;)id=(\d+)')
            $mRoute = [Regex]::Match(("" + $u), '(\?|&|&amp;)route=[^&]*?(\d+)')
            $candidate = $null
            if     ($mId.Success)    { $candidate = $mId.Groups[2].Value }
            elseif ($mRoute.Success) { $candidate = $mRoute.Groups[2].Value }
            if ($candidate) { $TicketId = To-IntOrNull $candidate; if ($TicketId) { if ($IsDebug) { LogDebug ("Derived TicketId from header {0}='{1}' -> {2}" -f $h, $u, $TicketId) }; break } }
        }
    }
}
if (-not $TicketId -or $TicketId -le 0) { LogError "TicketId is missing or invalid (parsed as null/0)."; New-JsonResponse -Code 400 -Message "TicketId is required; pass as body.TicketId, body.Ticket.TicketId, query ?id or ?ticketId, header TicketId, or include 'ID' / 'Entity.id' (or route=service<id>)."; return }

# ---------------------------
# Resolve CW ticket
# ---------------------------
$ticket = Get-CwTicket -TicketId $TicketId
if (-not $ticket) { New-JsonResponse -Code 500 -Message "Unable to read CW ticket."; return }

# ---------------------------
# TenantId: direct OR CloudRadial v2 OData by PSA company id -> CompanyTenantId
# ---------------------------
$TenantId       = $null
$TenantIdSource = $null
$TenantId = Get-Prop $body 'TenantId'
if ($TenantId) { $TenantIdSource = 'Body.TenantId' }
if (-not $TenantId -and $Request -and $Request.PSObject.Properties['Query'])   { $TenantId = $Request.Query['tenantId']; if ($TenantId) { $TenantIdSource = 'Query.tenantId' } }
if (-not $TenantId -and $Request -and $Request.PSObject.Properties['Headers']) { $TenantId = $Request.Headers['TenantId']; if ($TenantId) { $TenantIdSource = 'Header.TenantId' } }

if (-not $TenantId) {
    $cwCompanyId = $null
    if ($ticket.PSObject.Properties['company'] -and $ticket.company.PSObject.Properties['id']) { $cwCompanyId = ($ticket.company.id -as [int]) }

    if ($cwCompanyId) {
        # v2: look up company via psaKey; fallback: psaIdentifier
        $crCompany  = Get-CloudRadialCompanyByPsaId -PsaCompanyId $cwCompanyId
        if (-not $crCompany -and $ticket.company.PSObject.Properties['identifier']) {
            $crCompany = Get-CloudRadialCompanyByIdentifier -Identifier ("" + $ticket.company.identifier)
        }

        # Resolve TenantId via tokens (expand or companyToken feed)
        $crTenantId = $null
        if ($crCompany -and $crCompany.PSObject.Properties['companyId']) {
            $crTenantId = Get-CloudRadialTenantIdByCompanyId -CompanyId ($crCompany.companyId -as [int])
        }

        if ($crCompany -and $crTenantId) {
            $TenantId       = $crTenantId
            $TenantIdSource = "CloudRadial.v2.companyToken(CompanyTenantId)"
            LogInfo ("Resolved TenantId via CloudRadial v2 OData: {0} (companyId={1}, psaKey={2}, psaIdentifier='{3}')" -f $TenantId, $crCompany.companyId, $crCompany.psaKey, $crCompany.psaIdentifier)
        } else {
            LogInfo ("CloudRadial v2 did not return CompanyTenantId (companyId={0}, psaKey={1}, identifier='{2}')" -f ($crCompany?.companyId), ($crCompany?.psaKey), ($crCompany?.psaIdentifier))
        }
    } else {
        LogInfo ("CW ticket has no company.id; skipping CloudRadial lookup.")
    }
}

# Guardrails
$PartnerTenantId = [Environment]::GetEnvironmentVariable('Ms365_TenantId')
if ([string]::IsNullOrWhiteSpace($TenantId)) { LogError "TenantId missing; CloudRadial API lookup failed or not configured."; New-JsonResponse -Code 400 -Message "TenantId is required; ensure CloudRadial v2 has CompanyTenantId token or pass TenantId via body/query/header."; return }
if ($PartnerTenantId -and $TenantId -eq $PartnerTenantId) { LogError "TenantId equals partner Ms365_TenantId; refusing client lookup."; New-JsonResponse -Code 400 -Message "Client TenantId required; partner tenant not allowed."; return }

# ---------------------------
# UPN/Email
# ---------------------------
$UserUPN   = Get-Prop $body 'UserOfficeId'
$UserEmail = Get-Prop $body 'UserEmail'
$userObj   = Get-Prop $body 'User'
if (-not $UserUPN   -and $userObj) { $UserUPN   = Get-Prop $userObj 'UserOfficeId' }
if (-not $UserEmail -and $userObj) { $UserEmail = Get-Prop $userObj 'Email' }
if (-not $UserUPN   -and $Request -and $Request.PSObject.Properties['Query'])   { $UserUPN   = $Request.Query['userUpn'] }
if (-not $UserEmail -and $Request -and $Request.PSObject.Properties['Query'])   { $UserEmail = $Request.Query['userEmail'] }
if (-not $UserUPN   -and $Request -and $Request.PSObject.Properties['Headers']) { $UserUPN   = $Request.Headers['UserUPN'] }
if (-not $UserEmail -and $Request -and $Request.PSObject.Properties['Headers']) { $UserEmail = $Request.Headers['UserEmail'] }
if (-not $UserEmail) {
    if ($ticket.PSObject.Properties['contact'] -and $ticket.contact.PSObject.Properties['email']) { $UserEmail = $ticket.contact.email }
    elseif ($ticket.PSObject.Properties['companyContact'] -and $ticket.companyContact.PSObject.Properties['email']) { $UserEmail = $ticket.companyContact.email }
}

LogInfo ("Inputs: TicketId={0}, TenantId={1} (source={2}), UPN='{3}', Email='{4}'" -f $TicketId, $TenantId, $TenantIdSource, $UserUPN, $UserEmail)

# ---------------------------
# Graph & department
# ---------------------------
if (-not (Connect-GraphApp -TenantId $TenantId)) {
    New-JsonResponse -Code 500 -Message "Failed to connect to Microsoft Graph" -Extra @{ TicketId=$TicketId; TenantId=$TenantId; TenantIdSource=$TenantIdSource }; return
}
$department = Get-UserDepartment -UserPrincipalName $UserUPN -UserEmail $UserEmail

# ---------------------------
# Fallback via CW Contact UDF #53
# ---------------------------
$source = "EntraID"; $fallbackContactId = $null
if ([string]::IsNullOrWhiteSpace($department)) {
    LogInfo "Department blank from Graph; attempting CW contact UDF fallback (id=53)"
    $contactId = $null
    if     ($ticket.PSObject.Properties['contact'] -and $ticket.contact.PSObject.Properties['id']) { $contactId = ($ticket.contact.id -as [int]) }
    elseif ($ticket.PSObject.Properties['companyContact'] -and $ticket.companyContact.PSObject.Properties['id']) { $contactId = ($ticket.companyContact.id -as [int]) }
    elseif ($ticket.PSObject.Properties['contactId']) { $contactId = ($ticket.contactId -as [int]) }
    if ($contactId) {
        $contact = Get-CwContactById -ContactId $contactId
        if ($contact -and $contact.customFields) {
            foreach ($cf in @($contact.customFields)) {
                $cfIdInt = ($cf.id -as [int]); $cap = ("" + $cf.caption).Trim()
                if ($cfIdInt -eq 53 -or $cap -eq "Client Department") { $department = $cf.value; $source = "CWContactUDF53"; $fallbackContactId = $contactId; break }
            }
        }
    }
}

if (-not $department) { $department = "" }
LogInfo ("Final department (source={0}) = '{1}'" -f $source, (""+$department))

# ---------------------------
# Update UDF #54 when appropriate
# ---------------------------
$SkipEmpty = ([Environment]::GetEnvironmentVariable('SkipEmptyDepartment') -as [int]) -eq 1
if ($SkipEmpty -and [string]::IsNullOrWhiteSpace($department)) {
    LogInfo "SkipEmptyDepartment is ON; department is blank—skipping PATCH."
    $ok = $false
} else {
    $ok = Set-CwTicketDepartmentCustomField -TicketId $TicketId -DepartmentValue $department
}

# ---------------------------
# Verify UDF #54
# ---------------------------
$verifyUdfValue = $null
try {
    $verify = Get-CwTicket -TicketId $TicketId
    if ($verify -and $verify.customFields) {
        foreach ($cf in @($verify.customFields)) {
            $cfIdInt = ($cf.id -as [int]); $cap = ("" + $cf.caption).Trim()
            if ($cfIdInt -eq 54 -or $cap -eq "Client Department") { $verifyUdfValue = $cf.value; break }
        }
    }
    LogInfo ("Verify UDF #54 after PATCH: '{0}'" -f (""+$verifyUdfValue))
} catch { LogError ("Verify after PATCH failed: " + $_.Exception.Message) }

# ---------------------------
# Audit note
# ---------------------------
$timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss zzz")
$noteLines = @(
    "**Department sync audit**",
    ("- TenantId: {0} (source={1})" -f $TenantId, $TenantIdSource),
    ("- Source: {0}" -f $source),
    ("- Applied value: '{0}'" -f $department),
    ("- Submitter UPN: '{0}'" -f $UserUPN),
    ("- Submitter Email: '{0}'" -f $UserEmail),
    ("- Verify UDF #54 after PATCH: '{0}'" -f $verifyUdfValue),
    ("- Timestamp: {0}" -f $timestamp),
    ("- CorrelationId: {0}" -f $CorrelationId)
)
$noteText = ($noteLines -join [Environment]::NewLine)
$noteOk = Add-CwTicketNote -TicketId $TicketId -Text $noteText -InternalAnalysisFirst $true

# ---------------------------
# Respond
# ---------------------------
if ($ok) {
    New-JsonResponse -Code 200 -Message "Updated CW ticket UDF #54 (Client Department) and logged audit note." -Extra @{
        TicketId        = $TicketId
        TenantId        = $TenantId
        TenantIdSource  = $TenantIdSource
        Department      = $department
        Source          = $source
        NoteAdded       = $noteOk
        VerifyUdf54     = $verifyUdfValue
        JsonPatchMode   = $UseJsonPatch
    }
} else {
    New-JsonResponse -Code 500 -Message "Failed to update CW ticket UDF #54 (Client Department). Audit note was attempted." -Extra @{
        TicketId        = $TicketId
        TenantId        = $TenantId
        TenantIdSource  = $TenantIdSource
        Department      = $department
        Source          = $source
        NoteAdded       = $noteOk
        VerifyUdf54     = $verifyUdfValue
        JsonPatchMode   = $UseJsonPatch
    }
}
