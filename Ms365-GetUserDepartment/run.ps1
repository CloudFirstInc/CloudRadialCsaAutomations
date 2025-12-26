
# Ms365-GetUserDepartment/run.ps1
# TenantId resolution via OIDC discovery from domain (no CloudRadial).
# CW dept UDF update + Graph integration, with rate-limit aware retries.
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
    param([Parameter(Mandatory=$true)][string]$name,[Parameter()][object]$coll)
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
# Helpers
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
function Get-Prop { param([object]$obj,[string]$name) if ($null -eq $obj) { return $null } $p = $obj.PSObject.Properties[$name]; if ($null -ne $p) { return $p.Value }; return $null }
function Get-PropText { param([object]$obj,[string]$name) if ($null -eq $obj) { return "" } $p = $obj.PSObject.Properties[$name]; if ($null -ne $p -and $null -ne $p.Value) { return ("" + $p.Value) } return "" }

# ---------------------------
# Probe mode
# ---------------------------
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

# --- Generic retry wrapper for idempotent GETs (429/408/5xx) ---
function Invoke-HttpWithRetry {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][hashtable]$Headers,
        [int]$MaxRetries = 5,
        [int]$BaseDelayMs = 1000,     # 1s
        [int]$MaxDelayMs  = 15000     # 15s cap
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($IsDebug) { LogDebug ("HTTP GET [{0}/{1}] {2}" -f $attempt, $MaxRetries, $Uri) }
            $resp = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -TimeoutSec 20 -ErrorAction Stop
            return $resp
        } catch {
            $status = $null
            try { $status = $_.Exception.Response.StatusCode.value__ } catch {}
            # Read Retry-After if present
            $retryAfterSec = $null
            try {
                $hdrs = $_.Exception.Response.Headers
                if ($hdrs) {
                    $ra = $hdrs['Retry-After']
                    if ($ra) {
                        $sec = 0
                        if ([int]::TryParse(("" + $ra), [ref]$sec)) { $retryAfterSec = $sec }
                    }
                }
            } catch {}

            $isTransient = ($status -eq 429 -or $status -eq 408 -or ($status -ge 500 -and $status -lt 600))
            if (-not $isTransient -or $attempt -ge $MaxRetries) {
                LogError ("HTTP GET failed (status={0}, attempt={1}/{2}): {3}" -f $status, $attempt, $MaxRetries, $_.Exception.Message)
                if ($_.Exception.Response) {
                    try {
                        $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                        $body = $sr.ReadToEnd()
                        if ($body) { LogDebug ("HTTP error body: " + $body) }
                    } catch {}
                }
                throw
            }

            $delayMs = $BaseDelayMs * [math]::Pow(2, ($attempt - 1))
            if ($delayMs -gt $MaxDelayMs) { $delayMs = $MaxDelayMs }
            $jitter = [math]::Round($delayMs * (Get-Random -Minimum 0.0 -Maximum 0.30))
            if ($retryAfterSec) {
                $delayMs = [math]::Min($retryAfterSec * 1000, $MaxDelayMs)
                $jitter  = 0
                LogInfo ("Throttled (429). Honoring Retry-After={0}s; sleeping {1} ms before retry." -f $retryAfterSec, $delayMs)
            } else {
                LogInfo ("Transient HTTP {0}. Backing off (attempt {1}/{2}) for {3} ms (jitter {4} ms)." -f $status, $attempt, $MaxRetries, $delayMs, $jitter)
            }
            Start-Sleep -Milliseconds ($delayMs + $jitter)
        }
    }
}

function Get-CwTicket {
    param([int]$TicketId)
    $headers = Get-CwHeaders -ContentType 'application/json'
    $url     = 'https://{0}/v4_6_release/apis/3.0/service/tickets/{1}' -f $CwServer, $TicketId
    LogInfo ("CW GET ticket: id={0}, url={1}" -f $TicketId, $url)
    try {
        $resp = Invoke-HttpWithRetry -Uri $url -Headers $headers -MaxRetries 6 -BaseDelayMs 1000 -MaxDelayMs 20000
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
        $resp = Invoke-HttpWithRetry -Uri $url -Headers $headers -MaxRetries 6 -BaseDelayMs 1000 -MaxDelayMs 20000
        LogInfo ("CW GET contact -> OK")
        return $resp
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW GET contact failed: " + $_.Exception.Message + $suffix)
        return $null
    }
}

# --- NEW: Get contacts by company & extract email for domain ---
function Get-CwCompanyContacts {
    param(
        [Parameter(Mandatory=$true)][int]$CompanyId,
        [int]$PageSize = 50
    )
    $headers = Get-CwHeaders -ContentType 'application/json'
    $url = 'https://{0}/v4_6_release/apis/3.0/company/contacts?conditions=company/id={1}&pageSize={2}' -f $CwServer, $CompanyId, $PageSize
    LogInfo ("CW GET contacts by company: companyId={0}, url={1}" -f $CompanyId, $url)
    try {
        $resp = Invoke-HttpWithRetry -Uri $url -Headers $headers -MaxRetries 6 -BaseDelayMs 1000 -MaxDelayMs 20000
        if ($resp) { return @($resp) }
        return @()
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW GET contacts by company failed: " + $_.Exception.Message + $suffix)
        return @()
    }
}

function Get-CwContactPrimaryEmail {
    param([Parameter(Mandatory=$true)][object]$Contact)
    # Check embedded communicationItems first
    if ($Contact.PSObject.Properties['communicationItems'] -and $Contact.communicationItems) {
        foreach ($ci in @($Contact.communicationItems)) {
            $val = ("" + $ci.value)
            if ($val -and ($val -match '@')) { return $val }
        }
    }
    # Fallback to communications endpoint
    try {
        if ($Contact.PSObject.Properties['id']) {
            $contactId = ($Contact.id -as [int])
            if ($contactId -gt 0) {
                $headers = Get-CwHeaders -ContentType 'application/json'
                $urlComm = 'https://{0}/v4_6_release/apis/3.0/company/contacts/{1}/communications' -f $CwServer, $contactId
                LogInfo ("CW GET contact communications: id={0}, url={1}" -f $contactId, $urlComm)
                $comms = Invoke-HttpWithRetry -Uri $urlComm -Headers $headers -MaxRetries 6 -BaseDelayMs 1000 -MaxDelayMs 20000
                foreach ($c in @($comms)) {
                    $val = ("" + $c.value)
                    if ($val -and ($val -match '@')) { return $val }
                }
            }
        }
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW GET contact communications failed: " + $_.Exception.Message + $suffix)
    }
    return $null
}

# ---------------------------
# TenantId discovery (OIDC well-known)  — NEW
# ---------------------------
function Invoke-Json {
    param([string]$Uri)
    try {
        $headers = @{ 'Accept'='application/json' }
        if ($IsDebug) { LogDebug ("GET " + $Uri) }
        $resp = Invoke-HttpWithRetry -Uri $Uri -Headers $headers -MaxRetries 4 -BaseDelayMs 750 -MaxDelayMs 8000
        return $resp
    } catch {
        LogDebug ("OIDC discovery request failed for {0}: {1}" -f $Uri, $_.Exception.Message)
        return $null
    }
}
function Extract-GuidFromUrl {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $null }
    $m = [Regex]::Match($Url, '(?i)[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}')
    if ($m.Success) { return $m.Value }
    return $null
}
function Get-TenantIdFromDomain {
    param([Parameter(Mandatory=$true)][string]$Domain)
    # Try v2.0 well-known first (recommended by Microsoft identity platform)
    $urlV2 = "https://login.microsoftonline.com/$Domain/v2.0/.well-known/openid-configuration"
    $json  = Invoke-Json -Uri $urlV2
    $tid   = $null
    if ($json) {
        $tid = Extract-GuidFromUrl -Url ("" + $json.issuer)
        if (-not $tid) { $tid = Extract-GuidFromUrl -Url ("" + $json.token_endpoint) }
        if ($tid) { return $tid }
    }
    # Fallback to non-v2 well-known (some tenants)
    $urlV1   = "https://login.microsoftonline.com/$Domain/.well-known/openid-configuration"
    $jsonV1  = Invoke-Json -Uri $urlV1
    if ($jsonV1) {
        $tid = Extract-GuidFromUrl -Url ("" + $jsonV1.issuer)
        if (-not $tid) { $tid = Extract-GuidFromUrl -Url ("" + $jsonV1.token_endpoint) }
        if ($tid) { return $tid }
    }
    return $null
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

# ---------------------------
# TicketId extraction (robust)
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
# Resolve CW ticket (for contact/company fallback)
# ---------------------------
$ticket = Get-CwTicket -TicketId $TicketId
if (-not $ticket) { New-JsonResponse -Code 500 -Message "Unable to read CW ticket."; return }

# ---------------------------
# TenantId: direct OR OIDC discovery from domain (with CW contacts fallback)
# ---------------------------
$TenantId       = $null
$TenantIdSource = $null
$TenantId = Get-Prop $body 'TenantId'
if ($TenantId) { $TenantIdSource = 'Body.TenantId' }
if (-not $TenantId -and $Request -and $Request.PSObject.Properties['Query'])   { $TenantId = $Request.Query['tenantId']; if ($TenantId) { $TenantIdSource = 'Query.tenantId' } }
if (-not $TenantId -and $Request -and $Request.PSObject.Properties['Headers']) { $TenantId = $Request.Headers['TenantId']; if ($TenantId) { $TenantIdSource = 'Header.TenantId' } }

# If not passed, derive via OIDC discovery using a domain from email; else CW contacts fallback
if (-not $TenantId) {
    # Collect best-available email first
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

    # Derive domain from email
    $Domain = $null
    if ($UserEmail -and ($UserEmail -match '@')) {
        $Domain = ($UserEmail.Split('@') | Select-Object -Last 1).Trim()
    } elseif ($UserUPN -and ($UserUPN -match '@')) {
        $Domain = ($UserUPN.Split('@') | Select-Object -Last 1).Trim()
    }

    if ($Domain) {
        LogInfo ("Attempting OIDC discovery for tenant via domain: {0}" -f $Domain)
        $tidFromDomain = Get-TenantIdFromDomain -Domain $Domain
        if ($tidFromDomain) {
            $TenantId       = $tidFromDomain
            $TenantIdSource = "OIDCDiscovery(domain)"
            LogInfo ("Resolved TenantId via OIDC discovery: {0} (domain={1})" -f $TenantId, $Domain)
        } else {
            LogInfo ("OIDC discovery did not return TenantId for domain={0}" -f $Domain)
        }
    } else {
        LogInfo ("No email domain available; attempting CW company contacts for domain fallback.")
        # Try pulling a contact email from CW by company.id
        $cwCompanyId = $null
        if ($ticket.PSObject.Properties['company'] -and $ticket.company.PSObject.Properties['id']) {
            $cwCompanyId = ($ticket.company.id -as [int])
        }
        if ($cwCompanyId -gt 0) {
            $contacts = Get-CwCompanyContacts -CompanyId $cwCompanyId -PageSize 25
            $emailFromContacts = $null
            foreach ($ct in $contacts) {
                $emailFromContacts = Get-CwContactPrimaryEmail -Contact $ct
                if ($emailFromContacts) { break }
            }
            if ($emailFromContacts -and ($emailFromContacts -match '@')) {
                $Domain = ($emailFromContacts.Split('@') | Select-Object -Last 1).Trim()
                LogInfo ("Derived domain from CW contacts: {0}" -f $Domain)
                $tidFromDomain = Get-TenantIdFromDomain -Domain $Domain
                if ($tidFromDomain) {
                    $TenantId       = $tidFromDomain
                    $TenantIdSource = "OIDCDiscovery(domain-from-CWContact)"
                    LogInfo ("Resolved TenantId via OIDC discovery: {0} (domain={1})" -f $TenantId, $Domain)
                } else {
                    LogInfo ("OIDC discovery did not return TenantId for domain (from CW contacts)={0}" -f $Domain)
                }
            } else {
                LogInfo ("CW company contacts had no email; cannot derive domain.")
            }
        } else {
            LogInfo ("CW ticket had no company.id; cannot query contacts.")
        }
    }
}

# Guardrails
$PartnerTenantId = [Environment]::GetEnvironmentVariable('Ms365_TenantId')
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    LogError "TenantId missing; OIDC discovery failed or not configured."
    New-JsonResponse -Code 400 -Message "TenantId is required; pass TenantId via body/query/header or ensure a valid email domain (via payload or CW contacts) is available for discovery."
    return
}
if ($PartnerTenantId -and $TenantId -eq $PartnerTenantId) {
    LogError "TenantId equals partner Ms365_TenantId; refusing client lookup."
    New-JsonResponse -Code 400 -Message "Client TenantId required; partner tenant not allowed."
    return
}

# ---------------------------
# Recompute UPN/Email (for subsequent Graph work)
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
    # UDF patch helpers
    $ok = $false
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
        function ConvertTo-JsonArray { param([System.Collections.IList]$ops)
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
                try { Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $patch -ErrorAction Stop; LogInfo ("CW PATCH (JSON Patch) customFields -> OK"); return $true }
                catch {
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
                    Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $body -ErrorAction Stop
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
                Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $body -ErrorAction Stop
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

# Add Note with fallback
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
