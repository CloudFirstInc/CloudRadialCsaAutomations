
# Ms365-GetUserDepartment/run.ps1
# CloudRadial API tenant lookup (by PSA Company Id) + CW dept UDF update
# FIX: brace scoped variables in strings (e.g., ${script:CwServer}) to avoid colon parsing errors.

using namespace System.Net

param($Request, $TriggerMetadata)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------
# Correlation & Logging
# ---------------------------
$CorrelationId = [Guid]::NewGuid().ToString()
$IsDebug       = ([Environment]::GetEnvironmentVariable('DebugLogging') -as [int]) -eq 1

function LogInfo { param([string]$msg) Write-Information ("[$CorrelationId] " + $msg) }
function LogError { param([string]$msg) Write-Error -Message ("[$CorrelationId] " + $msg) -ErrorAction Continue }
function LogDebug { param([string]$msg) if ($IsDebug) { Write-Information ("[$CorrelationId][DEBUG] " + $msg) } }

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
        try { if ($raw.Trim().StartsWith('{')) { return $raw | ConvertFrom-Json -ErrorAction Stop } }
        catch { }
        return @{}
    }
    return $raw
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
        LogInfo ("Connecting to Graph (TenantId=$TenantId, AppId=$appId)")
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
            $filter    = "mail eq '$safeEmail' or userPrincipalName eq '$safeEmail'"
            LogInfo ("Graph lookup by filter: $filter")
            $user      = Get-MgUser -Filter $filter -Property department,userPrincipalName -Top 1 -ErrorAction Stop
        } catch { LogError ("Get-MgUser by email filter failed: " + $_.Exception.Message) }
    }
    if (-not $user -and $UserPrincipalName) {
        try { LogInfo ("Graph lookup by Id/UPN: $UserPrincipalName"); $user = Get-MgUser -UserId $UserPrincipalName -Property department,userPrincipalName -ErrorAction Stop }
        catch { LogError ("Get-MgUser by Id/UPN failed: " + $_.Exception.Message) }
    }
    if ($user) { LogInfo ("Graph user department = '" + (""+$user.Department) + "'"); return $user.Department }
    LogInfo ("Graph user not found or no department."); return $null
}

# ---------------------------
# ConnectWise helpers
# ---------------------------
$script:CwServer   = 'api-na.myconnectwise.net'
$UseJsonPatch      = ([Environment]::GetEnvironmentVariable('ConnectWise_UseJsonPatch') -as [int]) -eq 1

function Get-CwHeaders {
    param([string]$ContentType = 'application/json')
    $required = 'ConnectWisePsa_ApiCompanyId','ConnectWisePsa_ApiPublicKey','ConnectWisePsa_ApiPrivateKey','ConnectWisePsa_ApiClientId'
    foreach ($n in $required) { if ([string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($n))) { throw "Missing ConnectWise app setting: $n" } }
    $companyId = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiCompanyId')
    $pubKey    = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiPublicKey')
    $privKey   = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiPrivateKey')
    $clientId  = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiClientId')
    $authString  = "${companyId}+${pubKey}:${privKey}"
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
    return @{
        "Authorization" = "Basic $encodedAuth"
        "ClientID"      = $clientId
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
    $url     = "https://${script:CwServer}/v4_6_release/apis/3.0/service/tickets/$TicketId"
    LogInfo ("CW GET ticket: id=$TicketId, url=$url")
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
    $url     = "https://${script:CwServer}/v4_6_release/apis/3.0/company/contacts/$ContactId"
    LogInfo ("CW GET contact: id=$ContactId, url=$url")
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
# CloudRadial API helpers
# ---------------------------
function Get-CloudRadialHeaders {
    $publicKey  = [Environment]::GetEnvironmentVariable('CloudRadialCsa_ApiPublicKey')
    $privateKey = [Environment]::GetEnvironmentVariable('CloudRadialCsa_ApiPrivateKey')
    if ([string]::IsNullOrWhiteSpace($publicKey))  { throw "Missing app setting: CloudRadialCsa_ApiPublicKey" }
    if ([string]::IsNullOrWhiteSpace($privateKey)) { throw "Missing app setting: CloudRadialCsa_ApiPrivateKey" }
    $pair    = "$publicKey:$privateKey"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
    return @{
        Authorization = "Basic $encoded"
        "Content-Type" = "application/json"
        "Accept"       = "application/json"
    }
}

function Get-CloudRadialBaseUrl {
    $base = [Environment]::GetEnvironmentVariable('CloudRadialCsa_BaseUrl')
    if ([string]::IsNullOrWhiteSpace($base)) { $base = "https://api.us.cloudradial.com" } # default US endpoint
    return $base.TrimEnd('/')
}

function Get-CloudRadialCompanyByPsaId {
    param([Parameter(Mandatory=$true)][int]$PsaCompanyId)
    $headers = Get-CloudRadialHeaders
    $baseUrl = Get-CloudRadialBaseUrl
    # CloudRadial API supports Filter/Condition/Value params on /api/company. [4](https://www.reddit.com/r/ConnectWise/comments/vi96w3/connectwise_api_triggers/)
    $uri = "$baseUrl/api/company?Filter=psaid&Condition=eq&Value=$PsaCompanyId&Take=1"
    LogInfo ("CloudRadial GET company by PSAId: $PsaCompanyId | $uri")
    try {
        $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
        if ($resp -is [System.Collections.IEnumerable]) {
            return ($resp | Select-Object -First 1)
        } elseif ($resp.PSObject.Properties['data']) {
            return ($resp.data | Select-Object -First 1)
        } else {
            return $resp
        }
    } catch {
        LogError ("CloudRadial company lookup failed: " + $_.Exception.Message)
        return $null
    }
}

function Get-CloudRadialTenantIdFromCompany {
    param([object]$Company)
    if (-not $Company) { return $null }

    # Predefined token @CompanyTenantId = client's Microsoft 365 Tenant ID. [5](https://docs.webhook.site/custom-actions/variables.html)
    $paths = @('tokens.CompanyTenantId','Tokens.CompanyTenantId','companyTokens.CompanyTenantId')
    foreach ($p in $paths) {
        $parts = $p.Split('.')
        $node  = $Company
        foreach ($part in $parts) {
            if ($node -and $node.PSObject.Properties[$part]) { $node = $node.$part } else { $node = $null; break }
        }
        if ($node) { return ("" + $node).Trim() }
    }
    foreach ($prop in $Company.PSObject.Properties) {
        if ($prop.Name -match 'CompanyTenantId' -and $prop.Value) {
            return ("" + $prop.Value).Trim()
        }
    }
    return $null
}

# ---------------------------
# Ticket-level UDF #54 updater (Client Department)
# ---------------------------
function Set-CwTicketDepartmentCustomField {
    param([int]$TicketId,[string]$DepartmentValue)

    $url    = "https://${script:CwServer}/v4_6_release/apis/3.0/service/tickets/$TicketId"
    $ticket = Get-CwTicket -TicketId $TicketId
    if (-not $ticket) { return $false }

    $targetId    = 54
    $existing    = @()
    if ($ticket.customFields) { $existing = @($ticket.customFields) }

    $targetIndex = -1
    for ($i = 0; $i -lt $existing.Count; $i++) {
        if ( ($existing[$i].id -as [int]) -eq $targetId ) { $targetIndex = $i; break }
    }

    function ConvertTo-JsonArray {
        param([System.Collections.IList]$ops)
        if ($ops.Count -gt 1) { return ($ops | ConvertTo-Json -Depth 6) }
        $single = $ops[0] | ConvertTo-Json -Depth 6
        return ("[" + $single + "]")
    }

    try {
        if ($UseJsonPatch) {
            $ops = New-Object System.Collections.ArrayList
            if ($targetIndex -ge 0) {
                [void]$ops.Add(@{ op = "replace"; path = "/customFields/$targetIndex/value"; value = $DepartmentValue })
            } else {
                [void]$ops.Add(@{ op = "add"; path = "/customFields/-"; value = @{ id = $targetId; value = $DepartmentValue } })
            }
            $patch   = ConvertTo-JsonArray -ops $ops
            $headers = Get-CwHeaders -ContentType 'application/json'
            LogDebug ("CW JSON Patch body: " + $patch)
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
                LogDebug ("CW object replace body: " + $body)
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
            LogDebug ("CW object replace body: " + $body)
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
    param(
        [int]$TicketId,
        [string]$Text,
        [bool]$InternalAnalysisFirst = $true
    )
    $headers = Get-CwHeaders -ContentType 'application/json; charset=utf-8'
    $url     = "https://${script:CwServer}/v4_6_release/apis/3.0/service/tickets/$TicketId/notes"

    $memberId    = [Environment]::GetEnvironmentVariable('ConnectWisePsa_MemberId')
    $memberIdent = [Environment]::GetEnvironmentVariable('ConnectWisePsa_MemberIdentifier')
    function AttachMember {
        param([hashtable]$payload)
        if ($memberId -or $memberIdent) {
            $member = @{}
            if ($memberId)    { $member.id        = ($memberId -as [int]) }
            if ($memberIdent) { $member.identifier = $memberIdent }
            $payload.member = $member
        }
        return $payload
    }

    function New-NotePayload {
        param([bool]$internal,[bool]$discussion,[bool]$resolution,[string]$text)
        $maxLen = 2000
        if ($text.Length -gt $maxLen) { $text = $text.Substring(0,$maxLen) + "`n(truncated)" }
        $payload = @{
            ticketId              = $TicketId
            text                  = $text
            internalAnalysisFlag  = $internal
            detailDescriptionFlag = $discussion
            resolutionFlag        = $resolution
            externalFlag          = $false
            customerUpdatedFlag   = $false
        }
        $payload = AttachMember -payload $payload
        return ($payload | ConvertTo-Json -Depth 5)
    }

    $body1 = New-NotePayload -internal $InternalAnalysisFirst -discussion $false -resolution $false -text $Text
    LogInfo  ("CW Add Note: TicketId=$TicketId, internalAnalysisFlag=$InternalAnalysisFirst")
    LogDebug ("CW Note body (truncated): " + $body1.Substring(0, [Math]::Min(600, $body1.Length)))
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body1 -ErrorAction Stop
        LogInfo ("CW Add Note -> OK")
        return $true
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW Add Note failed: " + $_.Exception.Message + $suffix)
    }

    $body2 = New-NotePayload -internal $false -discussion $true -resolution $false -text $Text
    LogInfo  ("CW Add Note (Discussion): TicketId=$TicketId")
    LogDebug ("CW Note body (truncated): " + $body2.Substring(0, [Math]::Min(600, $body2.Length)))
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body2 -ErrorAction Stop
        LogInfo ("CW Add Note (Discussion) -> OK")
        return $true
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW Add Note (Discussion) failed: " + $_.Exception.Message + $suffix)
    }

    $body3 = New-NotePayload -internal $false -discussion $false -resolution $true -text $Text
    LogInfo  ("CW Add Note (Resolution): TicketId=$TicketId")
    LogDebug ("CW Note body (truncated): " + $body3.Substring(0, [Math]::Min(600, $body3.Length)))
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body3 -ErrorAction Stop
        LogInfo ("CW Add Note (Resolution) -> OK")
        return $true
    } catch {
        $errBody = Get-CwErrorBody -ex $_.Exception
        $suffix  = $errBody ? (" | Body: " + $errBody) : ""
        LogError ("CW Add Note (Resolution) failed: " + $_.Exception.Message + $suffix)
        return $false
    }
}

# ---------------------------
# Optional SecurityKey validation (header OR query)
# ---------------------------
$secKey = [Environment]::GetEnvironmentVariable('SecurityKey')
$reqKey = $null
if ($Request -and $Request.PSObject.Properties['Headers']) { $reqKey = $Request.Headers['SecurityKey'] }
if ($secKey -and (-not $reqKey) -and $Request -and $Request.PSObject.Properties['Query']) { $reqKey = $Request.Query['securityKey'] }
if ($secKey) {
    if (-not $reqKey -or $reqKey -ne $secKey) {
        LogError "Invalid or missing SecurityKey"
        New-JsonResponse -Code 401 -Message "Invalid or missing SecurityKey"; return
    }
}

# ---------------------------
# Inputs from request (CloudRadial or CW callback)
# ---------------------------
$body = Get-RequestBodyObject -Request $Request
function Get-Prop { param([object]$obj,[string]$name) if ($null -eq $obj) { return $null } $p = $obj.PSObject.Properties[$name]; if ($null -ne $p) { return $p.Value }; return $null }

# TicketId (CloudRadial, CW, query/header/body)
[int]$TicketId = 0
$TicketId = (Get-Prop $body 'TicketId') -as [int]
if (-not $TicketId) {
    $ticketObj = Get-Prop $body 'Ticket'
    if ($ticketObj) { $TicketId = (Get-Prop $ticketObj 'TicketId') -as [int] }
}
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Query'])   { $TicketId = ($Request.Query['ticketId'] -as [int]) }
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Headers']) { $TicketId = ($Request.Headers['TicketId'] -as [int]) }
# CW common (?id) and body.ID/Entity.id
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Query']) { $TicketId = ($Request.Query['id'] -as [int]) }
if (-not $TicketId) {
    $CwId   = Get-Prop $body 'ID'
    $Entity = Get-Prop $body 'Entity'
    if ($CwId) { $TicketId = ($CwId -as [int]) }
    elseif ($Entity -and $Entity.PSObject.Properties['id']) { $TicketId = ($Entity.id -as [int]) }
}
if (-not $TicketId) {
    LogError "TicketId not found in body/query/header"
    New-JsonResponse -Code 400 -Message "TicketId is required (body.TicketId, body.Ticket.TicketId, query ?ticketId or ?id, or header TicketId)."; return
}

# Resolve CW ticket (for PSA Company Id & fallbacks)
$ticket = Get-CwTicket -TicketId $TicketId
if (-not $ticket) { New-JsonResponse -Code 500 -Message "Unable to read CW ticket."; return }

# TenantId: prefer direct inputs, else CloudRadial API lookup by PSA Company Id -> @CompanyTenantId
$TenantId       = $null
$TenantIdSource = $null
$TenantId = Get-Prop $body 'TenantId'
if ($TenantId) { $TenantIdSource = 'Body.TenantId' }
if (-not $TenantId -and $Request -and $Request.PSObject.Properties['Query']) { $TenantId = $Request.Query['tenantId']; if ($TenantId) { $TenantIdSource = 'Query.tenantId' } }
if (-not $TenantId -and $Request -and $Request.PSObject.Properties['Headers']) { $TenantId = $Request.Headers['TenantId']; if ($TenantId) { $TenantIdSource = 'Header.TenantId' } }

if (-not $TenantId) {
    $cwCompanyId = $null
    if ($ticket.PSObject.Properties['company'] -and $ticket.company.PSObject.Properties['id']) { $cwCompanyId = ($ticket.company.id -as [int]) }
    elseif ($ticket.PSObject.Properties['company'] -and $ticket.company.PSObject.Properties['identifier']) { $cwCompanyId = ($ticket.company.identifier -as [int]) }
    if ($cwCompanyId) {
        $crCompany  = Get-CloudRadialCompanyByPsaId -PsaCompanyId $cwCompanyId
        $crTenantId = Get-CloudRadialTenantIdFromCompany -Company $crCompany
        if ($crCompany -and $crTenantId) {
            $TenantId       = $crTenantId
            $TenantIdSource = "CloudRadial.CompanyToken(@CompanyTenantId)"
            LogInfo ("Resolved TenantId via CloudRadial API: $TenantId (PSA CompanyId=$cwCompanyId)")
        } else {
            LogInfo ("CloudRadial did not return @CompanyTenantId for PSAId=$cwCompanyId")
        }
    } else {
        LogInfo ("CW ticket has no company.id; skipping CloudRadial lookup.")
    }
}

# Partner tenant guardrails
$PartnerTenantId = [Environment]::GetEnvironmentVariable('Ms365_TenantId')
if ([string]::IsNullOrWhiteSpace($TenantId)) { LogError "TenantId missing; CloudRadial API lookup failed or not configured."; New-JsonResponse -Code 400 -Message "TenantId is required; ensure CloudRadial has @CompanyTenantId or pass TenantId via body/query/header."; return }
if ($PartnerTenantId -and $TenantId -eq $PartnerTenantId) { LogError "TenantId equals partner Ms365_TenantId; refusing client lookup."; New-JsonResponse -Code 400 -Message "Client TenantId required; partner tenant not allowed."; return }

# UPN/Email
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
    if ($ticket.PSObject.Properties['contact'] -and $ticket.contact.PSObject.Properties['email']) {
        $UserEmail = $ticket.contact.email
    } elseif ($ticket.PSObject.Properties['companyContact'] -and $ticket.companyContact.PSObject.Properties['email']) {
        $UserEmail = $ticket.companyContact.email
    }
}

LogInfo ("Inputs: TicketId=$TicketId, TenantId=$TenantId (source=$TenantIdSource), UPN='$UserUPN', Email='$UserEmail'")

# Connect to Graph & get department
if (-not (Connect-GraphApp -TenantId $TenantId)) {
    New-JsonResponse -Code 500 -Message "Failed to connect to Microsoft Graph" -Extra @{ TicketId=$TicketId; TenantId=$TenantId; TenantIdSource=$TenantIdSource }; return
}
$department = Get-UserDepartment -UserPrincipalName $UserUPN -UserEmail $UserEmail

# Fallback via CW Contact UDF #53 (Client Department)
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
                $cfIdInt = ($cf.id -as [int])
                $cap     = ("" + $cf.caption).Trim()
                if ($cfIdInt -eq 53 -or $cap -eq "Client Department") {
                    $department = $cf.value; $source = "CWContactUDF53"; $fallbackContactId = $contactId; break
                }
            }
        }
    }
}

if (-not $department) { $department = "" }
LogInfo ("Final department (source=$source) = '" + (""+$department) + "'")

# Optional: skip updating if department is blank
$SkipEmpty = ([Environment]::GetEnvironmentVariable('SkipEmptyDepartment') -as [int]) -eq 1
if ($SkipEmpty -and [string]::IsNullOrWhiteSpace($department)) {
    LogInfo "SkipEmptyDepartment is ON; department is blank—skipping PATCH."
    $ok = $false
} else {
    $ok = Set-CwTicketDepartmentCustomField -TicketId $TicketId -DepartmentValue $department
}

# Verify UDF #54
$verifyUdfValue = $null
try {
    $verify = Get-CwTicket -TicketId $TicketId
    if ($verify -and $verify.customFields) {
        foreach ($cf in @($verify.customFields)) {
            $cfIdInt = ($cf.id -as [int]); $cap = ("" + $cf.caption).Trim()
            if ($cfIdInt -eq 54 -or $cap -eq "Client Department") { $verifyUdfValue = $cf.value; break }
        }
    }
    LogInfo ("Verify UDF #54 after PATCH: '" + (""+$verifyUdfValue) + "'")
} catch { LogError ("Verify after PATCH failed: " + $_.Exception.Message) }

# Audit note
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

# Respond
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
