
# Ms365-GetUserDepartment/run.ps1
# Updated: ticket-level UDF #54 update via JSON Patch (array, leading slashes, application/json),
# per-call Content-Type, Add Note fallback (Internal -> Discussion), StrictMode-safe CW,
# robust error-body capture, email-first Graph lookup, forced client tenant, no inline 'if'.

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
function LogError { param([string]$msg) Write-Error ("[$CorrelationId] " + $msg) -ErrorAction Continue }
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
    # 1) Prefer email
    if ($UserEmail) {
        try {
            $safeEmail = $UserEmail.Replace("'","''")
            $filter    = "mail eq '$safeEmail' or userPrincipalName eq '$safeEmail'"
            LogInfo ("Graph lookup by filter: $filter")
            $user      = Get-MgUser -Filter $filter -Property department,userPrincipalName -Top 1 -ErrorAction Stop
        } catch { LogError ("Get-MgUser by email filter failed: " + $_.Exception.Message) }
    }
    # 2) Fallback: by Id/UPN (GUID or UPN)
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
$script:CwServer  = 'api-na.myconnectwise.net'
$UseJsonPatch     = ([Environment]::GetEnvironmentVariable('ConnectWise_UseJsonPatch') -as [int]) -eq 1

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
    $url     = "https://$($script:CwServer)/v4_6_release/apis/3.0/service/tickets/$TicketId"
    LogInfo ("CW GET ticket: id=$TicketId, url=$url")
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
        LogInfo ("CW GET ticket -> OK")
        $hasCustom    = ($resp.PSObject.Properties['customFields'] -ne $null)
        $contactIdStr = ""
        if     ($resp.PSObject.Properties['contactId'])                                            { $contactIdStr = "" + $resp.contactId }
        elseif ($resp.PSObject.Properties['contact'] -and $resp.contact.PSObject.Properties['id']) { $contactIdStr = "" + $resp.contact.id }
        elseif ($resp.PSObject.Properties['companyContact'] -and $resp.companyContact.PSObject.Properties['id']) { $contactIdStr = "" + $resp.companyContact.id }
        LogDebug ("CW GET ticket fields: has customFields=$hasCustom, contactId=$contactIdStr")
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
    $url     = "https://$($script:CwServer)/v4_6_release/apis/3.0/company/contacts/$ContactId"
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

function Get-FallbackDepartmentFromContactUdf {
    param([object]$Ticket,[int]$ContactUdfId,[string]$ContactUdfCaption)

    # Precompute values (no inline 'if' inside strings)
    $cap_contactId        = ""
    $cap_companyContactId = ""
    $cap_contactIdLegacy  = ""

    if ($Ticket.PSObject.Properties['contact'] -and $Ticket.contact.PSObject.Properties['id']) { $cap_contactId = "" + $Ticket.contact.id }
    if ($Ticket.PSObject.Properties['companyContact'] -and $Ticket.companyContact.PSObject.Properties['id']) { $cap_companyContactId = "" + $Ticket.companyContact.id }
    if ($Ticket.PSObject.Properties['contactId']) { $cap_contactIdLegacy = "" + $Ticket.contactId }

    $caps = @("contact.id=$cap_contactId","companyContact.id=$cap_companyContactId","contactId=$cap_contactIdLegacy")
    LogDebug ("CW ticket contact candidates: " + ($caps -join ", "))

    $contactId = $null
    if     ($cap_contactId)        { $contactId = [int]$cap_contactId }
    elseif ($cap_companyContactId) { $contactId = [int]$cap_companyContactId }
    elseif ($cap_contactIdLegacy)  { $contactId = [int]$cap_contactIdLegacy }

    if (-not $contactId) { return @{ Value=$null; ContactId=$null } }

    $contact = Get-CwContactById -ContactId $contactId
    if (-not $contact) { return @{ Value=$null; ContactId=$contactId } }

    $Normalize = { param([object]$s) ("" + $s).Trim() }
    foreach ($cf in @($contact.customFields)) {
        $cfIdInt       = ($cf.id -as [int])
        $cfCaptionNorm = & $Normalize $cf.caption
        if ( ($cfIdInt -eq $ContactUdfId) -or ($cfCaptionNorm -eq $ContactUdfCaption) ) {
            LogInfo ("Fallback department from contact UDF (id=$ContactUdfId, caption='$ContactUdfCaption') = '" + (""+$cf.value) + "'")
            return @{ Value=$cf.value; ContactId=$contactId }
        }
    }
    return @{ Value=$null; ContactId=$contactId }
}

# ---------------------------
# Ticket-level UDF #54 updater
# ---------------------------
function Set-CwTicketDepartmentCustomField {
    param([int]$TicketId,[string]$DepartmentValue)

    $url = "https://$($script:CwServer)/v4_6_release/apis/3.0/service/tickets/$TicketId"

    # Read ticket to find UDF #54 index
    $ticket = Get-CwTicket -TicketId $TicketId
    if (-not $ticket) { return $false }

    $targetId    = 54
    $existing    = @()
    if ($ticket.customFields) { $existing = @($ticket.customFields) }

    $targetIndex = -1
    for ($i = 0; $i -lt $existing.Count; $i++) {
        if ( ($existing[$i].id -as [int]) -eq $targetId ) { $targetIndex = $i; break }
    }

    # Helper: force JSON array even when there's a single op
    function ConvertTo-JsonArray {
        param([System.Collections.IList]$ops)
        if ($ops.Count -gt 1) {
            return ($ops | ConvertTo-Json -Depth 6)
        } else {
            # Single element -> wrap explicitly in [...]
            $single = $ops[0] | ConvertTo-Json -Depth 6
            return ("[" + $single + "]")
        }
    }

    try {
        if ($UseJsonPatch) {
            # RFC-6902 ARRAY with leading slash paths; CW prefers application/json
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
                # Log, then FALL BACK to object-replace full array
                $errBody = Get-CwErrorBody -ex $_.Exception
                $suffix  = $errBody ? (" | Body: " + $errBody) : ""
                LogError ("CW PATCH (JSON Patch) failed: " + $_.Exception.Message + $suffix)
                LogInfo  ("Falling back to full-array object replace for customFields")

                $headers      = Get-CwHeaders -ContentType 'application/json'
                $customFields = @()
                if ($existing.Count -gt 0) { $customFields = @($existing) }
                if ($targetIndex -ge 0) {
                    $customFields[$targetIndex].value = $DepartmentValue
                } else {
                    $customFields += @{ id = $targetId; value = $DepartmentValue }
                }

                $body = @{ customFields = $customFields } | ConvertTo-Json -Depth 6
                LogDebug ("CW object replace body: " + $body)
                $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $body -ErrorAction Stop
                LogInfo ("CW PATCH (object replace) customFields -> OK")
                return $true
            }
        } else {
            # Object-replace the full array
            $headers = Get-CwHeaders -ContentType 'application/json'
            $customFields = @()
            if ($existing.Count -gt 0) { $customFields = @($existing) }
            if ($targetIndex -ge 0) {
                $customFields[$targetIndex].value = $DepartmentValue
            } else {
                $customFields += @{ id = $targetId; value = $DepartmentValue }
            }
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
# Add Note with fallback: Internal -> Discussion
# ---------------------------



function Add-CwTicketNote {
    param(
        [int]$TicketId,
        [string]$Text,
        [bool]$InternalAnalysisFirst = $true  # first attempt: Internal (internalAnalysisFlag)
    )

    # Prefer charset for CW note handling
    $headers = Get-CwHeaders -ContentType 'application/json; charset=utf-8'
    $url     = "https://$($script:CwServer)/v4_6_release/apis/3.0/service/tickets/$TicketId/notes"

    # Optional: include a member on the note if provided in app settings
    $memberId   = [Environment]::GetEnvironmentVariable('ConnectWisePsa_MemberId')
    $memberIdent= [Environment]::GetEnvironmentVariable('ConnectWisePsa_MemberIdentifier')
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

    # Helper: safely build a serviceNote with exactly one flag ON
    function New-NotePayload {
        param([bool]$internal,[bool]$discussion,[bool]$resolution,[string]$text)
        $maxLen = 2000
        if ($text.Length -gt $maxLen) { $text = $text.Substring(0,$maxLen) + "`n(truncated)" }

        $payload = @{
            ticketId              = $TicketId
            text                  = $text

            # Exactly one of the three flags must be true on create:
            internalAnalysisFlag  = $internal
            detailDescriptionFlag = $discussion
            resolutionFlag        = $resolution

            # Safe defaults
            externalFlag          = $false
            customerUpdatedFlag   = $false
        }
        $payload = AttachMember -payload $payload
        return ($payload | ConvertTo-Json -Depth 5)
    }

    # Attempt 1: Internal (internalAnalysisFlag=true)
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

    # Attempt 2: Discussion (detailDescriptionFlag=true)
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

    # Attempt 3: Resolution (resolutionFlag=true)
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
# Optional SecurityKey validation
# ---------------------------
$secKey = [Environment]::GetEnvironmentVariable('SecurityKey')
if ($secKey) {
    $reqKey = $null
    if ($Request -and $Request.PSObject.Properties['Headers']) { $reqKey = $Request.Headers['SecurityKey'] }
    if (-not $reqKey -or $reqKey -ne $secKey) {
        LogError "Invalid or missing SecurityKey"
        New-JsonResponse -Code 401 -Message "Invalid or missing SecurityKey"; return
    }
}

# ---------------------------
# Inputs from CloudRadial webhook
# ---------------------------
$body = Get-RequestBodyObject -Request $Request
function Get-Prop { param([object]$obj,[string]$name) if ($null -eq $obj) { return $null } $p = $obj.PSObject.Properties[$name]; if ($null -ne $p) { return $p.Value }; return $null }

# TicketId
[int]$TicketId = 0
$TicketId = (Get-Prop $body 'TicketId') -as [int]
if (-not $TicketId) {
    $ticketObj = Get-Prop $body 'Ticket'
    if ($ticketObj) { $TicketId = (Get-Prop $ticketObj 'TicketId') -as [int] }
}
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Query'])   { $TicketId = ($Request.Query['ticketId'] -as [int]) }
if (-not $TicketId -and $Request -and $Request.PSObject.Properties['Headers']) { $TicketId = ($Request.Headers['TicketId'] -as [int]) }
if (-not $TicketId) { LogError "TicketId not found in body/query/header"; New-JsonResponse -Code 400 -Message "TicketId is required. Provide it in body.TicketId, body.Ticket.TicketId, query ?ticketId, or header TicketId."; return }

# TenantId (forced client; no partner fallback)
$TenantId       = $null
$TenantIdSource = $null
$TenantId = Get-Prop $body 'TenantId'
if ($TenantId) { $TenantIdSource = 'Body.TenantId' }
if (-not $TenantId -and $Request -and $Request.PSObject.Properties['Query']) { $TenantId = $Request.Query['tenantId']; if ($TenantId) { $TenantIdSource = 'Query.tenantId' } }
if (-not $TenantId -and $Request -and $Request.PSObject.Properties['Headers']) { $TenantId = $Request.Headers['TenantId']; if ($TenantId) { $TenantIdSource = 'Header.TenantId' } }
$PartnerTenantId = [Environment]::GetEnvironmentVariable('Ms365_TenantId')
if ([string]::IsNullOrWhiteSpace($TenantId)) { LogError "TenantId missing; refusing to use partner Ms365_TenantId. Pass @CompanyTenantId."; New-JsonResponse -Code 400 -Message "TenantId is required (CloudRadial token @CompanyTenantId)."; return }
if ($PartnerTenantId -and $TenantId -eq $PartnerTenantId) { LogError "TenantId equals partner Ms365_TenantId; refusing client lookup. Ensure CloudRadial is sending @CompanyTenantId."; New-JsonResponse -Code 400 -Message "Client TenantId required; partner tenant not allowed."; return }

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
LogInfo ("Inputs: TicketId=$TicketId, TenantId=$TenantId (source=$TenantIdSource), UPN='$UserUPN', Email='$UserEmail'")

# Connect to Graph & get department
if (-not (Connect-GraphApp -TenantId $TenantId)) {
    New-JsonResponse -Code 500 -Message "Failed to connect to Microsoft Graph" -Extra @{ TicketId=$TicketId; TenantId=$TenantId; TenantIdSource=$TenantIdSource }; return
}
$department = Get-UserDepartment -UserPrincipalName $UserUPN -UserEmail $UserEmail

# Fallback via CW Contact UDF #53
$source = "EntraID"; $fallbackContactId = $null
if ([string]::IsNullOrWhiteSpace($department)) {
    LogInfo "Department blank from Graph; attempting CW contact UDF fallback (id=53)"
    $ticketForFb = Get-CwTicket -TicketId $TicketId
    if ($ticketForFb) {
        $fb = Get-FallbackDepartmentFromContactUdf -Ticket $ticketForFb -ContactUdfId 53 -ContactUdfCaption "Client Department"
        if ($fb.Value) { $department = $fb.Value; $source = "CWContactUDF53"; $fallbackContactId = $fb.ContactId }
        else { LogInfo "No department in UDF #53 on contact." }
    } else { LogError "Unable to read ticket for fallback." }
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

# Read-after-write verification (ticket-level UDF #54)
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
    ("- Fallback ContactId (if used): '{0}'" -f $fallbackContactId),
    ("- Verify UDF #54 after PATCH: '{0}'" -f $verifyUdfValue),
    ("- Timestamp: {0}" -f $timestamp),
    ("- CorrelationId: {0}" -f $CorrelationId)
)
$noteText = ($noteLines -join [Environment]::NewLine)

$noteOk = Add-CwTicketNote -TicketId $TicketId -Text $noteText -InternalFlag $true

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
