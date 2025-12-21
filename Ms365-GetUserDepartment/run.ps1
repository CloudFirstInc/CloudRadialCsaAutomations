
#Created by DCG using CoPilot

using namespace System.Net

param($Request, $TriggerMetadata)

# Be strict and fail fast on uninitialized vars
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------
# Common helpers
# ---------------------------
function New-JsonResponse {
    param(
        [int]$Code,
        [string]$Message,
        [hashtable]$Extra = @{}
    )

    $body = @{
        Message      = $Message
        ResultCode   = $Code
        ResultStatus = if ($Code -ge 200 -and $Code -lt 300) { "Success" } else { "Failure" }
    } + $Extra

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]$Code
        Body       = $body
        Headers    = @{ "Content-Type" = "application/json" }
    })
}

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
        } catch {
            return @{}
        }
    }

    if ($raw -is [string]) {
        try {
            if ($raw.Trim().StartsWith('{')) {
                return $raw | ConvertFrom-Json -ErrorAction Stop
            }
        } catch { }
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

        if (-not $TenantId) { $TenantId = [Environment]::GetEnvironmentVariable('Ms365_TenantId') }

        if ([string]::IsNullOrWhiteSpace($appId))     { throw "Missing Microsoft Graph app setting: Ms365_AuthAppId" }
        if ([string]::IsNullOrWhiteSpace($appSecret)) { throw "Missing Microsoft Graph app setting: Ms365_AuthSecretId" }
        if ([string]::IsNullOrWhiteSpace($TenantId))  { throw "Missing Microsoft Graph TenantId (body or Ms365_TenantId)" }

        Connect-MgGraph -TenantId $TenantId -ClientId $appId -ClientSecret $appSecret -NoWelcome -ErrorAction Stop
        return $true
    } catch {
        Write-Error "Graph connect failed: $($_.Exception.Message)"
        return $false
    }
}

function Get-UserDepartment {
    param([string]$UserPrincipalName,[string]$UserEmail)

    $user = $null

    if ($UserPrincipalName) {
        try {
            $user = Get-MgUser -UserId $UserPrincipalName -Property department,userPrincipalName -ErrorAction Stop
        } catch {
            $user = $null
        }
    }

    if (-not $user -and $UserEmail) {
        try {
            $safeEmail = $UserEmail.Replace("'","''")
            $filter    = "mail eq '$safeEmail' or userPrincipalName eq '$safeEmail'"
            $user      = Get-MgUser -Filter $filter -Property department,userPrincipalName -Top 1 -ErrorAction Stop
        } catch {
            $user = $null
        }
    }

    if ($user) { return $user.Department }
    return $null
}

# ---------------------------
# ConnectWise helpers (uses your ConnectWisePsa_* env vars)
# ---------------------------
$script:CwServer = 'api-na.myconnectwise.net'

function Get-CwHeaders {
    $required = 'ConnectWisePsa_ApiCompanyId','ConnectWisePsa_ApiPublicKey','ConnectWisePsa_ApiPrivateKey','ConnectWisePsa_ApiClientId'
    foreach ($n in $required) {
        $val = [Environment]::GetEnvironmentVariable($n)
        if ([string]::IsNullOrWhiteSpace($val)) { throw "Missing ConnectWise app setting: $n" }
    }

    $companyId = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiCompanyId')
    $pubKey    = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiPublicKey')
    $privKey   = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiPrivateKey')
    $clientId  = [Environment]::GetEnvironmentVariable('ConnectWisePsa_ApiClientId')

    # Delimit variables in interpolated string to avoid $var: parsing issues
    $authString  = "${companyId}+${pubKey}:${privKey}"
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))

    return @{
        "Authorization" = "Basic $encodedAuth"
        "ClientID"      = $clientId
        "Content-Type"  = "application/json; charset=utf-8"
        "Accept"        = "application/vnd.connectwise.com+json; version=2022.1"
    }
}

function Get-CwTicket {
    param([int]$TicketId)
    $headers = Get-CwHeaders
    $url     = "https://$($script:CwServer)/v4_6_release/apis/3.0/service/tickets/$TicketId"
    try {
        Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
    } catch {
        Write-Error "CW GET ticket failed: $($_.Exception.Message)"
        return $null
    }
}

function Get-CwContactById {
    param([int]$ContactId)
    $headers = Get-CwHeaders
    $url     = "https://$($script:CwServer)/v4_6_release/apis/3.0/company/contacts/$ContactId"
    try {
        Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
    } catch {
        Write-Error "CW GET contact failed: $($_.Exception.Message)"
        return $null
    }
}

function Get-FallbackDepartmentFromContactUdf {
    param([object]$Ticket,[int]$ContactUdfId,[string]$ContactUdfCaption)

    $contactId = $null
    if     ($Ticket.contact       -and $Ticket.contact.id)         { $contactId = [int]$Ticket.contact.id }
    elseif ($Ticket.companyContact -and $Ticket.companyContact.id) { $contactId = [int]$Ticket.companyContact.id }
    elseif ($Ticket.contactId)                                     { $contactId = [int]$Ticket.contactId }

    if (-not $contactId) { return @{ Value=$null; ContactId=$null } }

    $contact = Get-CwContactById -ContactId $contactId
    if (-not $contact) { return @{ Value=$null; ContactId=$contactId } }

    $Normalize = { param($s) ("" + $s).Trim() }

    foreach ($cf in @($contact.customFields)) {
        $cfIdInt       = ($cf.id -as [int])
        $cfCaptionNorm = & $Normalize $cf.caption
        if ( ($cfIdInt -eq 53) -or ($cfCaptionNorm -eq "Client Department") ) {
            return @{ Value=$cf.value; ContactId=$contactId }
        }
    }

    return @{ Value=$null; ContactId=$contactId }
}

function Set-CwTicketDepartmentCustomField {
    param([int]$TicketId,[string]$DepartmentValue)

    $headers = Get-CwHeaders
    $url     = "https://$($script:CwServer)/v4_6_release/apis/3.0/service/tickets/$TicketId"

    $ticket = Get-CwTicket -TicketId $TicketId
    if (-not $ticket) { return $false }

    $targetId      = 54
    $targetCaption = "Client Department"

    $customFields = @()
    if ($ticket.customFields) { $customFields = @($ticket.customFields) }

    $found = $false
    for ($i = 0; $i -lt $customFields.Count; $i++) {
        $cf = $customFields[$i]
        $cfIdInt       = ($cf.id -as [int])
        $cfCaptionNorm = ("" + $cf.caption).Trim()
        if ($cfIdInt -eq $targetId -or $cfCaptionNorm -eq $targetCaption) {
            $customFields[$i].value = $DepartmentValue
            $found = $true
        }
    }

    if (-not $found) {
        $cfObj = @{
            id               = $targetId
            caption          = $targetCaption
            type             = "Text"
            entryMethod      = "EntryField"
            numberOfDecimals = 0
            value            = $DepartmentValue
        }
        $customFields += $cfObj
    }

    $patchBody = @{ customFields = $customFields } | ConvertTo-Json -Depth 6

    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $patchBody -ErrorAction Stop
        Write-Host "CW PATCH customFields -> OK (TicketId=$TicketId, UDF=54)"
        return $true
    } catch {
        Write-Error "CW PATCH customFields failed: $($_.Exception.Message)"
        return $false
    }
}

function Add-CwTicketNote {
    param(
        [int]$TicketId,
        [string]$Text,
        [bool]$InternalFlag = $true,
        [bool]$DetailDescriptionFlag = $false,
        [bool]$ResolutionFlag = $false
    )
    $headers = Get-CwHeaders
    $url     = "https://$($script:CwServer)/v4_6_release/apis/3.0/service/tickets/$TicketId/notes"

    $noteBody = @{
        ticketId              = $TicketId
        text                  = $Text
        internalFlag          = $InternalFlag
        detailDescriptionFlag = $DetailDescriptionFlag
        resolutionFlag        = $ResolutionFlag
    } | ConvertTo-Json -Depth 4

    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $noteBody -ErrorAction Stop
        Write-Host "CW Add Note -> OK (TicketId=$TicketId)"
        return $true
    } catch {
        Write-Error "CW Add Note failed: $($_.Exception.Message)"
        return $false
    }
}

# ---------------------------
# Optional SecurityKey validation
# ---------------------------
$secKey = [Environment]::GetEnvironmentVariable('SecurityKey')
if ($secKey) {
    $reqKey = $Request.Headers['SecurityKey']
    if (-not $reqKey -or $reqKey -ne $secKey) {
        New-JsonResponse -Code 401 -Message "Invalid or missing SecurityKey"; return
    }
}

# ---------------------------
# Inputs from CloudRadial webhook
# ---------------------------
$body = Get-RequestBodyObject -Request $Request

[int]$TicketId = $body.TicketId
if (-not $TicketId -and $body.Ticket -and $body.Ticket.TicketId) { [int]$TicketId = $body.Ticket.TicketId }
if (-not $TicketId) { New-JsonResponse -Code 400 -Message "TicketId is required"; return }

$TenantId = $body.TenantId
if (-not $TenantId) { $TenantId = [Environment]::GetEnvironmentVariable('Ms365_TenantId') }

$UserUPN   = $body.UserOfficeId
$UserEmail = $body.UserEmail
if (-not $UserUPN   -and $body.User) { $UserUPN   = $body.User.UserOfficeId }
if (-not $UserEmail -and $body.User) { $UserEmail = $body.User.Email }

# ---------------------------
# Connect to Graph and get Department
# ---------------------------
if (-not (Connect-GraphApp -TenantId $TenantId)) {
    New-JsonResponse -Code 500 -Message    New-JsonResponse -Code 500 -Message "Failed to connect to Microsoft Graph"; return
}
$department = Get-UserDepartment -UserPrincipalName $UserUPN -UserEmail $UserEmail

# ---------------------------
# Fallback: if Department is blank, read CW contact UDF (ID 53)
# ---------------------------
$source = "EntraID"
$fallbackContactId = $null

if ([string]::IsNullOrWhiteSpace($department)) {
    $ticketObj = Get-CwTicket -TicketId $TicketId
    if ($ticketObj) {
        $fb = Get-FallbackDepartmentFromContactUdf -Ticket $ticketObj -ContactUdfId 53 -ContactUdfCaption "Client Department"
        if ($fb.Value) {
            $department        = $fb.Value
            $source            = "CWContactUDF53"
            $fallbackContactId = $fb.ContactId
        }
    }
}

if (-not $department) { $department = "" }

# ---------------------------
# Update CW ticket UDF #54 (Client Department) with the final value
# ---------------------------
$ok = Set-CwTicketDepartmentCustomField -TicketId $TicketId -DepartmentValue $department

# ---------------------------
# Audit logging note (avoid here-string to simplify parsing)
# ---------------------------
$timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss zzz")
$noteLines = @(
    "**Department sync audit**",
    "- Source: $source",
    "- Applied value: '$department'",
    "- Submitter UPN: '$UserUPN'",
    "- Submitter Email: '$UserEmail'",
    "- Fallback ContactId (if used): '$fallbackContactId'",
    "- Timestamp: $timestamp"
)
$noteText = ($noteLines -join [Environment]::NewLine)

$noteOk = Add-CwTicketNote -TicketId $TicketId -Text $noteText -InternalFlag $true

# ---------------------------
# Respond to caller
# ---------------------------
if ($ok) {
    New-JsonResponse -Code 200 -Message "Updated CW ticket UDF #54 (Client Department) and logged audit note." -Extra @{
        TicketId   = $TicketId
        Department = $department
        Source     = $source
        NoteAdded  = $noteOk
    }
}
else {
    New-JsonResponse -Code 500 -Message "Failed to update CW ticket UDF #54 (Client Department). Audit note was attempted." -Extra @{
        TicketId   = $TicketId
        Department = $department
        Source     = $source
        NoteAdded  = $noteOk
    }
