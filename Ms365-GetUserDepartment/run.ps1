
#Created by DCG using CoPilot

using namespace System.Net

param($Request, $TriggerMetadata)

# ---------------------------
# Common helpers
# ---------------------------
function New-JsonResponse {
    param([int]$Code,[string]$Message,[hashtable]$Extra=@{})
    $body=@{Message=$Message;ResultCode=$Code;ResultStatus=if($Code -ge 200 -and $Code -lt 300){"Success"}else{"Failure"}}+$Extra
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::($Code)
        Body        = $body
        ContentType = "application/json"
    })
}

function Connect-GraphApp {
    param([string]$TenantId)
    try{
        $secureSecret = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
        $cred         = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId,$secureSecret)
        Connect-MgGraph -ClientSecretCredential $cred -TenantId $TenantId -ErrorAction Stop
        return $true
    }catch{
        Write-Error "Graph connect failed: $($_.Exception.Message)"
        return $false
    }
}

function Get-UserDepartment {
    param([string]$UserPrincipalName,[string]$UserEmail)
    try{
        $user=$null
        if($UserPrincipalName){
            # Department is not returned by default; explicitly select it ($select).
            $user=Get-MgUser -UserId $UserPrincipalName -Property department,displayName,userPrincipalName -ErrorAction Stop
        }elseif($UserEmail){
            $user=Get-MgUser -Search "mail:$UserEmail" -Property department,displayName,userPrincipalName -Top 1 -ErrorAction Stop
        }
        if(-not $user){return $null}
        return $user.Department
    }catch{
        Write-Error "Get-UserDepartment failed: $($_.Exception.Message)"
        return $null
    }
}

# ---------------------------
# ConnectWise helpers
# ---------------------------
function Get-CwHeaders {
    # ConnectWise REST: Basic auth with "company+publicKey:privateKey", plus ClientID and Accept version.
    $authString  = "$($env:Cw_Company)+$($env:Cw_PublicKey):$($env:Cw_PrivateKey)"
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
    return @{
        "Authorization" = "Basic $encodedAuth"
        "ClientID"      = $env:Cw_ClientId
        "Content-Type"  = "application/json; charset=utf-8"
        "Accept"        = "application/vnd.connectwise.com+json; version=2022.1"
    }
}

function Get-CwTicket {
    param([int]$TicketId)
    $headers=Get-CwHeaders
    $url="https://$($env:Cw_Server)/v4_6_release/apis/3.0/service/tickets/$TicketId"
    try{ Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop }
    catch{ Write-Error "CW GET ticket failed: $($_.Exception.Message)"; $null }
}

function Get-CwContactById {
    param([int]$ContactId)
    $headers=Get-CwHeaders
    $url="https://$($env:Cw_Server)/v4_6_release/apis/3.0/company/contacts/$ContactId"
    try{ Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop }
    catch{ Write-Error "CW GET contact failed: $($_.Exception.Message)"; $null }
}

function Get-FallbackDepartmentFromContactUdf {
    param([object]$Ticket,[int]$ContactUdfId,[string]$ContactUdfCaption)
    # Try to identify the primary contact on the ticket.
    $contactId = $null
    if($Ticket.contact -and $Ticket.contact.id){ $contactId = [int]$Ticket.contact.id }
    elseif($Ticket.companyContact -and $Ticket.companyContact.id){ $contactId = [int]$Ticket.companyContact.id }
    elseif($Ticket.contactId){ $contactId = [int]$Ticket.contactId }

    if(-not $contactId){ return @{ Value=$null; ContactId=$null } }

    $contact = Get-CwContactById -ContactId $contactId
    if(-not $contact){ return @{ Value=$null; ContactId=$contactId } }

    # Scan the contact's customFields for ID=53 ("Client Department")
    $Normalize = { param($s) (""+$s).Trim() }
    foreach($cf in @($contact.customFields)){
        $cfIdInt = ($cf.id -as [int])
        $cfCaptionNorm = & $Normalize $cf.caption   # <-- fixed
        if( ($cfIdInt -eq 53) -or ($cfCaptionNorm -eq "Client Department") ){
            return @{ Value=$cf.value; ContactId=$contactId }
        }
    }
    return @{ Value=$null; ContactId=$contactId }
}

function Set-CwTicketDepartmentCustomField {
    param([int]$TicketId,[string]$DepartmentValue)

    $headers=Get-CwHeaders
    $url="https://$($env:Cw_Server)/v4_6_release/apis/3.0/service/tickets/$TicketId"

    # 1) Read existing ticket (for current customFields array)
    $ticket=Get-CwTicket -TicketId $TicketId
    if(-not $ticket){ return $false }

    # ---- Target Ticket UDF (Client Department on ticket): hard-coded ID 54 ----
    $targetId      = 54
    $targetCaption = "Client Department"

    # Prepare full array (CW requires replacing entire customFields on PATCH)
    $customFields = @()
    if($ticket.customFields){ $customFields = @($ticket.customFields) }

    $found=$false
    for($i=0; $i -lt $customFields.Count; $i++){
        $cf = $customFields[$i]
        $cfIdInt = ($cf.id -as [int])
        $cfCaptionNorm = (""+$cf.caption).Trim()

        if($cfIdInt -eq $targetId -or $cfCaptionNorm -eq $targetCaption){
            $customFields[$i].value = $DepartmentValue
            $found=$true
        }
    }

    if(-not $found){
        # Append entry if the array didn't include the target UDF yet
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

    $patchBody = @(
        @{
            op    = "replace"
            path  = "customFields"
            value = $customFields
        }
    ) | ConvertTo-Json -Depth 6

    try{
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Patch -Body $patchBody -ErrorAction Stop
        Write-Host "CW PATCH customFields -> OK (TicketId=$TicketId, UDF=54)"
        return $true
    }catch{
        Write-Error "CW PATCH customFields failed: $($_.Exception.Message)"
        return $false
    }
}

# ---------------------------
# NEW: Add audit note to ticket
# ---------------------------
function Add-CwTicketNote {
    param(
        [int]$TicketId,
        [string]$Text,
        [bool]$InternalFlag = $true,
        [bool]$DetailDescriptionFlag = $false,
        [bool]$ResolutionFlag = $false
    )
    $headers = Get-CwHeaders
    $url     = "https://$($env:Cw_Server)/v4_6_release/apis/3.0/service/tickets/$TicketId/notes"

    $noteBody = @{
        ticketId               = $TicketId
        text                   = $Text
        internalFlag           = $InternalFlag
        detailDescriptionFlag  = $DetailDescriptionFlag
        resolutionFlag         = $ResolutionFlag
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
if($env:SecurityKey){
    $reqKey=$Request.Headers.SecurityKey
    if(-not $reqKey -or $reqKey -ne $env:SecurityKey){
        New-JsonResponse -Code 401 -Message "Invalid or missing SecurityKey"; return
    }
}

# ---------------------------
# Inputs from CloudRadial webhook
# ---------------------------
[int]$TicketId = $Request.Body.TicketId
if(-not $TicketId -and $Request.Body.Ticket.TicketId){ [int]$TicketId = $Request.Body.Ticket.TicketId }

$TenantId = $Request.Body.TenantId; if(-not $TenantId){ $TenantId = $env:Ms365_TenantId }

# Prefer UPN (@UserOfficeId), else email (@UserEmail)
$UserUPN   = $Request.Body.UserOfficeId
$UserEmail = $Request.Body.UserEmail
if(-not $UserUPN   -and $Request.Body.User.UserOfficeId){ $UserUPN   = $Request.Body.User.UserOfficeId }
if(-not $UserEmail -and $Request.Body.User.Email){       $UserEmail = $Request.Body.User.Email }

if(-not $TicketId){ New-JsonResponse -Code 400 -Message "TicketId is required"; return }   # <-- fixed

# ---------------------------
# Connect to Graph and get Department
# ---------------------------
if(-not (Connect-GraphApp -TenantId $TenantId)){
    New-JsonResponse -Code 500 -Message "Failed to connect to Microsoft Graph"; return
}
$department = Get-UserDepartment -UserPrincipalName $UserUPN -UserEmail $UserEmail

# ---------------------------
# Fallback: if Department is blank, read CW contact UDF (ID 53)
# ---------------------------
$source = "EntraID"
$fallbackContactId = $null

if([string]::IsNullOrWhiteSpace($department)){
    $ticketObj = Get-CwTicket -TicketId $TicketId
    if($ticketObj){
        $fb = Get-FallbackDepartmentFromContactUdf -Ticket $ticketObj -ContactUdfId 53 -ContactUdfCaption "Client Department"
        if($fb.Value){
            $department = $fb.Value
            $source = "CWContactUDF53"
            $fallbackContactId = $fb.ContactId
        }
    }
}

# Optional: default to empty string if still blank
if(-not $department){ $department = "" }

# ---------------------------
# Update CW ticket UDF #54 (Client Department) with the final value
# ---------------------------
$ok = Set-CwTicketDepartmentCustomField -TicketId $TicketId -DepartmentValue $department

# ---------------------------
# NEW: Audit logging note
# ---------------------------
$timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss zzz")
$noteText = @"
**Department sync audit**
- Source: $source
- Applied value: '$department'
- Submitter UPN: '$UserUPN'
- Submitter Email: '$UserEmail'
- Fallback ContactId (if used): '$fallbackContactId'
- Timestamp: $timestamp
"@

# Write audit note (internal by default)
void

# ---------------------------
# Respond to caller
# ---------------------------
if ($ok) {
    New-JsonResponse -Code 200 -Message "Updated CW ticket UDF #54 (Client Department) and logged audit note." -Extra @{ TicketId=$TicketId; Department=$department; Source=$source }
}
else {
    New-JsonResponse -Code 500 -Message "Failed to update CW ticket UDF #54 (Client Department). Audit note was attempted." -Extra @{ TicketId=$TicketId; Department=$department; Source=$source }
}
