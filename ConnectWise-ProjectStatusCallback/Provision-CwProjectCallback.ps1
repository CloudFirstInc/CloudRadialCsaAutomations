#Requires -Version 7.1
<#
.SYNOPSIS
  Provisions the "project status -> callback -> Azure Function" design in ConnectWise.
  DRY RUN by default. Pass -Apply to make changes.

.DESCRIPTION
  1. Creates missing global project statuses (project/statuses): SharePoint Update, SP Update In-Progress, SP Update Failed
  2. Creates (or reuses/updates) a project callback to your Azure Function, scoped to the trigger status
     (type=project, level=status, objectId=<trigger status id>).
  Idempotent: existing statuses/callbacks are reused. Safe to re-run.

  Findings from discovery (cloudfirst, CW v2025.1): statuses are global (no board), no project callbacks exist yet.
  Callback type/level combos are not discoverable via the API, so the script reports the API error verbatim
  if the tenant rejects a combination; override with -CallbackType / -CallbackLevel.

  Credentials: parameters, env vars CW_COMPANY_ID, CW_PUBLIC_KEY, CW_PRIVATE_KEY, CW_CLIENT_ID, or prompted.

.EXAMPLE
  pwsh -File .\Provision-CwProjectCallback.ps1 -AzureFunctionUrl 'https://myapp.azurewebsites.net/api/ConnectWiseProjectCallback?code=KEY'
  pwsh -File .\Provision-CwProjectCallback.ps1 -AzureFunctionUrl '...' -Apply
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AzureFunctionUrl,
    [string]$CwApiBase   = 'https://api-na.myconnectwise.net/v4_6_release/apis/3.0',
    [string]$CompanyId   = $env:CW_COMPANY_ID,
    [string]$PublicKey   = $env:CW_PUBLIC_KEY,
    [string]$PrivateKey  = $env:CW_PRIVATE_KEY,
    [string]$ClientId    = $env:CW_CLIENT_ID,
    [string]$ApiVersion  = '2025.1',
    [string]$TriggerStatus = 'SharePoint Update',
    [string[]]$ExtraStatuses = @('SP Update In-Progress', 'SP Update Failed'),
    [string]$CallbackType  = 'project',
    [string]$CallbackLevel = 'status',
    [string]$CallbackDescription = 'Project status -> SharePoint update (Azure Function)',
    [switch]$Apply
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

if ($AzureFunctionUrl -notmatch '^https://') { throw 'AzureFunctionUrl must be an https:// URL (ConnectWise requires HTTPS callbacks).' }

if (-not $CompanyId)  { $CompanyId  = Read-Host 'CW company id' }
if (-not $PublicKey)  { $PublicKey  = Read-Host 'CW public key' }
if (-not $PrivateKey) { $PrivateKey = Read-Host 'CW private key' -MaskInput }
if (-not $ClientId)   { $ClientId   = Read-Host 'CW client id' }

$auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${CompanyId}+${PublicKey}:${PrivateKey}"))
$headers = @{
    Authorization  = "Basic $auth"
    ClientID       = $ClientId
    'Content-Type' = 'application/json'
    Accept         = "application/vnd.connectwise.com+json; version=$ApiVersion"
}

function Invoke-Cw([string]$Method, [string]$Path, $Body = $null, [hashtable]$Query = @{}) {
    $qs = ($Query.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join '&'
    $uri = "$CwApiBase/$Path" + $(if ($qs) { "?$qs" } else { '' })
    $req = @{ Uri = $uri; Headers = $headers; Method = $Method; TimeoutSec = 60 }
    if ($null -ne $Body) { $req.Body = ($Body | ConvertTo-Json -Depth 6 -Compress) }
    try { Invoke-RestMethod @req }
    catch {
        $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { '' }
        throw "CW $Method $Path failed: HTTP $code $($_.ErrorDetails.Message)"
    }
}

function Get-CwAll([string]$Path, [hashtable]$Query = @{}) {
    $all = [System.Collections.Generic.List[object]]::new()
    $page = 1
    while ($true) {
        $q = @{} + $Query; $q.page = $page; $q.pageSize = 1000
        $items = @(Invoke-Cw Get $Path -Query $q)
        foreach ($i in $items) { $all.Add($i) }
        if ($items.Count -lt 1000) { break }
        $page++
    }
    $all.ToArray()
}

$mode = if ($Apply) { 'APPLY' } else { 'DRY RUN (no changes; pass -Apply)' }
Write-Host "`nMode: $mode" -ForegroundColor Cyan

Write-Host "`n=== Auth check ===" -ForegroundColor Cyan
$info = Invoke-Cw Get 'system/info'
Write-Host "  Connected. CW $($info.version)"

# ---- Statuses ----
Write-Host "`n=== Project statuses ===" -ForegroundColor Cyan
$statuses = @(Get-CwAll 'project/statuses')
$wanted = @($TriggerStatus) + $ExtraStatuses | Select-Object -Unique
$statusIds = @{}
foreach ($name in $wanted) {
    $hit = @($statuses | Where-Object { $_.name -eq $name })
    if ($hit.Count) {
        $statusIds[$name] = $hit[0].id
        Write-Host "  exists   $name (id $($hit[0].id))"
        continue
    }
    if (-not $Apply) { Write-Host "  would create  $name"; continue }
    $new = Invoke-Cw Post 'project/statuses' @{ name = $name; defaultFlag = $false; closedFlag = $false; inactiveFlag = $false }
    $statusIds[$name] = $new.id
    Write-Host "  created  $name (id $($new.id))" -ForegroundColor Green
}

# ---- Callback ----
Write-Host "`n=== Callback ===" -ForegroundColor Cyan
$triggerId = $statusIds[$TriggerStatus]
if (-not $triggerId) {
    Write-Host "  Trigger status '$TriggerStatus' does not exist yet (dry run). Callback would be created after the status."
    Write-Host "`nDry run complete. Re-run with -Apply."
    return
}

$base = ($AzureFunctionUrl -split '\?')[0]
# Filtered server-side: the tenant has ~17k callbacks.
$existing = @(Get-CwAll 'system/callbacks' @{ conditions = "type='$CallbackType'" } | Where-Object { ($_.url -split '\?')[0] -eq $base })

$body = [ordered]@{
    description  = $CallbackDescription
    url          = $AzureFunctionUrl
    objectId     = $triggerId
    type         = $CallbackType
    level        = $CallbackLevel
    inactiveFlag = $false
}

$match = @($existing | Where-Object { $_.level -eq $CallbackLevel -and $_.objectId -eq $triggerId })
if ($match.Count) {
    Write-Host "  Callback already registered (id $($match[0].id)) for status $triggerId. Nothing to do."
} elseif ($existing.Count) {
    Write-Host "  Callback(s) to this function exist with different level/objectId: ids $((($existing).id) -join ',')." -ForegroundColor Yellow
    if ($Apply) {
        $upd = Invoke-Cw Put "system/callbacks/$($existing[0].id)" ($body + @{ id = $existing[0].id })
        Write-Host "  Updated callback id $($upd.id)." -ForegroundColor Green
    } else { Write-Host "  would update callback id $($existing[0].id)" }
} else {
    if ($Apply) {
        $cb = Invoke-Cw Post 'system/callbacks' $body
        Write-Host "  Created callback id $($cb.id): $CallbackType/$CallbackLevel objectId=$triggerId" -ForegroundColor Green
    } else { Write-Host "  would create callback: $CallbackType/$CallbackLevel objectId=$triggerId -> $base" }
}

Write-Host "`n$(if ($Apply) { 'Provisioning complete.' } else { 'Dry run complete. Re-run with -Apply.' })"
Write-Host "Reminder: the function must set projects to '$($wanted -join ''', ''')' itself; it must also ignore its own status updates to avoid loops."
