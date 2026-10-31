#Requires -Version 7.0
<#
.SYNOPSIS
  READ-ONLY discovery for the "project status -> callback -> Azure Function" design.
  Makes only GET requests (and an optional OPTIONS/GET probe of your function URL). Changes nothing.

.DESCRIPTION
  Reports what the provisioning script will need:
    1. Auth + API member sanity check, CW version
    2. Project boards (service/boards with projectFlag) and their ids
    3. Existing project statuses (project/statuses) and which target statuses already exist
    4. Existing callbacks (system/callbacks) - any already pointing at your function, duplicates
    5. Callback metadata probe (allowed types/levels) where the tenant exposes it
    6. A sample project: its status field shape and board
    7. Optional reachability test of the Azure Function URL
  Output: console summary + JSON/CSV under .\Discovery-Output

  Credentials: parameters, or env vars CW_COMPANY_ID, CW_PUBLIC_KEY, CW_PRIVATE_KEY, CW_CLIENT_ID,
  or you are prompted. Nothing is stored in this file.

.EXAMPLE
  $env:CW_COMPANY_ID='cloudfirst'; $env:CW_PUBLIC_KEY='...'; $env:CW_PRIVATE_KEY='...'; $env:CW_CLIENT_ID='...'
  pwsh -File .\Discover-CwProjectCallback.ps1 -AzureFunctionUrl 'https://myapp.azurewebsites.net/api/ConnectWiseProjectCallback'
#>
[CmdletBinding()]
param(
    [string]$CwApiBase   = 'https://api-na.myconnectwise.net/v4_6_release/apis/3.0',
    [string]$CompanyId   = $env:CW_COMPANY_ID,
    [string]$PublicKey   = $env:CW_PUBLIC_KEY,
    [string]$PrivateKey  = $env:CW_PRIVATE_KEY,
    [string]$ClientId    = $env:CW_CLIENT_ID,
    [string]$ApiVersion  = '2025.1',
    [string]$AzureFunctionUrl,
    [string[]]$TargetStatuses = @('SharePoint Update', 'SP Update In-Progress', 'SP Update Failed'),
    [string]$OutputDir   = (Join-Path $PSScriptRoot 'Discovery-Output')
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

if (-not $CompanyId)  { $CompanyId  = Read-Host 'CW company id' }
if (-not $PublicKey)  { $PublicKey  = Read-Host 'CW public key' }
if (-not $PrivateKey) { $PrivateKey = Read-Host 'CW private key' -MaskInput }
if (-not $ClientId)   { $ClientId   = Read-Host 'CW client id' }

$auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$CompanyId+$PublicKey`:$PrivateKey"))
$headers = @{
    Authorization  = "Basic $auth"
    ClientID       = $ClientId
    'Content-Type' = 'application/json'
    Accept         = "application/vnd.connectwise.com+json; version=$ApiVersion"
}

$report = [ordered]@{ RunUtc = (Get-Date).ToUniversalTime().ToString('o'); Notes = @() }
function Note($m) { $report.Notes += $m; Write-Host "  ! $m" -ForegroundColor Yellow }

# GET only. Returns @{ Ok; Data; Error }
function Get-Cw([string]$Path, [hashtable]$Query = @{}) {
    $qs = ($Query.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString([string]$_.Value))" }) -join '&'
    $uri = "$CwApiBase/$Path" + $(if ($qs) { "?$qs" } else { '' })
    try {
        $data = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -TimeoutSec 60
        return [pscustomobject]@{ Ok = $true; Data = $data; Error = $null; Uri = $uri }
    } catch {
        $code = $null; $body = $null
        if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
        if ($_.ErrorDetails) { $body = $_.ErrorDetails.Message }
        return [pscustomobject]@{ Ok = $false; Data = $null; Error = "HTTP $code $body".Trim(); Uri = $uri }
    }
}

# Pages through a list endpoint
function Get-CwAll([string]$Path, [hashtable]$Query = @{}) {
    $all = [System.Collections.Generic.List[object]]::new()
    $page = 1
    while ($true) {
        $q = @{} + $Query; $q.page = $page; $q.pageSize = 1000
        $r = Get-Cw $Path $q
        if (-not $r.Ok) { return [pscustomobject]@{ Ok = $false; Data = @(); Error = $r.Error } }
        $items = @($r.Data)
        foreach ($i in $items) { $all.Add($i) }
        if ($items.Count -lt 1000) { break }
        $page++
    }
    [pscustomobject]@{ Ok = $true; Data = $all.ToArray(); Error = $null }
}

Write-Host "`n=== 1. Auth / system info ===" -ForegroundColor Cyan
$info = Get-Cw 'system/info'
if (-not $info.Ok) { throw "Auth or connectivity failed: $($info.Error)" }
$report.System = $info.Data
Write-Host "  Connected. Version: $($info.Data.version)  Cloud: $($info.Data.isCloud)  Company: $CompanyId"

Write-Host "`n=== 2. Project boards ===" -ForegroundColor Cyan
$boards = Get-CwAll 'project/boardSetups'   # not present in every version; fall back below
if (-not $boards.Ok) {
    $boards = Get-CwAll 'service/boards' @{ conditions = 'projectFlag=true' }
}
if (-not $boards.Ok) { Note "Could not read project boards: $($boards.Error)"; $boards = [pscustomobject]@{ Data = @() } }
$report.ProjectBoards = @($boards.Data | ForEach-Object { [pscustomobject]@{ id = $_.id; name = $_.name; inactive = $_.inactiveFlag; projectFlag = $_.projectFlag } })
$report.ProjectBoards | Format-Table -AutoSize | Out-String | Write-Host

Write-Host "=== 3. Project statuses ===" -ForegroundColor Cyan
$statuses = Get-CwAll 'project/statuses'
if (-not $statuses.Ok) {
    Note "project/statuses failed ($($statuses.Error)); trying project/projectStatuses"
    $statuses = Get-CwAll 'project/projectStatuses'
}
if (-not $statuses.Ok) { Note "Could not read project statuses: $($statuses.Error)"; $statuses = [pscustomobject]@{ Data = @() } }
$report.ProjectStatuses = @($statuses.Data | ForEach-Object {
    [pscustomobject]@{
        id = $_.id; name = $_.name; closedFlag = $_.closedFlag; inactive = $_.inactiveFlag
        defaultFlag = $_.defaultFlag; boardId = $(if ($_.PSObject.Properties['board']) { $_.board.id } else { $null })
        fields = ($_.PSObject.Properties.Name | Where-Object { $_ -ne '_info' }) -join ','
    } })
$report.ProjectStatuses | Select-Object id, name, closedFlag, inactive, defaultFlag, boardId | Format-Table -AutoSize | Out-String | Write-Host
$report.StatusFieldShape = if ($report.ProjectStatuses.Count) { $report.ProjectStatuses[0].fields } else { $null }

Write-Host "  Target statuses:" -ForegroundColor Cyan
$report.TargetStatusCheck = foreach ($t in $TargetStatuses) {
    $hit = @($report.ProjectStatuses | Where-Object { $_.name -eq $t })
    $row = [pscustomobject]@{ Name = $t; Exists = ($hit.Count -gt 0); Ids = (@($hit | ForEach-Object { $_.id }) -join ',') }
    Write-Host ("    {0,-26} {1}" -f $t, $(if ($row.Exists) { "EXISTS (id $($row.Ids))" } else { 'missing' }))
    $row
}

Write-Host "`n=== 4. Existing callbacks ===" -ForegroundColor Cyan
$callbacks = Get-CwAll 'system/callbacks'
if (-not $callbacks.Ok) {
    Note "system/callbacks failed: $($callbacks.Error). The API member may lack 'System > Manage Hooks/Callbacks' permission."
    $callbacks = [pscustomobject]@{ Data = @() }
}
$report.Callbacks = @($callbacks.Data | ForEach-Object {
    [pscustomobject]@{
        id = $_.id; description = $_.description
        urlHost = $(if ($_.url) { try { ([uri]$_.url).Host } catch { 'invalid' } } else { $null })
        urlBase = $(if ($_.url) { ($_.url -split '\?')[0] } else { $null })   # query strings dropped: they hold vendor keys
        type = $_.type; level = $_.level
        objectId = $_.objectId; memberId = $(if ($_.PSObject.Properties['memberId']) { $_.memberId } else { $null })
        inactive = $(if ($_.PSObject.Properties['inactiveFlag']) { $_.inactiveFlag } else { $null })
    } })
Write-Host "  $($report.Callbacks.Count) callback(s) registered. Summary by type/level:"
$report.CallbackTypesInUse = @($report.Callbacks | Group-Object type, level | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) x$($_.Count)" })
$report.CallbackTypesInUse | ForEach-Object { Write-Host "    $_" }
Write-Host "  By destination host:"
$report.Callbacks | Group-Object urlHost | Sort-Object Count -Descending | Select-Object -First 15 | ForEach-Object { Write-Host ("    {0,-50} {1}" -f $_.Name, $_.Count) }
Write-Host "  Project-type callbacks (the ones relevant to us):"
$projCb = @($report.Callbacks | Where-Object { $_.type -match 'project' })
if ($projCb.Count) { $projCb | Select-Object id, description, urlHost, level, objectId, inactive | Format-Table -AutoSize | Out-String -Width 200 | Write-Host } else { Write-Host '    none' }
$report.ProjectCallbacks = $projCb
if ($AzureFunctionUrl) {
    $base = ($AzureFunctionUrl -split '\?')[0]
    $dupes = @($report.Callbacks | Where-Object { $_.urlBase -eq $base })
    if ($dupes) { Note "A callback already points at this function (ids: $((@($dupes | ForEach-Object { $_.id })) -join ',')). Provisioning must reuse/update it, not add another." }
}


Write-Host "=== 5. Callback metadata probe ===" -ForegroundColor Cyan
$report.CallbackProbe = foreach ($p in 'system/callbacks/info', 'system/callbacks/count') {
    $r = Get-Cw $p
    Write-Host ("  {0,-26} {1}" -f $p, $(if ($r.Ok) { 'OK' } else { $r.Error }))
    [pscustomobject]@{ Path = $p; Ok = $r.Ok; Data = $r.Data; Error = $r.Error }
}
Write-Host "  (Valid type/level combos are not listed by the API; the provisioning run is where they get tested.)"

Write-Host "`n=== 6. Sample projects (status/board field shape) ===" -ForegroundColor Cyan
$proj = Get-Cw 'project/projects' @{ pageSize = 5; orderBy = 'id desc' }
if ($proj.Ok) {
    $report.SampleProjects = @($proj.Data | ForEach-Object {
        [pscustomobject]@{ id = $_.id; name = $_.name; status = $_.status.name; statusId = $_.status.id; board = $_.board.name; boardId = $_.board.id; company = $_.company.name }
    })
    $report.SampleProjects | Format-Table -AutoSize | Out-String | Write-Host
    if (@($proj.Data).Count) { $report.ProjectFieldNames = ($proj.Data[0].PSObject.Properties.Name) -join ',' }
} else { Note "Could not read projects: $($proj.Error)" }

Write-Host "=== 7. Azure Function reachability ===" -ForegroundColor Cyan
if ($AzureFunctionUrl) {
    $base = ($AzureFunctionUrl -split '\?')[0]
    try {
        $resp = Invoke-WebRequest -Uri $base -Method Get -SkipHttpErrorCheck -TimeoutSec 20
        $report.FunctionProbe = "HTTP $([int]$resp.StatusCode)"
        Write-Host "  GET $base -> HTTP $([int]$resp.StatusCode)  (401/403/405 means reachable and protected/POST-only; 404 means wrong route or not deployed)"
    } catch {
        $report.FunctionProbe = "Unreachable: $($_.Exception.Message)"
        Note $report.FunctionProbe
    }
} else {
    Write-Host '  Skipped (pass -AzureFunctionUrl to test).'
}

# ---- Findings ----
Write-Host "`n=== Findings ===" -ForegroundColor Cyan
$missing = @($report.TargetStatusCheck | Where-Object { -not $_.Exists } | ForEach-Object { $_.Name })
Write-Host "  Project boards found:        $($report.ProjectBoards.Count)"
Write-Host "  Statuses to create:          $(if ($missing) { $missing -join ', ' } else { 'none' })"
Write-Host "  Statuses are board-scoped:   $(if (($report.ProjectStatuses | Where-Object boardId)) { 'YES - create per board' } else { 'no board field seen (global)' })"
Write-Host "  Existing callbacks:          $($report.Callbacks.Count)"

New-Item -ItemType Directory -Force $OutputDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$json = Join-Path $OutputDir "Discovery-$stamp.json"
$report | ConvertTo-Json -Depth 8 | Set-Content -Path $json -Encoding UTF8
$report.ProjectStatuses | Export-Csv (Join-Path $OutputDir "ProjectStatuses-$stamp.csv") -NoTypeInformation
$report.Callbacks       | Export-Csv (Join-Path $OutputDir "Callbacks-$stamp.csv") -NoTypeInformation
Write-Host "`nNo changes were made to ConnectWise. Report: $json"
