
param($Timer)

<#
  Azure Functions (PowerShell) — Timer Trigger
  Pipeline: Auvik Stats API (device availability) --> CSV --> SharePoint (Drive) via Microsoft Graph

  Notes & references:
   - Auvik API calls must target the regional host: https://auvikapi.{region}.my.auvik.com  [Auvik API Integration Guide]  [1](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
   - Device Availability stats (statId=uptime|outage) return percent/seconds with device/tenant relationships; filters include time range, interval, deviceType. [Auvik Statistics - Device API]  [2](https://elischei.com/how-to-get-site-id-with-graph-explorer-and-other-sharepoint-info/)
   - SharePoint document libraries are Drives; resolve drive from listId via /sites/{site-id}/lists/{list-id}/drive. [Working with SharePoint sites in Graph]  
   - Upload CSV via PUT /drives/{drive-id}/root:/folder/file.csv:/content (≤ ~250 MB). [Graph upload small files]  [3](https://www.fortinet.com/content/dam/fortinet/assets/alliances/sb-fortinet-auvik.pdf)
   - BrightGauge CSV: UTC dates (YYYY-MM-DD), stable headers, no commas in numerics. [BG CSV requirements]  [4](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/v1.0/api/drive-get.md)
#>

# -------------------------------------------------------------
# Enforce TLS 1.2 (Auvik requires TLS 1.2+)
# -------------------------------------------------------------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# -------------------------------------------------------------
# Configuration (env vars) — PS-safe helper
# -------------------------------------------------------------
function Get-EnvVal([string]$name, [string]$default = "") {
  $raw = [System.Environment]::GetEnvironmentVariable($name)
  if ([string]::IsNullOrWhiteSpace($raw)) { return $default }
  return $raw.Trim()
}

$AuvikRegion       = Get-EnvVal "AUVIK_REGION"           # e.g., us5  [1](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
$AuvikUser         = Get-EnvVal "AUVIK_USERNAME"
$AuvikApiKey       = Get-EnvVal "AUVIK_API_KEY"
$TenantsCsv        = Get-EnvVal "AUVIK_TENANTS"
$Interval          = Get-EnvVal "AUVIK_INTERVAL"         # 'day' recommended  [2](https://elischei.com/how-to-get-site-id-with-graph-explorer-and-other-sharepoint-info/)
$DeviceTypesCsv    = Get-EnvVal "AUVIK_DEVICE_TYPES"     # optional: firewall,router
$WindowDays        = if ($env:WINDOW_DAYS) { [int]$env:WINDOW_DAYS } else { 30 }

$TenantId          = Get-EnvVal "Ms365_TenantId"
$ClientId          = Get-EnvVal "Ms365_AuthAppId"
$ClientSecret      = Get-EnvVal "Ms365_AuthSecretId"

$SP_SiteHost       = Get-EnvVal "SP_SiteHost"            # e.g., palaisparc.sharepoint.com
$SP_SitePath       = Get-EnvVal "SP_SitePath"            # e.g., /automation
$SP_ListId         = Get-EnvVal "SP_ListId"              # library listId (GUID)
$SP_FolderPath     = Get-EnvVal "SP_FolderPath"          # e.g., /Reports/Uptime
$OutputFileName    = Get-EnvVal "OUTPUT_FILE_NAME"       # e.g., auvik-uptime-month.csv

# -------------------------------------------------------------
# Diagnostics (no secrets)
# -------------------------------------------------------------
$intervalDisplay = if ([string]::IsNullOrWhiteSpace($Interval)) { 'day' } else { $Interval }
Write-Host "---------- ENV CHECK ----------"
Write-Host ("Auvik Region      : {0}" -f $AuvikRegion)
Write-Host ("Auvik Username    : {0}" -f $AuvikUser)
Write-Host ("Auvik Tenants     : {0}" -f $TenantsCsv)
Write-Host ("Auvik Interval    : {0}" -f $intervalDisplay)
Write-Host ("Window Days       : {0}" -f $WindowDays)
Write-Host ("TenantId          : {0}" -f $TenantId)
Write-Host ("ClientId          : {0}" -f $ClientId)
Write-Host ("ClientSecret Len  : {0}" -f $ClientSecret.Length)
Write-Host ("SP Site Host      : {0}" -f $SP_SiteHost)
Write-Host ("SP Site Path      : {0}" -f $SP_SitePath)
Write-Host ("SP ListId         : {0}" -f $SP_ListId)
Write-Host ("SP Folder Path    : {0}" -f $SP_FolderPath)
Write-Host ("Output File       : {0}" -f $OutputFileName)
Write-Host "-------------------------------"

# -------------------------------------------------------------
# Helpers (Auvik + Graph + CSV)
# -------------------------------------------------------------

# Regional Auvik base (use auvikapi.{region}.my.auvik.com)  [1](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
$BaseAuvik = ("https://auvikapi.{0}.my.auvik.com" -f $AuvikRegion)

# Basic auth header (username:apiKey)  [1](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
$AuthHeader   = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${AuvikUser}:${AuvikApiKey}"))
$HeadersAuvik = @{ Authorization = $AuthHeader; Accept = 'application/json' }

# Time window & interval
$FromUtc   = (Get-Date).ToUniversalTime().AddDays(-$WindowDays)
$ThruUtc   = (Get-Date).ToUniversalTime()
$Interval  = $intervalDisplay

# Lists for tenants/device types
$Tenants     = if ([string]::IsNullOrWhiteSpace($TenantsCsv))    { @($null) } else { $TenantsCsv.Split(',')     | ForEach-Object { $_.Trim() } }
$DeviceTypes = if ([string]::IsNullOrWhiteSpace($DeviceTypesCsv)) { @()       } else { $DeviceTypesCsv.Split(',') | ForEach-Object { $_.Trim() } }

# Convert Auvik "unix minutes" -> UTC day bucket  [2](https://elischei.com/how-to-get-site-id-with-graph-explorer-and-other-sharepoint-info/)
function Convert-FromUnixMinutes([long]$unixMinutes) {
  $secs = $unixMinutes * 60L
  return [DateTimeOffset]::FromUnixTimeSeconds($secs).UtcDateTime.Date
}

# Percent-encode for query params using WebUtility
function UrlEnc([string]$s) {
  return [System.Net.WebUtility]::UrlEncode($s)
}

function Build-Query {
  param([hashtable]$kv)
  $pairs = $kv.GetEnumerator() | ForEach-Object {
    "{0}={1}" -f $_.Key, (UrlEnc([string]$_.Value))
  }
  ($pairs -join '&')
}

# Auvik GET with paging via links.next
function Invoke-AuvikGet {
  param([string]$Url)
  $results = @()
  $includedAll = @()
  $next = $Url
  do {
    $resp = Invoke-RestMethod -Uri $next -Headers $HeadersAuvik -Method GET
    if ($resp.data)     { $results     += $resp.data }
    if ($resp.included) { $includedAll += $resp.included }
    $next = $resp.links.next
  } while ($next)
  return @{ data = $results; included = $includedAll }
}

# Verify Auvik credentials early  [1](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
function Test-AuvikAuth {
  try {
    $verifyUrl = "$BaseAuvik/authentication/verify"
    $resp = Invoke-RestMethod -Uri $verifyUrl -Headers $HeadersAuvik -Method GET
    Write-Host ("Auvik auth verify: {0}" -f ($resp | ConvertTo-Json -Depth 4))
  }
  catch {
    Write-Host ("Auvik auth verify failed: {0}" -f $_.Exception.Message)
    throw
  }
}

# Discover tenants when none provided  [1](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
function Get-AuvikTenants {
  $url = "$BaseAuvik/v1/tenants?page[first]=200"
  try {
    $r = Invoke-AuvikGet -Url $url
    $ids = @()
    foreach ($t in $r.data) { $ids += $t.id }
    Write-Host ("Discovered {0} tenants: {1}" -f $ids.Count, ($ids -join ","))
    return $ids
  } catch {
    Write-Host ("Failed to fetch tenants: {0}" -f $_.Exception.Message)
    return @()
  }
}

# Graph token (client credentials; requires admin-consented app permissions)  
function Get-GraphToken {
  if ([string]::IsNullOrWhiteSpace($TenantId))    { throw "Ms365_TenantId is empty." }
  if ([string]::IsNullOrWhiteSpace($ClientId))    { throw "Ms365_AuthAppId is empty." }
  if ([string]::IsNullOrWhiteSpace($ClientSecret)){ throw "Ms365_AuthSecretId (secret value) is empty." }

  $body = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
  }
  $url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
  Write-Host "Requesting Graph token: $url"

  $tok = Invoke-RestMethod -Uri $url -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
  if (-not $tok.access_token) { throw "Token response missing access_token." }
  return $tok.access_token
}

# Resolve Drive from ListId (document library) via list->drive relationship  
function Resolve-DriveByListId {
  param(
    [string]$GraphToken,
    [string]$SiteHost,   # e.g., palaisparc.sharepoint.com
    [string]$SitePath,   # e.g., /automation
    [string]$ListId
  )
  $gh = @{ Authorization = "Bearer $GraphToken" }

  # 1) Resolve site by path
  $siteUrl = ("https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f $SiteHost, $SitePath)
  $site    = Invoke-RestMethod -Uri $siteUrl -Headers $gh -Method GET

  # 2) Get drive associated with the listId
  $driveRelUrl = ("https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/drive" -f $site.id, $ListId)
  try {
    $driveObj = Invoke-RestMethod -Uri $driveRelUrl -Headers $gh -Method GET
  } catch {
    $driveObj = $null
  }

  if ($driveObj -and $driveObj.id) {
    return @{ driveId = $driveObj.id; siteWebUrl = $site.webUrl; driveName = $driveObj.name }
  }

  # Diagnostics: list lists & drives if resolution fails
  $lists  = Invoke-RestMethod -Uri ("https://graph.microsoft.com/v1.0/sites/{0}/lists?$select=id,name" -f $site.id) -Headers $gh -Method GET
  $drives = Invoke-RestMethod -Uri ("https://graph.microsoft.com/v1.0/sites/{0}/drives?$select=id,name,driveType,webUrl" -f $site.id) -Headers $gh -Method GET

  Write-Host ("No drive matched ListId '{0}' on site '{1}'. Lists:" -f $ListId, $site.webUrl)
  foreach ($l in $lists.value) { Write-Host ("- {0}  id:{1}" -f $l.name,$l.id) }
  Write-Host "Drives:"
  foreach ($d in $drives.value) { Write-Host ("- {0}  id:{1}  type:{2}" -f $d.name,$d.id,$d.driveType) }

  throw "Provided SP_ListId did not resolve to a drive. Verify the ListId belongs to a Document Library on this site."
}

# Optional: ensure nested folder path exists (create missing segments)  
function Ensure-FolderPath {
  param(
    [string]$GraphToken,
    [string]$DriveId,
    [string]$FolderPath  # e.g., /Reports/Uptime
  )
  $ghJson = @{ Authorization = "Bearer $GraphToken"; "Content-Type" = "application/json" }
  $gh     = @{ Authorization = "Bearer $GraphToken" }

  $segments = ($FolderPath.Trim()) -split '/'
  $segments = $segments | Where-Object { $_ -ne '' }

  $currentPath = ""
  $parentId = "root"

  foreach ($seg in $segments) {
    $currentPath = if ($currentPath) { "$currentPath/$seg" } else { $seg }
    $getUrl = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}" -f $DriveId, $currentPath)

    try {
      $item = Invoke-RestMethod -Uri $getUrl -Headers $gh -Method GET
      $parentId = $item.id
    } catch {
      # Create folder under parentId
      $createUrl = ("https://graph.microsoft.com/v1.0/drives/{0}/items/{1}/children" -f $DriveId, $parentId)
      $body = @{ name = $seg; folder = @{}; "@microsoft.graph.conflictBehavior" = "fail" } | ConvertTo-Json
      $newItem = Invoke-RestMethod -Uri $createUrl -Headers $ghJson -Method POST -Body $body
      $parentId = $newItem.id
      Write-Host ("Created folder segment: {0}" -f $seg)
    }
  }
  return $currentPath
}

# Upload CSV to the resolved drive via PUT /content (supports up to ~250 MB)  [3](https://www.fortinet.com/content/dam/fortinet/assets/alliances/sb-fortinet-auvik.pdf)
function Upload-CsvToDrive {
  param(
    [string]$GraphToken,
    [string]$DriveId,
    [string]$FolderPath,
    [string]$FileName,
    [string]$CsvText
  )
  $headers    = @{ Authorization = "Bearer $GraphToken"; "Content-Type" = "text/csv" }

  # Ensure folder path exists (optional but robust)
  $createdPath = Ensure-FolderPath -GraphToken $GraphToken -DriveId $DriveId -FolderPath $FolderPath

  $uploadUrl  = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}:/content" -f $DriveId, $createdPath, $FileName)
  Write-Host "Uploading CSV to: $uploadUrl"
  Invoke-RestMethod -Uri $uploadUrl -Method PUT -Headers $headers -Body $CsvText | Out-Null
}

# -------------------------------------------------------------
# Collect Auvik availability stats (uptime% + outage seconds)
# -------------------------------------------------------------

# Resolve tenant list if not supplied
if ($Tenants.Count -eq 1 -and $Tenants[0] -eq $null) {
  Test-AuvikAuth
  $Tenants = Get-AuvikTenants
  if ($Tenants.Count -eq 0) {
    Write-Host "No Auvik tenants available to this API user; Stats may be empty."
  }
} else {
  # Preflight auth check even if Tenants provided
  Test-AuvikAuth
}

$rows = New-Object System.Collections.Generic.List[object]

foreach ($tenant in $Tenants) {

  # If device types provided, iterate; else single call without filter
  $deviceTypeLoop = if ($DeviceTypes.Count -gt 0) { $DeviceTypes } else { @($null) }

  foreach ($devType in $deviceTypeLoop) {

    $q = @{
      "filter[fromTime]" = $FromUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
      "filter[thruTime]" = $ThruUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
      "filter[interval]" = $Interval
      "page[first]"      = 500
    }
    if ($tenant)  { $q["tenants"] = $tenant }
    if ($devType) { $q["filter[deviceType]"] = $devType }  # optional device type filter  [2](https://elischei.com/how-to-get-site-id-with-graph-explorer-and-other-sharepoint-info/)

    # Uptime %
    $uptUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($q + @{ "statId" = "uptime" }))
    # Outage (seconds)
    $outUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($q + @{ "statId" = "outage" }))

    Write-Host ("Uptime URL: {0}" -f $uptUrl)
    Write-Host ("Outage URL: {0}" -f $outUrl)

    $upt = Invoke-AuvikGet -Url $uptUrl
    $out = Invoke-AuvikGet -Url $outUrl

    Write-Host ("Tenant {0} | devType {1} | uptime items: {2} | outage items: {3}" -f ($tenant ?? "<none>"), ($devType ?? "<none>"), ($upt.data.Count), ($out.data.Count))

    # If data is zero with a device type filter, retry without the filter (diagnostics)
    if ($devType -and $upt.data.Count -eq 0 -and $out.data.Count -eq 0) {
      $q.Remove("filter[deviceType]")
      $uptUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($q + @{ "statId" = "uptime" }))
      $outUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($q + @{ "statId" = "outage" }))
      Write-Host ("Retry without deviceType: Uptime URL: {0}" -f $uptUrl)
      Write-Host ("Retry without deviceType: Outage URL: {0}" -f $outUrl)
      $upt = Invoke-AuvikGet -Url $uptUrl
      $out = Invoke-AuvikGet -Url $outUrl
      Write-Host ("Tenant {0} | devType <none> | uptime items: {1} | outage items: {2}" -f ($tenant ?? "<none>"), ($upt.data.Count), ($out.data.Count))
    }

    # Build outage lookup by device+date
    $outLookup = @{}
    foreach ($o in $out.data) {
      $devId = $o.relationships.device.data.id
      $tenId = $o.relationships.tenant.data.id
      $ts    = Convert-FromUnixMinutes([long]$o.attributes.time)
      $key   = "$tenId|$devId|$($ts.ToString('yyyy-MM-dd'))"

      $total = $o.attributes.total
      if ($null -eq $total) { $total = 0 }
      $outLookup[$key] = $total
    }

    # Map uptime rows
    foreach ($u in $upt.data) {
      $devId = $u.relationships.device.data.id
      $tenId = $u.relationships.tenant.data.id
      $ts    = Convert-FromUnixMinutes([long]$u.attributes.time)
      $key   = "$tenId|$devId|$($ts.ToString('yyyy-MM-dd'))"

      $avg = $u.attributes.average
      if ($null -eq $avg) { $avg = 0 }
      $uptPct = [math]::Round($avg, 3)

      $outSec = if ($outLookup.ContainsKey($key)) { $outLookup[$key] } else { 0 }
      $outMin = [math]::Round(($outSec / 60.0), 2)

      $rows.Add([pscustomobject]@{
        date           = $ts.ToString('yyyy-MM-dd')   # BrightGauge date (UTC)  [4](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/v1.0/api/drive-get.md)
        tenant_id      = $tenId
        tenant_name    = ""                           # optional
        site_id        = ""                           # optional
        site_name      = ""                           # optional
        device_id      = $devId
        device_name    = ""                           # optional
        device_type    = if ($devType) { $devType } else { "" }
        uptime_percent = $uptPct
        outage_minutes = $outMin
        interval       = $Interval
      })
    }
  }
}

Write-Host ("Auvik rows collected: {0}" -f $rows.Count)

# -------------------------------------------------------------
# CSV: BrightGauge-friendly (stable headers, no commas)
# -------------------------------------------------------------
$headers  = @('date','tenant_id','tenant_name','site_id','site_name','device_id','device_name','device_type','uptime_percent','outage_minutes','interval')
$csvLines = New-Object System.Collections.Generic.List[string]
$csvLines.Add(($headers -join ','))

foreach ($r in $rows) {
  $vals = foreach ($h in $headers) {
    $val = if ($null -eq $r.$h) { "" } else { [string]$r.$h }
    $val -replace ',', ' '  # BG: no commas in numbers/fields  [4](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/v1.0/api/drive-get.md)
  }
  $csvLines.Add(($vals -join ','))
}
$csv = $csvLines -join "`n"

# -------------------------------------------------------------
# Upload to SharePoint — resolve drive by listId and PUT /content
# -------------------------------------------------------------
$graphToken = Get-GraphToken
$driveInfo  = Resolve-DriveByListId -GraphToken $graphToken -SiteHost $SP_SiteHost -SitePath $SP_SitePath -ListId $SP_ListId
Upload-CsvToDrive -GraphToken $graphToken -DriveId $driveInfo.driveId -FolderPath $SP_FolderPath -FileName $OutputFileName -CsvText $csv

Write-Host ("Uploaded {0} rows to drive '{1}' at {2}/{3}" -f $rows.Count, $driveInfo.driveName, $SP_FolderPath, $OutputFileName)
