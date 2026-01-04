
param($Timer)

<#
  Azure Functions (PowerShell 5.1) — Timer Trigger
  Multi-tenant Auvik → CSVs (firewall device uptime & WAN internet uptime) → SharePoint via Graph

  References:
   - Auvik Requested Time Ranges (midnight alignment required)     https://support.auvik.com/hc/en-us/articles/360059283312-Statistics-API-Requested-Time-Ranges
   - Device Availability stats (uptime/outage)                      https://support.auvik.com/hc/en-us/articles/360044579852-Statistics-Device-API
   - Service stats (cloud ping: pingTime, pingPacket)               https://support.auvik.com/hc/en-us/articles/360045023551-Statistics-Service-API
   - Auvik API integration guidance (headers, region, TLS)          https://support.auvik.com/hc/en-us/articles/360031007111-Auvik-API-Integration-Guide
   - Microsoft Graph upload (PUT /content)                          https://learn.microsoft.com/graph/api/driveitem-put-content
#>

# -------------------------------------------------------------
# TLS 1.2 (required by Auvik)
# -------------------------------------------------------------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# -------------------------------------------------------------
# Env helper (PowerShell-safe)
# -------------------------------------------------------------
function Get-EnvVal([string]$name, [string]$default = "") {
  $raw = [System.Environment]::GetEnvironmentVariable($name)
  if ([string]::IsNullOrWhiteSpace($raw)) { return $default }
  return $raw.Trim()
}

# -------------------------------------------------------------
# Configuration (Auvik)
# -------------------------------------------------------------
$AuvikRegion    = Get-EnvVal "AUVIK_REGION" "us5"
$AuvikUser      = Get-EnvVal "AUVIK_USERNAME" "dgentles@cloudfirstinc.com"
$AuvikApiKey    = Get-EnvVal "AUVIK_API_KEY"  "<PASTE_API_KEY_IN_APP_SETTINGS>"
# Optional Accept header override (e.g., application/vnd.api+json)
$AuvikAccept    = Get-EnvVal "AUVIK_ACCEPT" "application/json"

# Tenants: blank = auto-discover all
$TenantsCsv     = Get-EnvVal "AUVIK_TENANTS" ""
# Device types to target; start with firewalls (add 'router' later if desired)
$DeviceTypesCsv = Get-EnvVal "AUVIK_DEVICE_TYPES" "firewall"
$DeviceTypes    = if ([string]::IsNullOrWhiteSpace($DeviceTypesCsv)) { @() } else { $DeviceTypesCsv.Split(',') | ForEach-Object { $_.Trim() } }

# Time window & interval (Patch A: align to midnight UTC boundaries)
$WindowDays     = if ($env:WINDOW_DAYS) { [int]$env:WINDOW_DAYS } else { 30 }
$Interval       = Get-EnvVal "AUVIK_INTERVAL" "day"

# Midnight UTC alignment required by Auvik stats API
$todayMidnightUtc     = (Get-Date).ToUniversalTime().Date
$FromUtcAligned       = $todayMidnightUtc.AddDays(-$WindowDays)  # midnight N days ago
$ThruUtcAligned       = $todayMidnightUtc.AddDays(1)             # next midnight to include full current day
$FromStr = $FromUtcAligned.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
$ThruStr = $ThruUtcAligned.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")

# Regional host + Basic (wrap vars with ${} to avoid ':' parser issues)
$BaseAuvik      = "https://auvikapi.$AuvikRegion.my.auvik.com"
$BasicToken     = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${AuvikUser}:${AuvikApiKey}"))
$HeadersAuvik   = @{ Authorization = "Basic $BasicToken"; Accept = $AuvikAccept }

# -------------------------------------------------------------
# Configuration (Graph / SharePoint)
# -------------------------------------------------------------
$TenantId       = Get-EnvVal "Ms365_TenantId"
$ClientId       = Get-EnvVal "Ms365_AuthAppId"
$ClientSecret   = Get-EnvVal "Ms365_AuthSecretId"

$SP_SiteHost    = Get-EnvVal "SP_SiteHost"
$SP_SitePath    = Get-EnvVal "SP_SitePath"
$SP_ListId      = Get-EnvVal "SP_ListId"
$SP_FolderPath  = Get-EnvVal "SP_FolderPath" "/Reports/Uptime"

# Output CSVs (BG datasets)
$OutputFileDev  = Get-EnvVal "OUTPUT_FILE_NAME" "auvik-uptime-month.csv"
$OutputFileWan  = "wan-internet-uptime-month.csv"

# -------------------------------------------------------------
# Diagnostics (no secrets)
# -------------------------------------------------------------
Write-Host "---------- ENV CHECK ----------"
Write-Host ("Auvik Region      : {0}" -f $AuvikRegion)
Write-Host ("Auvik Username    : {0}" -f $AuvikUser)
Write-Host ("Auvik Tenants CSV : {0}" -f $TenantsCsv)
Write-Host ("Device Types      : {0}" -f ($DeviceTypes -join ','))
Write-Host ("Interval          : {0}" -f $Interval)
Write-Host ("Window Days       : {0}" -f $WindowDays)
Write-Host ("TenantId          : {0}" -f $TenantId)
Write-Host ("ClientId          : {0}" -f $ClientId)
Write-Host ("ClientSecret Len  : {0}" -f $ClientSecret.Length)
Write-Host ("SP Site Host      : {0}" -f $SP_SiteHost)
Write-Host ("SP Site Path      : {0}" -f $SP_SitePath)
Write-Host ("SP ListId         : {0}" -f $SP_ListId)
Write-Host ("SP Folder Path    : {0}" -f $SP_FolderPath)
Write-Host ("CSV (devices)     : {0}" -f $OutputFileDev)
Write-Host ("CSV (wan)         : {0}" -f $OutputFileWan)
Write-Host "-------------------------------"

# -------------------------------------------------------------
# Utility helpers
# -------------------------------------------------------------
function UrlEnc([string]$s) { [System.Net.WebUtility]::UrlEncode($s) }
function Build-Query {
  param([hashtable]$kv)
  $pairs = $kv.GetEnumerator() | ForEach-Object {
    "{0}={1}" -f $_.Key, (UrlEnc([string]$_.Value))
  }
  ($pairs -join '&')
}
function Convert-FromUnixMinutes([long]$unixMinutes) {
  $secs = $unixMinutes * 60L
  return [DateTimeOffset]::FromUnixTimeSeconds($secs).UtcDateTime.Date
}

# Inspect HTTP status for verify (Auvik often returns 200 with empty body)
function Test-AuvikAuth {
  try {
    $verifyUrl = "$BaseAuvik/authentication/verify"
    $resp = Invoke-WebRequest -Uri $verifyUrl -Headers $HeadersAuvik -Method GET
    Write-Host ("Auvik auth verify status: {0}" -f $resp.StatusCode)
    if ($resp.StatusCode -ne 200) { throw ("Auth verify returned {0}" -f $resp.StatusCode) }
  } catch {
    Write-Host ("Auvik auth verify failed: {0}" -f $_.Exception.Message)
    throw
  }
}

# Auvik GET with paging — skip 403; log known 400/404; bubble up others
function Invoke-AuvikGet {
  param([string]$Url)
  $results = @()
  $next = $Url
  do {
    try {
      $resp = Invoke-RestMethod -Uri $next -Headers $HeadersAuvik -Method GET
      if ($resp.data) { $results += $resp.data }
      $next = $resp.links.next
    } catch {
      $httpError = $_.Exception.Response
      if ($httpError) {
        $code = $httpError.StatusCode.value__
        if ($code -eq 403) {
          Write-Host ("[SKIP] 403 Forbidden on {0} — tenant not authorized" -f $next)
          break
        }
        elseif ($code -eq 400) {
          try {
            $reader = New-Object IO.StreamReader($httpError.GetResponseStream())
            $body = $reader.ReadToEnd()
          } catch { $body = "" }
          Write-Host ("[WARN] 400 on {0} → {1}" -f $next, $body)
          break
        }
        elseif ($code -eq 404) {
          Write-Host ("[WARN] 404 Not Found on {0}" -f $next)
          break
        }
      }
      throw
    }
  } while ($next)
  return $results
}

# Discover tenants; returns an array of @{ id; name } and a name map
function Get-AuvikTenants {
  $url = "$BaseAuvik/v1/tenants?page[first]=500"
  try {
    $data = Invoke-AuvikGet -Url $url
    $tenantsOut = @()
    foreach ($t in $data) {
      $tenantsOut += [pscustomobject]@{
        id   = $t.id
        name = $t.attributes.name
      }
    }
    Write-Host ("Discovered {0} tenants" -f $tenantsOut.Count)
    return $tenantsOut
  } catch {
    Write-Host ("Failed to fetch tenants: {0}" -f $_.Exception.Message)
    return @()
  }
}

# 🔒 Pre-filter: remove tenant IDs that 403
function Filter-AuthorizedTenants {
  param([string[]]$TenantIds)
  $authorized = New-Object System.Collections.Generic.List[object]
  foreach ($tid in $TenantIds) {
    $probeUrl = "$BaseAuvik/v1/inventory/device/info?tenants=$tid&page[first]=1"
    try {
      $probe = Invoke-RestMethod -Uri $probeUrl -Headers $HeadersAuvik -Method GET
      $authorized.Add($tid)
    } catch {
      $httpError = $_.Exception.Response
      if ($httpError -and $httpError.StatusCode.value__ -eq 403) {
        Write-Host ("[FILTER] removing unauthorized tenant {0}" -f $tid)
        continue
      }
      throw
    }
  }
  return $authorized.ToArray()
}

# Firewall inventory (names/types/vendor/model) per tenant
function Get-FirewallInventoryForTenant {
  param([string]$TenantId)
  $q = @{
    "tenants"            = $TenantId
    "filter[deviceType]" = "firewall"
    "page[first]"        = 1000
  }
  $url = "$BaseAuvik/v1/inventory/device/info?" + (Build-Query $q)
  Write-Host ("Inventory URL (tenant {0}): {1}" -f $TenantId, $url)
  $data = Invoke-AuvikGet -Url $url
  $map  = @{}
  foreach ($d in $data) {
    $siteId = ""
    if ($d.relationships -and $d.relationships.site -and $d.relationships.site.data) {
      $siteId = $d.relationships.site.data.id
    }
    $map[$d.id] = @{
      deviceName = $d.attributes.deviceName
      deviceType = $d.attributes.deviceType
      vendorName = $d.attributes.vendorName
      makeModel  = $d.attributes.makeModel
      siteId     = $siteId
      siteName   = ""  # optional enrichment
    }
  }
  return $map
}

# Graph token (client creds)
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

# Resolve Drive via list->drive relationship
function Resolve-DriveByListId {
  param([string]$GraphToken,[string]$SiteHost,[string]$SitePath,[string]$ListId)
  $gh = @{ Authorization = "Bearer $GraphToken" }
  $siteUrl  = ("https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f $SiteHost, $SitePath)
  $site     = Invoke-RestMethod -Uri $siteUrl -Headers $gh -Method GET
  $driveUrl = ("https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/drive" -f $site.id, $ListId)
  $driveObj = Invoke-RestMethod -Uri $driveUrl -Headers $gh -Method GET
  return @{ driveId = $driveObj.id; driveName = $driveObj.name; siteWebUrl = $site.webUrl }
}

# Ensure folder path exists
function Ensure-FolderPath {
  param([string]$GraphToken,[string]$DriveId,[string]$FolderPath)
  $ghJson = @{ Authorization = "Bearer $GraphToken"; "Content-Type" = "application/json" }
  $gh     = @{ Authorization = "Bearer $GraphToken" }
  $segments = ($FolderPath.Trim()) -split '/'
  $segments = $segments | Where-Object { $_ -ne '' }
  $currentPath = ""
  $parentId = "root"
  foreach ($seg in $segments) {
    $currentPath = if ($currentPath) { "$currentPath/$seg" } else { $seg }
    $getUrl = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}" -f $DriveId, $currentPath)
    try { $item = Invoke-RestMethod -Uri $getUrl -Headers $gh -Method GET; $parentId = $item.id }
    catch {
      $createUrl = ("https://graph.microsoft.com/v1.0/drives/{0}/items/{1}/children" -f $DriveId, $parentId)
      $body = @{ name = $seg; folder = @{}; "@microsoft.graph.conflictBehavior" = "fail" } | ConvertTo-Json
      $newItem = Invoke-RestMethod -Uri $createUrl -Headers $ghJson -Method POST -Body $body
      $parentId = $newItem.id
      Write-Host ("Created folder segment: {0}" -f $seg)
    }
  }
  return $currentPath
}

# Upload CSV via PUT /content  (Patch B: idempotent overwrite + retry on 409)
function Upload-CsvToDrive {
  param([string]$GraphToken,[string]$DriveId,[string]$FolderPath,[string]$FileName,[string]$CsvText)
  $headers = @{ Authorization = "Bearer $GraphToken"; "Content-Type" = "text/csv" }
  $createdPath = Ensure-FolderPath -GraphToken $GraphToken -DriveId $DriveId -FolderPath $FolderPath

  # Ensure overwrite behavior for concurrent writes
  $uploadUrl  = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}:/content?@microsoft.graph.conflictBehavior=replace" -f $DriveId, $createdPath, $FileName)
  Write-Host "Uploading CSV to: $uploadUrl"

  $maxRetries = 3; $attempt = 0
  do {
    try {
      Invoke-RestMethod -Uri $uploadUrl -Method PUT -Headers $headers -Body $CsvText | Out-Null
      return
    } catch {
      $resp = $_.Exception.Response
      if ($resp -and $resp.StatusCode.value__ -eq 409 -and $attempt -lt $maxRetries) {
        Start-Sleep -Seconds 2
        $attempt++
        Write-Host "[RETRY] 409 Conflict on upload; attempt $attempt of $maxRetries"
        continue
      }
      throw
    }
  } while ($true)
}

# -------------------------------------------------------------
# Auvik — Auth verify
# -------------------------------------------------------------
Test-AuvikAuth

# -------------------------------------------------------------
# Tenants — discover or use provided, then pre-filter to authorized IDs
# -------------------------------------------------------------
$tenantList     = @()
$tenantNameMap  = @{}
if ([string]::IsNullOrWhiteSpace($TenantsCsv)) {
  $discovered = Get-AuvikTenants
  foreach ($t in $discovered) { $tenantList += $t.id; $tenantNameMap[$t.id] = $t.name }
} else {
  $tenantList = $TenantsCsv.Split(',') | ForEach-Object { $_.Trim() }
  # Best-effort name map from discovery
  $discovered = Get-AuvikTenants
  foreach ($t in $discovered) { $tenantNameMap[$t.id] = $t.name }
}
if ($tenantList.Count -gt 0) {
  $tenantList = Filter-AuthorizedTenants -TenantIds $tenantList
  Write-Host ("Authorized tenants after filter: {0}" -f ($tenantList -join ','))
}
if ($tenantList.Count -eq 0) {
  Write-Host "No authorized tenants available for this API user; output may be empty."
}

# -------------------------------------------------------------
# Collect DEVICE availability (firewalls) across authorized tenants
# -------------------------------------------------------------
$rowsDevices = New-Object System.Collections.Generic.List[object]

foreach ($tenant in $tenantList) {
  $tenantName  = if ($tenantNameMap.ContainsKey($tenant)) { $tenantNameMap[$tenant] } else { "" }

  # Inventory enrichment (firewalls)
  $fwInv = Get-FirewallInventoryForTenant -TenantId $tenant

  # Loop device types (or single null if none specified)
  $deviceTypeLoop = if ($DeviceTypes.Count -gt 0) { $DeviceTypes } else { @($null) }

  foreach ($devType in $deviceTypeLoop) {
    $devTypeLabel = if ([string]::IsNullOrWhiteSpace($devType)) { "<none>" } else { $devType }

    # Common filters (aligned to midnight)
    $qFilters = @{
      "filter[fromTime]" = $FromStr
      "filter[thruTime]" = $ThruStr
      "filter[interval]" = $Interval
      "page[first]"      = 500
      "tenants"          = $tenant
    }
    if (-not [string]::IsNullOrWhiteSpace($devType)) { $qFilters["filter[deviceType]"] = $devType }

    # Device Availability endpoint with statId=uptime|outage
    $uptUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($qFilters + @{ "statId" = "uptime" }))
    $outUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($qFilters + @{ "statId" = "outage" }))
    Write-Host ("[DEV] Tenant {0} | devType {1}" -f $tenant, $devTypeLabel)
    Write-Host ("[DEV] Uptime URL:  {0}" -f $uptUrl)
    Write-Host ("[DEV] Outage URL:  {0}" -f $outUrl)

    $uptData = Invoke-AuvikGet -Url $uptUrl
    $outData = Invoke-AuvikGet -Url $outUrl
    Write-Host ("[DEV] Items: uptime={0} outage={1}" -f $uptData.Count, $outData.Count)

    # If filtered and still zero, retry without deviceType
    if (-not [string]::IsNullOrWhiteSpace($devType) -and $uptData.Count -eq 0 -and $outData.Count -eq 0) {
      $qFilters.Remove("filter[deviceType]")
      $uptUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($qFilters + @{ "statId" = "uptime" }))
      $outUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($qFilters + @{ "statId" = "outage" }))
      Write-Host ("[DEV] Retry without deviceType → Uptime URL: {0}" -f $uptUrl)
      Write-Host ("[DEV] Retry without deviceType → Outage URL: {0}" -f $outUrl)
      $uptData = Invoke-AuvikGet -Url $uptUrl
      $outData = Invoke-AuvikGet -Url $outUrl
      Write-Host ("[DEV] Items (no filter): uptime={0} outage={1}" -f $uptData.Count, $outData.Count)
    }

    # Build outage lookup
    $outLookup = @{}
    foreach ($o in $outData) {
      $devId = $o.relationships.device.data.id
      $tenId = $o.relationships.tenant.data.id
      $ts    = Convert-FromUnixMinutes([long]$o.attributes.time)
      $key   = "$tenId|$devId|$($ts.ToString('yyyy-MM-dd'))"
      $total = $o.attributes.total
      if ($null -eq $total) { $total = 0 }
      $outLookup[$key] = $total
    }

    # Map rows
    foreach ($u in $uptData) {
      $devId = $u.relationships.device.data.id
      $tenId = $u.relationships.tenant.data.id
      $ts    = Convert-FromUnixMinutes([long]$u.attributes.time)
      $key   = "$tenId|$devId|$($ts.ToString('yyyy-MM-dd'))"

      # If we have firewall inventory, only include those devices
      if ($fwInv.Count -gt 0 -and -not $fwInv.ContainsKey($devId)) { continue }

      $avg   = $u.attributes.average
      if ($null -eq $avg) { $avg = 0 }
      $uptPct= [math]::Round([double]$avg, 3)

      $outSec= if ($outLookup.ContainsKey($key)) { $outLookup[$key] } else { 0 }
      $outMin= [math]::Round(([double]$outSec / 60.0), 2)

      $inv   = if ($fwInv.ContainsKey($devId)) { $fwInv[$devId] } else { @{} }
      $deviceTypeOut = if ($inv.deviceType) { $inv.deviceType } else { if ([string]::IsNullOrWhiteSpace($devType)) { "" } else { $devType } }

      $rowsDevices.Add([pscustomobject]@{
        date           = $ts.ToString('yyyy-MM-dd')
        tenant_id      = $tenId
        tenant_name    = $tenantName
        site_id        = $inv.siteId
        site_name      = $inv.siteName
        device_id      = $devId
        device_name    = $inv.deviceName
        device_type    = $deviceTypeOut
        vendor_name    = $inv.vendorName
        model          = $inv.makeModel
        uptime_percent = $uptPct
        outage_minutes = $outMin
        interval       = $Interval
      })
    }
  }
}

Write-Host ("[DEV] Rows collected: {0}" -f $rowsDevices.Count)

# -------------------------------------------------------------
# Collect WAN Internet stats (service ping) across authorized tenants
# -------------------------------------------------------------
$rowsWan = New-Object System.Collections.Generic.List[object]

foreach ($tenant in $tenantList) {
  $tenantName  = if ($tenantNameMap.ContainsKey($tenant)) { $tenantNameMap[$tenant] } else { "" }

  $qBase = @{
    "filter[fromTime]" = $FromStr
    "filter[thruTime]" = $ThruStr
    "filter[interval]" = $Interval
    "page[first]"      = 500
    "tenants"          = $tenant
  }

  # Service stats (cloud ping)
  $rttUrl = "$BaseAuvik/v1/stat/service/pingTime?"   + (Build-Query ($qBase))   # RTT avg/min/max
  $pktUrl = "$BaseAuvik/v1/stat/service/pingPacket?" + (Build-Query ($qBase))   # packets TX/RX
  Write-Host ("[WAN] Tenant {0}" -f $tenant)
  Write-Host ("[WAN] RTT URL: {0}" -f $rttUrl)
  Write-Host ("[WAN] PKT URL: {0}" -f $pktUrl)

  $rttData = Invoke-AuvikGet -Url $rttUrl
  $pktData = Invoke-AuvikGet -Url $pktUrl
  Write-Host ("[WAN] Items: rtt={0} pkt={1}" -f $rttData.Count, $pktData.Count)

  # Build RTT lookup (per day)
  $rttLookup = @{}
  foreach ($r in $rttData) {
    $siteId = $r.relationships.site.data.id
    $tenId  = $r.relationships.tenant.data.id
    $ts     = Convert-FromUnixMinutes([long]$r.attributes.time)
    $key    = "$tenId|$siteId|$($ts.ToString('yyyy-MM-dd'))"
    $rttLookup[$key] = @{
      avg = $r.attributes.average
      max = $r.attributes.maximum
      min = $r.attributes.minimum
    }
  }

  # Map per-day packet TX/RX + internet uptime %
  foreach ($p in $pktData) {
    $siteId = $p.relationships.site.data.id
    $tenId  = $p.relationships.tenant.data.id
    $ts     = Convert-FromUnixMinutes([long]$p.attributes.time)
    $key    = "$tenId|$siteId|$($ts.ToString('yyyy-MM-dd'))"

    $tx = $p.attributes.transmitted; if ($null -eq $tx) { $tx = 0 }
    $rx = $p.attributes.received;    if ($null -eq $rx) { $rx = 0 }
    $tx = [double]$tx
    $rx = [double]$rx
    $upt = if ($tx -gt 0) { [math]::Round(($rx / $tx) * 100.0, 3) } else { 0.0 }

    $rtt = if ($rttLookup.ContainsKey($key)) { $rttLookup[$key] } else { @{ avg=$null; max=$null; min=$null } }
    $avgRtt = if ($null -ne $rtt.avg) { [math]::Round([double]$rtt.avg, 3) } else { "" }
    $maxRtt = if ($null -ne $rtt.max) { [math]::Round([double]$rtt.max, 3) } else { "" }
    $minRtt = if ($null -ne $rtt.min) { [math]::Round([double]$rtt.min, 3) } else { "" }

    $rowsWan.Add([pscustomobject]@{
      date                   = $ts.ToString('yyyy-MM-dd')
      tenant_id              = $tenId
      tenant_name            = $tenantName
      site_id                = $siteId
      site_name              = ""             # optional enrichment
      avg_rtt_ms             = $avgRtt
      max_rtt_ms             = $maxRtt
      min_rtt_ms             = $minRtt
      packets_tx             = $tx
      packets_rx             = $rx
      internet_uptime_percent= $upt
      interval               = $Interval
    })
  }
}

Write-Host ("[WAN] Rows collected: {0}" -f $rowsWan.Count)

# -------------------------------------------------------------
# CSV builders (BrightGauge-friendly: stable headers, no commas)
# -------------------------------------------------------------
# Devices CSV (firewalls)
$devHeaders  = @('date','tenant_id','tenant_name','site_id','site_name','device_id','device_name','device_type','vendor_name','model','uptime_percent','outage_minutes','interval')
$devCsvLines = New-Object System.Collections.Generic.List[string]
$devCsvLines.Add(($devHeaders -join ','))
foreach ($r in $rowsDevices) {
  $vals = foreach ($h in $devHeaders) {
    $val = if ($null -eq $r.$h) { "" } else { [string]$r.$h }
    $val -replace ',', ' '
  }
  $devCsvLines.Add(($vals -join ','))
}
$devCsv = $devCsvLines -join "`n"

# WAN CSV (cloud ping per site)
$wanHeaders  = @('date','tenant_id','tenant_name','site_id','site_name','avg_rtt_ms','max_rtt_ms','min_rtt_ms','packets_tx','packets_rx','internet_uptime_percent','interval')
$wanCsvLines = New-Object System.Collections.Generic.List[string]
$wanCsvLines.Add(($wanHeaders -join ','))
foreach ($r in $rowsWan) {
  $vals = foreach ($h in $wanHeaders) {
    $val = if ($null -eq $r.$h) { "" } else { [string]$r.$h }
    $val -replace ',', ' '
  }
  $wanCsvLines.Add(($vals -join ','))
}
$wanCsv = $wanCsvLines -join "`n"

# -------------------------------------------------------------
# SharePoint upload (both CSVs)
# -------------------------------------------------------------
$graphToken = Get-GraphToken
$driveInfo  = Resolve-DriveByListId -GraphToken $graphToken -SiteHost $SP_SiteHost -SitePath $SP_SitePath -ListId $SP_ListId

Upload-CsvToDrive -GraphToken $graphToken -DriveId $driveInfo.driveId -FolderPath $SP_FolderPath -FileName $OutputFileDev -CsvText $devCsv
Write-Host ("[DEV] Uploaded {0} rows to drive '{1}' at {2}/{3}" -f $rowsDevices.Count, $driveInfo.driveName, $SP_FolderPath, $OutputFileDev)

Upload-CsvToDrive -GraphToken $graphToken -DriveId $driveInfo.driveId -FolderPath $SP_FolderPath -FileName $OutputFileWan -CsvText $wanCsv
Write-Host ("[WAN] Uploaded {0} rows to drive '{1}' at {2}/{3}" -f $rowsWan.Count, $driveInfo.driveName, $SP_FolderPath, $OutputFileWan)
