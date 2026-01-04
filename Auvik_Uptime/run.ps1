
param($Timer)

<#
  Azure Functions (PowerShell 5.1) — Timer Trigger
  Multi-tenant Auvik → CSVs (firewall uptime, WAN internet uptime) → SharePoint via Graph

  References:
   - Auvik API Integration Guide (regional host & Basic auth; TLS 1.2)            [1](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)
   - Auvik Statistics – Device Availability (uptime %, outage seconds)            [2](https://elischei.com/how-to-get-site-id-with-graph-explorer-and-other-sharepoint-info/)
   - Auvik Statistics – Service (cloud ping RTT, packets transmitted/received)    [3](https://robwindsor.hashnode.dev/access-a-sharepoint-site-by-server-relative-url-with-microsoft-graph)
   - BrightGauge CSV requirements (UTC dates, stable headers, no commas)          [4](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/v1.0/api/drive-get.md)
   - Microsoft Graph upload (PUT /content to OneDrive/SharePoint)                 [5](https://www.fortinet.com/content/dam/fortinet/assets/alliances/sb-fortinet-auvik.pdf)
#>

# -------------------------------------------------------------
# TLS 1.2
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
$AuvikApiKey    = Get-EnvVal "AUVIK_API_KEY"
$TenantsCsv     = Get-EnvVal "AUVIK_TENANTS" ""             # empty = auto-discover
$DeviceTypesCsv = Get-EnvVal "AUVIK_DEVICE_TYPES" "firewall"
$DeviceTypes    = if ([string]::IsNullOrWhiteSpace($DeviceTypesCsv)) { @() } else { $DeviceTypesCsv.Split(',') | ForEach-Object { $_.Trim() } }
$WindowDays     = if ($env:WINDOW_DAYS) { [int]$env:WINDOW_DAYS } else { 30 }
$Interval       = Get-EnvVal "AUVIK_INTERVAL" "day"

$FromUtc        = (Get-Date).ToUniversalTime().AddDays(-$WindowDays)
$ThruUtc        = (Get-Date).ToUniversalTime()

# Regional host + Basic
$BaseAuvik      = "https://auvikapi.$AuvikRegion.my.auvik.com"
$BasicToken     = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$AuvikUser:$AuvikApiKey"))
$HeadersAuvik   = @{ Authorization = "Basic $BasicToken"; Accept = "application/json" }

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

# Inspect HTTP status for verify (empty body is common on 200)
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

# Auvik GET with paging
function Invoke-AuvikGet {
  param([string]$Url)
  $results = @()
  $next = $Url
  do {
    $resp = Invoke-RestMethod -Uri $next -Headers $HeadersAuvik -Method GET
    if ($resp.data) { $results += $resp.data }
    $next = $resp.links.next
  } while ($next)
  return $results
}

# Discover tenants; returns @{ id=..., name=... }[] and a map
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

# Firewall inventory (name/type/vendor/model) per tenant
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
    $map[$d.id] = @{
      deviceName = $d.attributes.deviceName
      deviceType = $d.attributes.deviceType
      vendorName = $d.attributes.vendorName
      makeModel  = $d.attributes.makeModel
      siteId     = if ($d.relationships.site.data) { $d.relationships.site.data.id } else { "" }
      siteName   = ""  # optional enrichment if needed later
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

# Upload CSV via PUT /content
function Upload-CsvToDrive {
  param([string]$GraphToken,[string]$DriveId,[string]$FolderPath,[string]$FileName,[string]$CsvText)
  $headers = @{ Authorization = "Bearer $GraphToken"; "Content-Type" = "text/csv" }
  $createdPath = Ensure-FolderPath -GraphToken $GraphToken -DriveId $DriveId -FolderPath $FolderPath
  $uploadUrl  = ("https://graph.microsoft.com/v1.0/drives/{0}/root:/{1}/{2}:/content" -f $DriveId, $createdPath, $FileName)
  Write-Host "Uploading CSV to: $uploadUrl"
  Invoke-RestMethod -Uri $uploadUrl -Method PUT -Headers $headers -Body $CsvText | Out-Null
}

# -------------------------------------------------------------
# Auvik — Auth verify
# -------------------------------------------------------------
Test-AuvikAuth

# -------------------------------------------------------------
# Tenants — discover or use provided
# -------------------------------------------------------------
$tenantList = @()
$tenantNameMap = @{}

if ([string]::IsNullOrWhiteSpace($TenantsCsv)) {
  $discovered = Get-AuvikTenants
  if ($discovered.Count -eq 0) {
    Write-Host "No tenants discovered; running unscoped (may return empty if user not authorized)."
    $tenantList = @($null)
  } else {
    foreach ($t in $discovered) {
      $tenantList += $t.id
      $tenantNameMap[$t.id] = $t.name
    }
  }
} else {
  $tenantList = $TenantsCsv.Split(',') | ForEach-Object { $_.Trim() }
  # Try to fetch names for provided IDs (best-effort)
  $discovered = Get-AuvikTenants
  foreach ($t in $discovered) { $tenantNameMap[$t.id] = $t.name }
}

# -------------------------------------------------------------
# Collect DEVICE availability (firewalls) across all tenants
# -------------------------------------------------------------
$rowsDevices = New-Object System.Collections.Generic.List[object]

foreach ($tenant in $tenantList) {
  $tenantLabel = if ($tenant) { $tenant } else { "<none>" }
  $tenantName  = if ($tenant -and $tenantNameMap.ContainsKey($tenant)) { $tenantNameMap[$tenant] } else { "" }

  # Inventory enrichment (firewalls)
  $fwInv = if ($tenant) { Get-FirewallInventoryForTenant -TenantId $tenant } else { @{} }

  # Loop device types (or single null if none specified)
  $deviceTypeLoop = if ($DeviceTypes.Count -gt 0) { $DeviceTypes } else { @($null) }

  foreach ($devType in $deviceTypeLoop) {

    $qBase = @{
      "filter[fromTime]" = $FromUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
      "filter[thruTime]" = $ThruUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
      "filter[interval]" = $Interval
      "page[first]"      = 500
    }
    if ($tenant)  { $qBase["tenants"] = $tenant }
    if ($devType) { $qBase["filter[deviceType]"] = $devType }

    $uptUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($qBase + @{ "statId" = "uptime" }))
    $outUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($qBase + @{ "statId" = "outage" }))
    Write-Host ("[DEV] Tenant {0} | devType {1}" -f $tenantLabel, ($devType ?? "<none>"))
    Write-Host ("[DEV] Uptime URL:  {0}" -f $uptUrl)
    Write-Host ("[DEV] Outage URL:  {0}" -f $outUrl)

    $uptData = Invoke-AuvikGet -Url $uptUrl
    $outData = Invoke-AuvikGet -Url $outUrl
    Write-Host ("[DEV] Items: uptime={0} outage={1}" -f $uptData.Count, $outData.Count)

    # Build outage lookup
    $outLookup = @{}
    foreach ($o in $outData) {
      $devId = $o.relationships.device.data.id
      $tenId = $o.relationships.tenant.data.id
      $ts    = Convert-FromUnixMinutes([long]$o.attributes.time)
      $key   = "$tenId|$devId|$($ts.ToString('yyyy-MM-dd'))"
      $total = $o.attributes.total; if ($null -eq $total) { $total = 0 }
      $outLookup[$key] = $total
    }

    # Map rows (only firewalls if inventory map present)
    foreach ($u in $uptData) {
      $devId = $u.relationships.device.data.id
      $tenId = $u.relationships.tenant.data.id
      $ts    = Convert-FromUnixMinutes([long]$u.attributes.time)
      $key   = "$tenId|$devId|$($ts.ToString('yyyy-MM-dd'))"

      # If we have inventory, include only those devices (firewalls). If no inventory, include all (devType filter already applied).
      if ($fwInv.Count -gt 0 -and -not $fwInv.ContainsKey($devId)) { continue }

      $avg   = $u.attributes.average; if ($null -eq $avg) { $avg = 0 }
      $uptPct= [math]::Round($avg, 3)
      $outSec= if ($outLookup.ContainsKey($key)) { $outLookup[$key] } else { 0 }
      $outMin= [math]::Round(($outSec / 60.0), 2)

      $inv   = if ($fwInv.ContainsKey($devId)) { $fwInv[$devId] } else { @{} }

      $rowsDevices.Add([pscustomobject]@{
        date           = $ts.ToString('yyyy-MM-dd')
        tenant_id      = $tenId
        tenant_name    = $tenantName
        site_id        = $inv.siteId
        site_name      = $inv.siteName
        device_id      = $devId
        device_name    = $inv.deviceName
        device_type    = if ($inv.deviceType) { $inv.deviceType } else { ($devType ?? "") }
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
# Collect WAN Internet stats (service ping) across all tenants
# -------------------------------------------------------------
$rowsWan = New-Object System.Collections.Generic.List[object]

foreach ($tenant in $tenantList) {
  $tenantLabel = if ($tenant) { $tenant } else { "<none>" }
  $tenantName  = if ($tenant -and $tenantNameMap.ContainsKey($tenant)) { $tenantNameMap[$tenant] } else { "" }

  $qBase = @{
    "filter[fromTime]" = $FromUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    "filter[thruTime]" = $ThruUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    "filter[interval]" = $Interval
    "page[first]"      = 500
  }
  if ($tenant) { $qBase["tenants"] = $tenant }

  # RTT stats
  $rttUrl = "$BaseAuvik/v1/stat/service/pingTime?" + (Build-Query ($qBase))
  # Packet stats (transmitted/received)
  $pktUrl = "$BaseAuvik/v1/stat/service/pingPacket?" + (Build-Query ($qBase))
  Write-Host ("[WAN] Tenant {0}" -f $tenantLabel)
  Write-Host ("[WAN] RTT URL: {0}" -f $rttUrl)
