
param($Timer)

<#
  Azure Functions (PowerShell 5.1) — Timer Trigger
  Pipeline: Auvik Stats API (device availability) --> CSV --> SharePoint (Drive) via Microsoft Graph

  References:
   - Auvik Statistics API (Device Availability: uptime%, outage seconds; filters + paging)  [support]  [1](https://learn.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)
   - Auvik API regional hostname & Basic auth (username: apiKey)                            [support]  [2](https://sposcripts.com/how-to-upload-files-to-sharepoint-using-graph-api/)
   - Microsoft Graph files model (Drive/DriveItem, PUT /content)                           [docs]     [3](https://vectorlinux.com/how-to-upload-file-in-onedrive-in-c/)[4](https://docs.1stream.com/551377-brightgauge/brightgauge-integration/version/1?kb_language=en_US)
   - Resolve SharePoint site by path; list site drives; match drive via sharepointIds.listId [docs]    [5](https://support.brightgauge.com/hc/en-us/articles/204473769-How-to-Upload-a-CSV-Dataset-Dropbox-or-OneDrive?mobile_site=true)[6](https://documentation.n-able.com/passportal/userguide/Content/rmm-integ/Auvik-Integ.html)[7](https://www.techguy.at/upload-a-file-to-onedrive-via-graph-api-and-powershell/)
   - BrightGauge CSV requirements (date format, headers, numerics)                         [docs]     [8](https://stackoverflow.com/questions/41285403/upload-file-to-sharepoint-drive-using-microsoft-graph)
#>

# -------------------------------------------------------------
# Enforce TLS 1.2 for Auvik API calls (per their guidance)
# -------------------------------------------------------------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  # [2](https://sposcripts.com/how-to-upload-files-to-sharepoint-using-graph-api/)

# -------------------------------------------------------------
# Configuration (env vars) — PS 5.1 safe
# -------------------------------------------------------------
function Get-EnvVal([string]$name, [string]$default = "") {
  $raw = [System.Environment]::GetEnvironmentVariable($name)
  if ([string]::IsNullOrWhiteSpace($raw)) { return $default }
  return $raw.Trim()
}

$AuvikRegion       = Get-EnvVal "AUVIK_REGION"
$AuvikUser         = Get-EnvVal "AUVIK_USERNAME"
$AuvikApiKey       = Get-EnvVal "AUVIK_API_KEY"
$TenantsCsv        = Get-EnvVal "AUVIK_TENANTS"
$Interval          = Get-EnvVal "AUVIK_INTERVAL"
$DeviceTypesCsv    = Get-EnvVal "AUVIK_DEVICE_TYPES"

$WindowDays = if ($env:WINDOW_DAYS) { [int]$env:WINDOW_DAYS } else { 30 }

$TenantId          = Get-EnvVal "Ms365_TenantId"
$ClientId          = Get-EnvVal "Ms365_AuthAppId"
$ClientSecret      = Get-EnvVal "Ms365_AuthSecretId"

$SP_SiteHost       = Get-EnvVal "SP_SiteHost"       # e.g., palaisparc.sharepoint.com
$SP_SitePath       = Get-EnvVal "SP_SitePath"       # e.g., /automation
$SP_ListId         = Get-EnvVal "SP_ListId"         # document library listId (GUID)
$SP_FolderPath     = Get-EnvVal "SP_FolderPath"     # e.g., /Reports/Uptime
$OutputFileName    = Get-EnvVal "OUTPUT_FILE_NAME"  # e.g., auvik-uptime-month.csv

# -------------------------------------------------------------
# Diagnostics — show key env values (no secrets)
# -------------------------------------------------------------
Write-Host "---------- ENV CHECK ----------"
Write-Host ("Auvik Region      : {0}" -f $AuvikRegion)
Write-Host ("Auvik Username    : {0}" -f $AuvikUser)
Write-Host ("Auvik Tenants     : {0}" -f $TenantsCsv)
Write-Host ("Auvik Interval    : {0}" -f (if ($Interval) { $Interval } else { 'day' }))
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
# Helpers
# -------------------------------------------------------------

# Regional Auvik base (use auvikapi.{region}.my.auvik.com, not dashboard subdomain)  [2](https://sposcripts.com/how-to-upload-files-to-sharepoint-using-graph-api/)
$BaseAuvik = ("https://auvikapi.{0}.my.auvik.com" -f $AuvikRegion)

# Basic auth header (username:apiKey)
$AuthHeader   = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${AuvikUser}:${AuvikApiKey}"))  # [9](https://martin-machacek.com/blogPost/e632292b-c6d4-4819-b68f-2953f485dd13)
$HeadersAuvik = @{ Authorization = $AuthHeader; Accept = 'application/json' }

# Time window & interval
$FromUtc   = (Get-Date).ToUniversalTime().AddDays(-$WindowDays)
$ThruUtc   = (Get-Date).ToUniversalTime()
$Interval  = if ($Interval) { $Interval } else { 'day' }

# Lists for tenants/device types
$Tenants     = if ([string]::IsNullOrWhiteSpace($TenantsCsv))    { @($null) } else { $TenantsCsv.Split(',')     | ForEach-Object { $_.Trim() } }
$DeviceTypes = if ([string]::IsNullOrWhiteSpace($DeviceTypesCsv)) { @()       } else { $DeviceTypesCsv.Split(',') | ForEach-Object { $_.Trim() } }

# Convert unix minutes (Auvik stats time) to DateTime (UTC day bucket)  [1](https://learn.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)
function Convert-FromUnixMinutes([long]$unixMinutes) {
  $secs = $unixMinutes * 60L
  return [DateTimeOffset]::FromUnixTimeSeconds($secs).UtcDateTime.Date
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

# Safe querystring builder
function Build-Query {
  param([hashtable]$kv)
  $pairs = $kv.GetEnumerator() | ForEach-Object {
    "{0}={1}" -f $_.Key, [System.Web.HttpUtility]::UrlEncode([string]$_.Value)
  }
  ($pairs -join '&')
}

# Graph token (client credentials) — v2.0 token endpoint (requires admin-consented app permissions)  [10](https://docs.1stream.com/en_US/551377-brightgauge/importing-csv-files-in-brightgauge)
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

# Resolve SharePoint Drive by site path & listId (document library)  [5](https://support.brightgauge.com/hc/en-us/articles/204473769-How-to-Upload-a-CSV-Dataset-Dropbox-or-OneDrive?mobile_site=true)[6](https://documentation.n-able.com/passportal/userguide/Content/rmm-integ/Auvik-Integ.html)
function Resolve-DriveByListId {
  param(
    [string]$GraphToken,
    [string]$SiteHost,
    [string]$SitePath,
    [string]$ListId
  )
  $gh = @{ Authorization = "Bearer $GraphToken" }

  # 1) Get site by path (hostname + server-relative path)
  $siteUrl = ("https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f $SiteHost, $SitePath)
  $site    = Invoke-RestMethod -Uri $siteUrl -Headers $gh -Method GET

  # 2) List drives for site; select sharepointIds so we can match listId
  $drivesUrl = ("https://graph.microsoft.com/v1.0/sites/{0}/drives?$select=id,name,driveType,sharepointIds" -f $site.id)
  $drives    = Invoke-RestMethod -Uri $drivesUrl -Headers $gh -Method GET

  $target = $null
  foreach ($d in $drives.value) {
    $lid = $d.sharepointIds.listId  # present because of $select sharepointIds  [7](https://www.techguy.at/upload-a-file-to-onedrive-via-graph-api-and-powershell/)
    if ($lid -and ($lid -eq $ListId)) { $target = $d; break }
  }

  if (-not $target) {
    Write-Host "No drive matched listId '$ListId'. Drives under site '$($site.webUrl)':"
    $drives.value | ForEach-Object { Write-Host ("- {0} (type: {1}) listId: {2}" -f $_.name, $_.driveType, $_.sharepointIds.listId) }
    throw "Drive resolution failed."
  }

  return @{ driveId = $target.id; siteWebUrl = $site.webUrl; driveName = $target.name }
}

# Upload CSV to the resolved drive via PUT /content (supports up to 250 MB)  [4](https://docs.1stream.com/551377-brightgauge/brightgauge-integration/version/1?kb_language=en_US)
function Upload-CsvToDrive {
  param(
    [string]$GraphToken,
    [string]$DriveId,
    [string]$FolderPath,
    [string]$FileName,
    [string]$CsvText
  )
  $headers    = @{ Authorization = "Bearer $GraphToken"; "Content-Type" = "text/csv" }
  $safeFolder = if ($FolderPath.StartsWith("/")) { $FolderPath } else { "/$FolderPath" }

  $uploadUrl  = ("https://graph.microsoft.com/v1.0/drives/{0}/root:{1}/{2}:/content" -f $DriveId, $safeFolder, $FileName)
  Write-Host "Uploading CSV to: $uploadUrl"
  Invoke-RestMethod -Uri $uploadUrl -Method PUT -Headers $headers -Body $CsvText | Out-Null
}

# -------------------------------------------------------------
# Collect Auvik availability stats (uptime% + outage seconds)
# -------------------------------------------------------------
$rows = New-Object System.Collections.Generic.List[object]

foreach ($tenant in $Tenants) {

  $deviceTypeLoop = if ($DeviceTypes.Count -gt 0) { $DeviceTypes } else { @($null) }

  foreach ($devType in $deviceTypeLoop) {

    $q = @{
      "filter[fromTime]" = $FromUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
      "filter[thruTime]" = $ThruUtc.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
      "filter[interval]" = $Interval
      "page[first]"      = 500
    }
    if ($tenant)  { $q["tenants"] = $tenant }
    if ($devType) { $q["filter[deviceType]"] = $devType }  # e.g., firewall, router  [11](https://learn.microsoft.com/en-us/answers/questions/1189234/clarification-on-sharepoint-composite-site-id)

    # Uptime %
    $uptUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($q + @{ "statId" = "uptime" }))
    $upt    = Invoke-AuvikGet -Url $uptUrl  # .data: attributes.average, relationships (device, tenant)  [1](https://learn.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)

    # Outage (seconds)
    $outUrl = "$BaseAuvik/v1/stat/device/availability?" + (Build-Query ($q + @{ "statId" = "outage" }))
    $out    = Invoke-AuvikGet -Url $outUrl  # .data: attributes.total, relationships (device, tenant)      [1](https://learn.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0)

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
        date           = $ts.ToString('yyyy-MM-dd')  # BrightGauge date (UTC)  [8](https://stackoverflow.com/questions/41285403/upload-file-to-sharepoint-drive-using-microsoft-graph)
        tenant_id      = $tenId
        tenant_name    = ""                           # optional: populate via Inventory if desired
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
    $val -replace ',', ' '  # BG: no commas in numbers/fields  [8](https://stackoverflow.com/questions/41285403/upload-file-to-sharepoint-drive-using-microsoft-graph)
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
