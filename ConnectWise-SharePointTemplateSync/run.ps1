#requires -Version 7.2
using namespace System.Net

# One-way SharePoint -> ConnectWise push for recurring service templates and schedules.
# Called by Power Automate when an item in a "CW - <Dept> - Templates/Schedules" list is created, modified or deleted.
# Body: { "event": "created|modified|deleted", "siteUrl": "...", "listName": "...", "itemId": 12, "cwId": 123 (deleted only), "dryRun": true (optional) }

param($Request, $TriggerMetadata)

Set-StrictMode -Version 1
$ErrorActionPreference = 'Stop'

$script:CorrelationId = [guid]::NewGuid().ToString()
$script:Ops = [System.Collections.Generic.List[string]]::new()
$script:Warnings = [System.Collections.Generic.List[string]]::new()

# ---------------------------------------------------------------- settings
function Get-Setting {
    param([string]$Name, [string]$Default = '')
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v
}
function Test-SettingTrue {
    param([string]$Name, [string]$Default = 'false')
    return ((Get-Setting $Name $Default).Trim().ToLowerInvariant() -in @('true', '1', 'yes', 'on'))
}

# Department sites and lists (allowlist: requests for anything else are rejected)
$script:Departments = @{
    'Service Desk'       = 'https://palaisparc.sharepoint.com/Service/servicedesk'
    'Finance'            = 'https://palaisparc.sharepoint.com/finance'
    'Engineering'        = 'https://palaisparc.sharepoint.com/Service/Engineering'
    'Account Management' = 'https://palaisparc.sharepoint.com/sales/AccountManagement'
    'Sales'              = 'https://palaisparc.sharepoint.com/sales'
    'Marketing'          = 'https://palaisparc.sharepoint.com/Marketing'
    'Internal IT'        = 'https://palaisparc.sharepoint.com/it'
    'Human Resources'    = 'https://palaisparc.sharepoint.com/hr'
}
$script:CwDeptToSp = @{
    'Engineering' = 'Engineering'; 'Engineering and Project Management' = 'Engineering'
    'Account Management' = 'Account Management'; 'CW Sales' = 'Account Management'
    'Internal IT' = 'Internal IT'; 'Information Technology' = 'Internal IT'
    'Human Resources' = 'Human Resources'; 'Human Resource' = 'Human Resources'; 'HR' = 'Human Resources'
}
$script:RegistrySite = 'https://palaisparc.sharepoint.com/it'
$script:RegistryList = 'CW - Recurring Work Sync Registry'

# ---------------------------------------------------------------- helpers
function New-RequestException {
    param([int]$StatusCode, [string]$Message)
    $ex = [System.Exception]::new($Message)
    $ex.Data['HttpStatusCode'] = $StatusCode
    return $ex
}
function Block { param([string]$Message) throw (New-RequestException 422 $Message) }

function ConvertTo-Text {
    param($Value)
    if ($null -eq $Value) { return '' }
    $tn = $Value.GetType().FullName
    if ($tn -in @('Microsoft.SharePoint.Client.FieldLookupValue', 'Microsoft.SharePoint.Client.FieldUserValue')) { return [string]$Value.LookupValue }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    $s = [string]$Value
    return ($s -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', '').Trim()
}
function ConvertTo-IntOrNull {
    param($Value)
    $t = ConvertTo-Text $Value
    if ($t -eq '') { return $null }
    $n = 0
    if ([int]::TryParse(($t -replace '\.0+$', ''), [ref]$n)) { return $n }
    Block "Value '$t' is not a whole number."
}
function ConvertTo-BoolOrNull {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return $Value }
    $t = (ConvertTo-Text $Value).ToLowerInvariant()
    if ($t -eq '') { return $null }
    if ($t -in @('true', 'yes', '1')) { return $true }
    if ($t -in @('false', 'no', '0')) { return $false }
    Block "Value '$t' is not a Yes/No."
}
function Get-Sha256 {
    param([string]$Text)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    return ([System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::HashData($bytes)) -replace '-', '')
}
function Get-Field {
    # Reads a SharePoint field value, returning $null when the column is not on the item.
    param($Item, [string]$Name)
    if ($Item.FieldValues.ContainsKey($Name)) { return $Item.FieldValues[$Name] }
    return $null
}

# ---------------------------------------------------------------- ConnectWise IO
$script:CwHttpClient = $null
function Send-CWHttp {
    # HttpClient with redirects disabled: ConnectWise answers POST with 201, the created object in the body, and an http:// Location
    # header that Invoke-RestMethod refuses to follow (it would report failure for a write that actually succeeded).
    param([string]$Method, [string]$Uri, [hashtable]$Headers, [string]$BodyJson)
    if ($global:CwHttpOverride) { return & $global:CwHttpOverride $Method $Uri $BodyJson }
    if (-not $script:CwHttpClient) {
        $handler = [System.Net.Http.HttpClientHandler]::new()
        $handler.AllowAutoRedirect = $false
        $script:CwHttpClient = [System.Net.Http.HttpClient]::new($handler)
        $script:CwHttpClient.Timeout = [TimeSpan]::FromSeconds(60)
    }
    $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::new($Method), $Uri)
    foreach ($k in $Headers.Keys) { $null = $req.Headers.TryAddWithoutValidation($k, [string]$Headers[$k]) }
    if ($BodyJson) { $req.Content = [System.Net.Http.StringContent]::new($BodyJson, [Text.Encoding]::UTF8, 'application/json') }
    $resp = $script:CwHttpClient.SendAsync($req).GetAwaiter().GetResult()
    return [pscustomobject]@{ Status = [int]$resp.StatusCode; Content = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult() }
}
function Invoke-CWRequest {
    param([string]$Method, [string]$Path, $Body = $null, [switch]$AllowNotFound)
    $base = Get-Setting 'ConnectWisePsa_ApiBaseUrl' 'https://api-na.myconnectwise.net/v4_6_release/apis/3.0'
    $company = Get-Setting 'ConnectWisePsa_ApiCompanyId'
    $pub = Get-Setting 'ConnectWisePsa_ApiPublicKey'
    $priv = Get-Setting 'ConnectWisePsa_ApiPrivateKey'
    $clientId = Get-Setting 'ConnectWisePsa_ApiClientId'
    if (-not ($company -and $pub -and $priv -and $clientId)) { throw (New-RequestException 500 'ConnectWise API settings are incomplete.') }
    $version = Get-Setting 'ConnectWisePsa_ApiVersion' '2025.1'
    $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$company+$pub`:$priv"))
    $headers = @{ Authorization = "Basic $auth"; clientId = $clientId; Accept = "application/vnd.connectwise.com+json; version=$version" }
    # -InputObject keeps a one-operation JSON Patch as an array (piping would unroll it into an object)
    $json = if ($null -ne $Body) { ConvertTo-Json -InputObject $Body -Depth 20 -Compress } else { $null }
    $attempts = if ($Method -eq 'GET') { 3 } else { 1 } # never auto-retry writes
    for ($i = 1; $i -le $attempts; $i++) {
        $code = 0; $content = ''
        try { $r = Send-CWHttp -Method $Method -Uri "$base/$Path" -Headers $headers -BodyJson $json; $code = $r.Status; $content = "$($r.Content)" }
        catch { $content = $_.Exception.Message }
        if ($code -ge 200 -and $code -lt 300) {
            if ([string]::IsNullOrWhiteSpace($content)) { return $null }
            return ($content | ConvertFrom-Json)
        }
        if ($code -eq 404 -and $AllowNotFound) { return $null }
        if ($i -lt $attempts -and ($code -in 429, 500, 502, 503, 504 -or $code -eq 0)) { Start-Sleep -Seconds (2 * $i); continue }
        $detail = $content
        if ($detail.Length -gt 500) { $detail = $detail.Substring(0, 500) }
        throw (New-RequestException 502 "ConnectWise $Method $Path failed (HTTP $code): $detail")
    }
}
function Get-CWAll {
    param([string]$Path, [string]$Query = '')
    $all = [System.Collections.Generic.List[object]]::new()
    $page = 1
    while ($true) {
        $sep = if ($Path.Contains('?')) { '&' } else { '?' }
        $q = "pageSize=1000&page=$page"
        if ($Query) { $q += "&$Query" }
        $r = @(Invoke-CWRequest -Method GET -Path "$Path$sep$q")
        foreach ($x in $r) { $all.Add($x) }
        if ($r.Count -lt 1000) { break }
        $page++
    }
    return , $all.ToArray()
}
function Find-ByName {
    param($Rows, [string]$Name, [string]$What)
    $hits = @($Rows | Where-Object { $_.name -ieq $Name })
    if ($hits.Count -eq 0) { Block "$What '$Name' was not found in ConnectWise." }
    if ($hits.Count -gt 1) { Block "$What '$Name' is ambiguous in ConnectWise ($($hits.Count) matches)." }
    return $hits[0]
}

# ---------------------------------------------------------------- SharePoint IO
$script:SpConnections = @{}
function Get-SpConnection {
    param([string]$Url)
    if ($script:SpConnections.ContainsKey($Url)) { return $script:SpConnections[$Url] }
    $tenant = Get-Setting 'Ms365_TenantId'
    $app = Get-Setting 'Ms365_AuthAppId'
    $cert = Get-Setting 'Ms365_CertBase64'
    $pwd = Get-Setting 'Ms365_CertPassword'
    if (-not ($tenant -and $app -and $cert)) { throw (New-RequestException 500 'SharePoint (Ms365_*) settings are incomplete.') }
    $secure = ConvertTo-SecureString -String $pwd -AsPlainText -Force
    $c = Connect-PnPOnline -Url $Url -ClientId $app -Tenant $tenant -CertificateBase64Encoded $cert -CertificatePassword $secure -ReturnConnection
    $script:SpConnections[$Url] = $c
    return $c
}
function Get-SpItem {
    param([string]$Site, [string]$List, [int]$Id)
    return Get-PnPListItem -List $List -Id $Id -Connection (Get-SpConnection $Site)
}
function Get-SpFieldNames {
    param([string]$Site, [string]$List)
    return @(Get-PnPField -List $List -Connection (Get-SpConnection $Site) | ForEach-Object { $_.InternalName })
}
function Add-SpTextField {
    param([string]$Site, [string]$List, [string]$Name, [string]$Type = 'Text')
    $null = Add-PnPField -List $List -InternalName $Name -DisplayName $Name -Type $Type -AddToDefaultView:$false -Connection (Get-SpConnection $Site)
}
function Set-SpItem {
    # SystemUpdate keeps Modified/Editor unchanged so write-backs do not re-trigger the flow.
    param([string]$Site, [string]$List, [int]$Id, [hashtable]$Values, [string[]]$Columns)
    $v = @{}
    foreach ($k in $Values.Keys) { if ($k -in $Columns) { $v[$k] = $Values[$k] } }
    if ($v.Count -eq 0) { return }
    $null = Set-PnPListItem -List $List -Identity $Id -Values $v -UpdateType SystemUpdate -Connection (Get-SpConnection $Site)
}
function Find-SpItemsByField {
    param([string]$Site, [string]$List, [string]$Field, [string]$Value)
    $safe = [System.Security.SecurityElement]::Escape($Value)
    $caml = "<View><Query><Where><Eq><FieldRef Name='$Field'/><Value Type='Text'>$safe</Value></Eq></Where></Query></View>"
    return @(Get-PnPListItem -List $List -Query $caml -Connection (Get-SpConnection $Site))
}

# ---------------------------------------------------------------- CW lookups (cached)
$script:Cache = @{}
function Get-Cached {
    param([string]$Key, [scriptblock]$Loader)
    if (-not $script:Cache.ContainsKey($Key)) { $script:Cache[$Key] = & $Loader }
    return , $script:Cache[$Key]
}
function Resolve-Board {
    param([string]$Name)
    $rows = Get-Cached 'boards' { Get-CWAll 'service/boards' }
    return Find-ByName $rows $Name 'Board'
}
function Resolve-BoardChild {
    param([int]$BoardId, [string]$Kind, [string]$Name)
    $rows = Get-Cached "board$BoardId-$Kind" { Get-CWAll "service/boards/$BoardId/$Kind" }
    return Find-ByName $rows $Name $Kind
}
function Resolve-Simple {
    param([string]$Path, [string]$Name, [string]$What)
    $rows = Get-Cached $Path { Get-CWAll $Path }
    return Find-ByName $rows $Name $What
}
function Get-CwDeptSpName {
    param([string]$CwDeptName)
    if ($script:CwDeptToSp.ContainsKey($CwDeptName)) { return $script:CwDeptToSp[$CwDeptName] }
    return $CwDeptName
}

# ---------------------------------------------------------------- template logic
function Build-TemplateDesired {
    # Returns scalar fields and reference fields (name -> {id}) from non-blank SharePoint columns. Blank = leave CW value unchanged.
    param($Item, [string]$Dept)
    $scalar = [ordered]@{}
    $refs = [ordered]@{}
    $title = ConvertTo-Text $Item.FieldValues['Title']
    if (-not $title) { Block 'Title is required.' }
    $scalar['name'] = $title
    $map = @{ Summary = 'summary'; WorkInstructions = 'problem'; InternalAnalysis = 'internalAnalysis'; Resolution = 'resolution' }
    foreach ($k in $map.Keys) { $t = ConvertTo-Text (Get-Field $Item $k); if ($t) { $scalar[$map[$k]] = $t } }
    foreach ($pair in @(@('HoursBudget', 'hoursBudget', 'dec'), @('ScheduleDaysBefore', 'scheduleDaysBefore', 'int'), @('ServiceDaysBefore', 'serviceDaysBefore', 'int'))) {
        $raw = Get-Field $Item $pair[0]
        $t = ConvertTo-Text $raw
        if ($t -eq '') { continue }
        if ($pair[2] -eq 'int') { $scalar[$pair[1]] = ConvertTo-IntOrNull $raw }
        else {
            $d = 0.0
            if (-not [double]::TryParse($t, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { Block "HoursBudget '$t' is not a number." }
            $scalar[$pair[1]] = $d
        }
    }
    $attach = ConvertTo-BoolOrNull (Get-Field $Item 'AttachSchedule')
    if ($null -ne $attach) { $scalar['attachScheduleToNewServiceFlag'] = $attach }

    $boardName = ConvertTo-Text (Get-Field $Item 'CWBoard')
    $board = $null
    if ($boardName) {
        $board = Resolve-Board $boardName
        $refs['board'] = $board.id
        foreach ($c in @(@('CWStatus', 'statuses', 'status'), @('CWType', 'types', 'type'), @('CWSubtype', 'subtypes', 'subType'), @('CWItem', 'items', 'item'), @('CWTeam', 'teams', 'team'))) {
            $n = ConvertTo-Text (Get-Field $Item $c[0])
            if ($n) { $refs[$c[2]] = (Resolve-BoardChild $board.id $c[1] $n).id }
        }
    }
    $pr = ConvertTo-Text (Get-Field $Item 'CWPriority')
    if ($pr) { $refs['priority'] = (Resolve-Simple 'service/priorities' $pr 'Priority').id }
    $cwDept = ConvertTo-Text (Get-Field $Item 'CWDepartment')
    if ($cwDept) {
        $d = Resolve-Simple 'system/departments' $cwDept 'Department'
        if ((Get-CwDeptSpName $d.name) -ne $Dept) { Block "CWDepartment '$cwDept' does not belong to SharePoint department '$Dept'." }
        $refs['department'] = $d.id
    }
    $agr = ConvertTo-IntOrNull (Get-Field $Item 'AgreementId')
    if ($null -ne $agr) { $refs['agreement'] = $agr }
    $co = ConvertTo-IntOrNull (Get-Field $Item 'CompanyId')
    if ($null -ne $co) { $refs['company'] = $co }
    return @{ Scalar = $scalar; Refs = $refs; BoardId = $(if ($board) { $board.id } else { $null }) }
}

function New-TemplateBody {
    param($Desired, [bool]$Master)
    $b = [ordered]@{}
    foreach ($k in $Desired.Scalar.Keys) { $b[$k] = $Desired.Scalar[$k] }
    if (-not $b.Contains('summary')) { $b['summary'] = $b['name'] }
    foreach ($k in $Desired.Refs.Keys) {
        if ($Master -and $k -in @('company', 'agreement')) { continue }
        $b[$k] = @{ id = $Desired.Refs[$k] }
    }
    $b['templateFlag'] = $Master
    return $b
}

function Get-TemplatePatchOps {
    param($Current, $Desired)
    $ops = [System.Collections.Generic.List[object]]::new()
    foreach ($k in $Desired.Scalar.Keys) {
        $cur = $Current.PSObject.Properties[$k]
        $curVal = if ($cur) { $cur.Value } else { $null }
        if ("$curVal" -ne "$($Desired.Scalar[$k])") { $ops.Add(@{ op = 'replace'; path = "/$k"; value = $Desired.Scalar[$k] }) }
    }
    foreach ($k in $Desired.Refs.Keys) {
        if ($k -eq 'company') {
            if ($Current.company -and [int]$Current.company.id -ne [int]$Desired.Refs[$k]) { Block "CompanyId differs from ConnectWise (company cannot be changed on an existing template)." }
            continue
        }
        $cur = $Current.PSObject.Properties[$k]
        $curId = if ($cur -and $cur.Value -and $cur.Value.PSObject.Properties['id']) { $cur.Value.id } else { $null }
        if ($k -eq 'agreement' -and $null -eq $curId) { $curId = $null }
        if ("$curId" -ne "$($Desired.Refs[$k])") { $ops.Add(@{ op = 'replace'; path = "/$k/id"; value = $Desired.Refs[$k] }) }
    }
    return , $ops.ToArray()
}

function Find-Master {
    param([string]$Name)
    $safe = $Name.Replace('"', '\"')
    $rows = @(Get-CWAll 'service/templates' ('conditions=' + [uri]::EscapeDataString("name=""$safe"" and templateFlag=true")))
    $rows = @($rows | Where-Object { $_.name -ieq $Name -and $_.templateFlag })
    if ($rows.Count -gt 1) { Block "More than one master template named '$Name' exists in ConnectWise; resolve manually." }
    if ($rows.Count -eq 1) { return $rows[0] }
    return $null
}
function Find-CompanyTemplates {
    param([string]$Name, [int]$CompanyId)
    $safe = $Name.Replace('"', '\"')
    $rows = @(Get-CWAll 'service/templates' ('conditions=' + [uri]::EscapeDataString("name=""$safe"" and templateFlag=false")))
    return @($rows | Where-Object { $_.name -ieq $Name -and $_.company -and [int]$_.company.id -eq $CompanyId })
}

# ---------------------------------------------------------------- schedule logic
function Build-ScheduleDesired {
    param($Item, [string]$Site, [string]$Dept, [string]$TemplateList)
    $d = @{}
    $parentId = ConvertTo-IntOrNull (Get-Field $Item 'ParentTemplateId')
    if ($null -eq $parentId) {
        $pn = ConvertTo-Text (Get-Field $Item 'ParentTemplateName')
        if (-not $pn) { Block 'ParentTemplateId or ParentTemplateName is required.' }
        $hits = @(Find-SpItemsByField $Site $TemplateList 'Title' $pn)
        if ($hits.Count -ne 1) { Block "Parent template '$pn' matched $($hits.Count) SharePoint items; set ParentTemplateId." }
        $parentId = ConvertTo-IntOrNull (Get-Field $hits[0] 'CWTemplateId')
        if ($null -eq $parentId) { Block "Parent template '$pn' has not been synced to ConnectWise yet (no CWTemplateId)." }
    }
    $d.ParentId = $parentId
    $start = Get-Field $Item 'DateStart'
    if (-not $start) { Block 'DateStart is required.' }
    $d.Start = ([datetime]$start).ToUniversalTime()
    # DateEnd is not pushed: ConnectWise derives each appointment's endDateTime from start + hours.
    $d.End = $null
    $h = ConvertTo-Text (Get-Field $Item 'Hours')
    $d.Hours = $null
    if ($h) { $x = 0.0; if (-not [double]::TryParse($h, [Globalization.CultureInfo]::InvariantCulture, [ref]$x) -or $x -lt 0) { Block "Hours '$h' is invalid." }; $d.Hours = $x }

    $memberId = ConvertTo-IntOrNull (Get-Field $Item 'MemberId')
    $mi = ConvertTo-Text (Get-Field $Item 'MemberIdentifier')
    if ($null -eq $memberId -and $mi) {
        if ($mi -notmatch '^[A-Za-z0-9._-]{1,50}$') { Block "MemberIdentifier '$mi' has invalid characters." }
        $m = @(Invoke-CWRequest -Method GET -Path ("system/members?conditions=" + [uri]::EscapeDataString("identifier=""$mi""")))
        if ($m.Count -ne 1) { Block "Member '$mi' matched $($m.Count) ConnectWise members." }
        $memberId = [int]$m[0].id
    }
    $d.MemberId = $memberId
    $d.IsOwner = ConvertTo-BoolOrNull (Get-Field $Item 'TicketOwner')

    $statusName = ConvertTo-Text (Get-Field $Item 'ScheduleStatus')
    if (-not $statusName) { $statusName = 'Firm' }
    $d.StatusId = (Resolve-Simple 'schedule/statuses' $statusName 'Schedule status').id

    $rt = (ConvertTo-Text (Get-Field $Item 'RecurType')).ToLowerInvariant()
    if ($rt -notin @('d', 'w', 'm', 'y')) { Block "RecurType must be d, w, m or y (got '$rt')." }
    $iv = ConvertTo-IntOrNull (Get-Field $Item 'RecurInterval')
    if ($null -eq $iv) { $iv = 1 }
    if ($iv -lt 1) { Block 'RecurInterval must be 1 or more.' }
    $dow = ConvertTo-Text (Get-Field $Item 'DaysOfWeekRaw')
    $abs = ConvertTo-IntOrNull (Get-Field $Item 'AbsDayNbr')
    $mon = ConvertTo-IntOrNull (Get-Field $Item 'AbsMonthNbr')
    $sel = ConvertTo-IntOrNull (Get-Field $Item 'SelectInterval')
    switch ($rt) {
        'w' { if (-not $dow) { Block 'Weekly recurrence needs DaysOfWeekRaw.' } }
        'm' {
            if ($null -ne $abs) { if ($abs -lt 1 -or $abs -gt 31) { Block 'AbsDayNbr must be 1-31.' } }
            elseif ($null -eq $sel -or -not $dow) { Block 'Monthly recurrence needs AbsDayNbr, or SelectInterval with DaysOfWeekRaw.' }
        }
        'y' {
            if ($null -eq $mon -or $mon -lt 1 -or $mon -gt 12) { Block 'Yearly recurrence needs AbsMonthNbr 1-12.' }
            if ($null -eq $abs -or $abs -lt 1 -or $abs -gt 31) { Block 'Yearly recurrence needs AbsDayNbr 1-31.' }
        }
    }
    $rec = [ordered]@{ RecurType = $rt; RecurInterval = $iv }
    if ($dow) { $rec['DaysOfWeek'] = $dow }
    if ($null -ne $abs) { $rec['AbsDayNbr'] = $abs }
    if ($null -ne $mon) { $rec['AbsMonthNbr'] = $mon }
    if ($null -ne $sel) { $rec['SelectInterval'] = $sel }
    $d.Recurrence = $rec
    return $d
}

function New-ScheduleBody {
    param($D)
    $b = [ordered]@{
        serviceTemplate = @{ id = $D.ParentId }
        startDateTime   = $D.Start.ToString('yyyy-MM-ddTHH:mm:ssZ')
        scheduleStatus  = @{ id = $D.StatusId }
        recurrence      = $D.Recurrence
    }
    if ($null -ne $D.Hours) { $b['hours'] = $D.Hours }
    if ($null -ne $D.MemberId) { $b['member'] = @{ id = $D.MemberId } }
    if ($null -ne $D.IsOwner) { $b['IsTicketOwner'] = $D.IsOwner }
    return $b
}

# ---------------------------------------------------------------- write-back
function Get-CwDeptForTemplate {
    param($Template)
    if ($Template.PSObject.Properties['department'] -and $Template.department) { return (Get-CwDeptSpName $Template.department.name) }
    return $null
}
function Update-Registry {
    param([string]$Key, [string]$RecordType, $Cw, [string]$Dept, [string]$TargetList, [string]$TargetSite, [bool]$IsDryRun)
    $title = $Key
    if ($IsDryRun) { $script:Ops.Add("DRYRUN registry upsert $Key"); return }
    $cols = Get-SpFieldNames $script:RegistrySite $script:RegistryList
    $hits = @(Find-SpItemsByField $script:RegistrySite $script:RegistryList 'CWKey' $Key)
    $vals = @{
        Title = $title; CWKey = $Key; RecordType = $RecordType; OwnerDepartment = $Dept
        TargetSite = $TargetSite; TargetList = $TargetList; SyncStatus = 'Routed'
        LastSyncUtc = (Get-Date).ToUniversalTime().ToString('o')
    }
    if ($RecordType -eq 'Template') { $vals.CWTemplateId = [string]$Cw.id } else { $vals.CWResourceId = [string]$Cw.id; $vals.CWTemplateId = [string]$Cw.serviceTemplate.id }
    if ($hits.Count -gt 1) { $script:Warnings.Add("Registry has $($hits.Count) rows for $Key; not modified."); return }
    if ($hits.Count -eq 1) { Set-SpItem $script:RegistrySite $script:RegistryList $hits[0].Id $vals $cols; $script:Ops.Add("Registry updated $Key") }
    else {
        $v = @{}; foreach ($k in $vals.Keys) { if ($k -in $cols) { $v[$k] = $vals[$k] } }
        $null = Add-PnPListItem -List $script:RegistryList -Values $v -Connection (Get-SpConnection $script:RegistrySite)
        $script:Ops.Add("Registry created $Key")
    }
}

# ---------------------------------------------------------------- main
$statusCode = 200
$result = [ordered]@{ correlationId = $script:CorrelationId; outcome = 'Failed' }
$item = $null; $site = $null; $listName = $null; $itemId = 0; $cols = @(); $isDryRun = $true
$controlCols = @('PushHash', 'PushStatus', 'PushMessage', 'PushUtc')
$writeBackOk = $false

try {
    $body = $Request.Body
    if ($body -is [string]) { $body = $body | ConvertFrom-Json -AsHashtable }
    if ($body -isnot [System.Collections.IDictionary]) { throw (New-RequestException 400 'Request body must be a JSON object.') }
    $event = "$($body['event'])".ToLowerInvariant()
    if ($event -notin @('created', 'modified', 'deleted')) { throw (New-RequestException 400 "event must be created, modified or deleted.") }
    $site = "$($body['siteUrl'])".TrimEnd('/')
    $listName = "$($body['listName'])"
    $itemId = 0
    if (-not [int]::TryParse("$($body['itemId'])", [ref]$itemId) -or $itemId -lt 1) { throw (New-RequestException 400 'itemId must be a positive integer.') }

    $isDryRun = (Test-SettingTrue 'SpTemplateSync_DryRun' 'true') -or ($body.ContainsKey('dryRun') -and $body['dryRun'] -ne $false)
    $result.event = $event; $result.itemId = $itemId; $result.dryRun = $isDryRun

    # Allowlist: site -> department, list name -> list type
    $dept = $null
    foreach ($k in $script:Departments.Keys) { if ($script:Departments[$k].TrimEnd('/') -ieq $site) { $dept = $k } }
    if (-not $dept) { throw (New-RequestException 403 'siteUrl is not an approved department site.') }
    $site = $script:Departments[$dept]
    $templateLists = @("CW- $dept - Templates", "CW - $dept - Templates")
    $scheduleList = "CW - $dept - Schedules"
    $listType = $null
    if ($listName -in $templateLists) { $listType = 'Template' }
    elseif ($listName -ieq $scheduleList) { $listType = 'Schedule' }
    else { throw (New-RequestException 403 'listName is not an approved list for this department.') }
    $result.department = $dept; $result.listType = $listType
    $tplListName = $templateLists[0]
    if ($listType -eq 'Schedule') {
        foreach ($n in $templateLists) { if (Get-PnPList -Identity $n -Connection (Get-SpConnection $site) -ErrorAction SilentlyContinue) { $tplListName = $n; break } }
    } else { $tplListName = $listName }

    $createMaster = Test-SettingTrue 'SpTemplateSync_CreateMasterIfMissing' 'true'

    # ---- delete: ConnectWise's API cannot end a recurrence (no end field; resource DELETE returns 405), so nothing is changed
    if ($event -eq 'deleted') {
        $result.outcome = 'NoChange'
        $script:Warnings.Add('Deleted items are not propagated: the ConnectWise API cannot end a recurring schedule or deactivate a template. Remove or retire the record in ConnectWise manually.')
    }
    else {
        $item = Get-SpItem $site $listName $itemId
        $cols = Get-SpFieldNames $site $listName
        foreach ($c in $controlCols) {
            if ($c -notin $cols -and -not $isDryRun) { Add-SpTextField $site $listName $c ($(if ($c -eq 'PushMessage') { 'Note' } else { 'Text' })); $cols += $c; $script:Ops.Add("Created column $c") }
        }
        $writeBackOk = ($controlCols | Where-Object { $_ -in $cols }).Count -eq $controlCols.Count

        # input hash covers only inputs, never write-back columns, so our own updates are recognised as no-ops
        $inputCols = if ($listType -eq 'Template') {
            @('Title', 'Summary', 'WorkInstructions', 'InternalAnalysis', 'Resolution', 'HoursBudget', 'ScheduleDaysBefore', 'ServiceDaysBefore', 'AttachSchedule', 'CWBoard', 'CWStatus', 'CWType', 'CWSubtype', 'CWItem', 'CWTeam', 'CWPriority', 'CWDepartment', 'CompanyId', 'AgreementId')
        } else {
            @('ParentTemplateId', 'ParentTemplateName', 'DateStart', 'Hours', 'MemberId', 'MemberIdentifier', 'TicketOwner', 'ScheduleStatus', 'RecurType', 'RecurInterval', 'DaysOfWeekRaw', 'AbsDayNbr', 'AbsMonthNbr', 'SelectInterval')
        }
        $hashText = ($inputCols | ForEach-Object { "$_=" + (ConvertTo-Text (Get-Field $item $_)) }) -join '|'
        $hash = Get-Sha256 $hashText
        $prevHash = ConvertTo-Text (Get-Field $item 'PushHash')
        $pushStatus = ConvertTo-Text (Get-Field $item 'PushStatus')
        $existingId = ConvertTo-IntOrNull (Get-Field $item $(if ($listType -eq 'Template') { 'CWTemplateId' } else { 'CWResourceId' }))

        if ($prevHash -eq $hash -and $pushStatus -eq 'Synced' -and $null -ne $existingId) {
            $result.outcome = 'NoChange'; $result.cwId = $existingId
            $script:Ops.Add('Inputs unchanged since last push; skipped.')
        }
        else {
            if ($writeBackOk -and -not $isDryRun) { Set-SpItem $site $listName $itemId @{ PushStatus = 'Creating'; PushMessage = ''; PushUtc = (Get-Date).ToUniversalTime().ToString('o') } $cols }

            if ($listType -eq 'Template') {
                $desired = Build-TemplateDesired $item $dept
                $companyId = $desired.Refs['company']
                if ($null -eq $companyId) { Block 'CompanyId is required.' }
                $co = Invoke-CWRequest -Method GET -Path "company/companies/$companyId" -AllowNotFound
                if (-not $co) { Block "ConnectWise company $companyId was not found." }
                if ($desired.Refs.Contains('agreement')) {
                    $ag = Invoke-CWRequest -Method GET -Path "finance/agreements/$($desired.Refs['agreement'])" -AllowNotFound
                    if (-not $ag) { Block "Agreement $($desired.Refs['agreement']) was not found." }
                    if ([int]$ag.company.id -ne [int]$companyId) { Block 'Agreement belongs to a different company.' }
                }
                $name = $desired.Scalar['name']
                $cwTpl = $null
                if ($null -ne $existingId) {
                    $cwTpl = Invoke-CWRequest -Method GET -Path "service/templates/$existingId" -AllowNotFound
                    if (-not $cwTpl) { Block "CWTemplateId $existingId no longer exists in ConnectWise." }
                    if ($cwTpl.templateFlag) { Block 'CWTemplateId points at a master template; refusing to modify.' }
                    if ((Get-CwDeptForTemplate $cwTpl) -ne $dept -and (Get-CwDeptForTemplate $cwTpl)) { Block 'ConnectWise template belongs to a different department.' }
                }
                else {
                    $dups = @(Find-CompanyTemplates $name ([int]$companyId))
                    if ($dups.Count -gt 1) { Block "Multiple ConnectWise templates named '$name' exist for this company." }
                    if ($dups.Count -eq 1) {
                        if ($pushStatus -notin @('Creating', 'AwaitingCompanyApply')) { Block "A ConnectWise template named '$name' already exists for this company (id $($dups[0].id)). Set CWTemplateId to link it." }
                        $cwTpl = $dups[0]; $result.outcome = 'Adopted'
                    }
                }

                if ($null -eq $cwTpl) {
                    $master = Find-Master $name
                    $result.masterCreated = $false
                    if (-not $master) {
                        if (-not $createMaster) { Block "No master template named '$name' exists and master creation is disabled." }
                        $mBody = New-TemplateBody $desired $true
                        if ($isDryRun) { $script:Ops.Add("DRYRUN POST service/templates (master): $($mBody | ConvertTo-Json -Depth 10 -Compress)") }
                        else { $master = Invoke-CWRequest -Method POST -Path 'service/templates' -Body $mBody; $script:Ops.Add("Created master template $($master.id)"); $result.masterCreated = $true }
                    }
                    if ($master) { $result.masterTemplateId = $master.id }
                    # The ConnectWise REST API cannot create company-level templates (POST always yields a master and ignores
                    # company/templateFlag; PATCH ignores both), so stop here and wait for the master to be applied to the company in ConnectWise.
                    $result.outcome = if ($isDryRun) { 'DryRun' } else { 'AwaitingCompanyApply' }
                    $msg = "Master template $(if ($master) { $master.id } else { '(to be created)' }) is ready. Apply it to company $companyId in ConnectWise, then re-save this item to link the company template."
                    $script:Warnings.Add($msg)
                    if (-not $isDryRun) {
                        $wb = @{ PushStatus = 'AwaitingCompanyApply'; PushMessage = $msg; PushUtc = (Get-Date).ToUniversalTime().ToString('o') }
                        if ($master) { $wb.MasterTemplateId = [string]$master.id }
                        Set-SpItem $site $listName $itemId $wb $cols
                    }
                }
                elseif ($result.outcome -ne 'Adopted') {
                    $ops = @(Get-TemplatePatchOps $cwTpl $desired)
                    if ($ops.Count -eq 0) { $result.outcome = 'NoChange' }
                    elseif ($isDryRun) { $script:Ops.Add("DRYRUN PATCH service/templates/$($cwTpl.id): $($ops | ConvertTo-Json -Depth 6 -Compress)"); $result.outcome = 'DryRun' }
                    else { $null = Invoke-CWRequest -Method PATCH -Path "service/templates/$($cwTpl.id)" -Body $ops; $script:Ops.Add("Patched template $($cwTpl.id) ($($ops.Count) change(s))"); $result.outcome = 'Updated' }
                }

                if ($cwTpl -and -not $isDryRun) {
                    $result.cwId = $cwTpl.id
                    # write the ID back first so a later failure cannot cause a duplicate create on retry
                    Set-SpItem $site $listName $itemId @{ CWTemplateId = [string]$cwTpl.id } $cols
                    $fresh = Invoke-CWRequest -Method GET -Path "service/templates/$($cwTpl.id)"
                    $wb = @{
                        CWTemplateId = [string]$fresh.id; Fingerprint = (Get-Sha256 ($fresh | ConvertTo-Json -Depth 20 -Compress))
                        PushHash = $hash; PushStatus = 'Synced'; PushMessage = ''; PushUtc = (Get-Date).ToUniversalTime().ToString('o')
                    }
                    if ($fresh.PSObject.Properties['_info'] -and $fresh._info) {
                        if ($fresh._info.PSObject.Properties['lastUpdated']) { $wb.CWLastUpdated = $fresh._info.lastUpdated }
                        if ($fresh._info.PSObject.Properties['updatedBy']) { $wb.CWUpdatedBy = $fresh._info.updatedBy }
                    }
                    if ($result.Contains('masterTemplateId') -and $result.masterTemplateId) { $wb.MasterTemplateId = [string]$result.masterTemplateId }
                    Set-SpItem $site $listName $itemId $wb $cols
                    Update-Registry "T:$($fresh.id)" 'Template' $fresh $dept $listName $site $false
                } elseif ($cwTpl) { $result.cwId = $cwTpl.id }
            }
            else {
                $d = Build-ScheduleDesired $item $site $dept $tplListName
                $parent = Invoke-CWRequest -Method GET -Path "service/templates/$($d.ParentId)" -AllowNotFound
                if (-not $parent) { Block "Parent template $($d.ParentId) was not found in ConnectWise." }
                if ($parent.templateFlag) { Block 'Schedules can only be attached to company templates, not masters.' }
                $pd = Get-CwDeptForTemplate $parent
                if ($pd -and $pd -ne $dept) { Block 'Parent template belongs to a different department.' }
                $body2 = New-ScheduleBody $d
                $cwRes = $null
                $existing = @(Get-CWAll "service/templates/$($d.ParentId)/resources")
                if ($null -ne $existingId) {
                    $cwRes = @($existing | Where-Object { [int]$_.id -eq $existingId })[0]
                    if (-not $cwRes) { Block "CWResourceId $existingId was not found under template $($d.ParentId)." }
                } elseif ($pushStatus -eq 'Creating') {
                    $m = @($existing | Where-Object {
                            $mid = if ($_.PSObject.Properties['member'] -and $_.member) { [int]$_.member.id } else { $null }
                            $mid -eq $d.MemberId -and ([datetime]$_.startDateTime).ToUniversalTime() -eq $d.Start -and $_.recurrence.RecurType -eq $d.Recurrence.RecurType })
                    if ($m.Count -eq 1) { $cwRes = $m[0]; $result.outcome = 'Adopted' }
                }
                if ($null -eq $cwRes) {
                    if ($isDryRun) { $script:Ops.Add("DRYRUN POST service/templates/$($d.ParentId)/resources: $($body2 | ConvertTo-Json -Depth 10 -Compress)"); $result.outcome = 'DryRun' }
                    else { $cwRes = Invoke-CWRequest -Method POST -Path "service/templates/$($d.ParentId)/resources" -Body $body2; $script:Ops.Add("Created resource $($cwRes.id)"); $result.outcome = 'Created' }
                }
                elseif ($result.outcome -ne 'Adopted') {
                    # merge-then-PUT: overlay desired values on the current object
                    $merged = [ordered]@{}
                    foreach ($p in $cwRes.PSObject.Properties) { if ($p.Name -ne '_info') { $merged[$p.Name] = $p.Value } }
                    foreach ($k in $body2.Keys) { $merged[$k] = $body2[$k] }
                    if ($isDryRun) { $script:Ops.Add("DRYRUN PUT service/templates/$($d.ParentId)/resources/$($cwRes.id)"); $result.outcome = 'DryRun' }
                    else { $null = Invoke-CWRequest -Method PUT -Path "service/templates/$($d.ParentId)/resources/$($cwRes.id)" -Body $merged; $script:Ops.Add("Updated resource $($cwRes.id)"); $result.outcome = 'Updated' }
                }
                if ($cwRes -and -not $isDryRun) {
                    $result.cwId = $cwRes.id
                    Set-SpItem $site $listName $itemId @{ CWResourceId = [string]$cwRes.id } $cols
                    $fresh = Invoke-CWRequest -Method GET -Path "service/templates/$($d.ParentId)/resources/$($cwRes.id)"
                    $wb = @{
                        CWResourceId = [string]$fresh.id; ParentTemplateId = [string]$d.ParentId; Fingerprint = (Get-Sha256 ($fresh | ConvertTo-Json -Depth 20 -Compress))
                        PushHash = $hash; PushStatus = 'Synced'; PushMessage = ''; PushUtc = (Get-Date).ToUniversalTime().ToString('o')
                    }
                    if ($fresh.PSObject.Properties['recurrence'] -and $fresh.recurrence -and $fresh.recurrence.PSObject.Properties['id']) { $wb.RecurrenceId = [string]$fresh.recurrence.id }
                    Set-SpItem $site $listName $itemId $wb $cols
                    Update-Registry "R:$($fresh.id)" 'Schedule' $fresh $dept $listName $site $false
                } elseif ($cwRes) { $result.cwId = $cwRes.id }
            }
        }
    }
}
catch {
    $ex = $_.Exception
    $statusCode = 500
    if ($ex.Data.Contains('HttpStatusCode')) { $statusCode = [int]$ex.Data['HttpStatusCode'] }
    $result.outcome = if ($statusCode -eq 422) { 'Blocked' } else { 'Failed' }
    $result.error = if ($statusCode -ge 500 -and $statusCode -ne 502) { 'Internal error. See function logs with the correlation ID.' } else { $ex.Message }
    Write-Error "[$($script:CorrelationId)] $($ex.Message)" -ErrorAction Continue
    if ($item -and $writeBackOk -and -not $isDryRun) {
        try {
            $msg = $ex.Message; if ($msg.Length -gt 1000) { $msg = $msg.Substring(0, 1000) }
            Set-SpItem $site $listName $itemId @{ PushStatus = $(if ($statusCode -eq 422) { 'Blocked' } else { 'Failed' }); PushMessage = $msg; PushUtc = (Get-Date).ToUniversalTime().ToString('o') } $cols
        } catch { Write-Warning "[$($script:CorrelationId)] Could not record failure on the SharePoint item." }
    }
}
finally {
    foreach ($c in $script:SpConnections.Values) { try { Disconnect-PnPOnline -Connection $c } catch { } }
}

$result.operations = @($script:Ops)
$result.warnings = @($script:Warnings)
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]$statusCode
        Headers    = @{ 'Content-Type' = 'application/json' }
        Body       = ($result | ConvertTo-Json -Depth 8)
    })
