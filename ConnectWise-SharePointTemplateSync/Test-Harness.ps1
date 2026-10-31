#requires -Version 7.2
# Mock harness for run.ps1. No network calls: Invoke-RestMethod and the PnP cmdlets are replaced by stubs.
# Run: pwsh -File .\Test-Harness.ps1

$ErrorActionPreference = 'Stop'
$runPath = Join-Path $PSScriptRoot 'run.ps1'
Add-Type -TypeDefinition 'public class HttpResponseContext { public System.Net.HttpStatusCode StatusCode; public System.Collections.IDictionary Headers; public string Body; }' -ErrorAction SilentlyContinue

foreach ($kv in @{ ConnectWisePsa_ApiCompanyId = 'co'; ConnectWisePsa_ApiPublicKey = 'pub'; ConnectWisePsa_ApiPrivateKey = 'priv'; ConnectWisePsa_ApiClientId = 'cid'
        Ms365_TenantId = 't'; Ms365_AuthAppId = 'a'; Ms365_CertBase64 = 'c'; Ms365_CertPassword = 'p' }.GetEnumerator()) { [Environment]::SetEnvironmentVariable($kv.Key, $kv.Value) }

# ---- stubs
function Invoke-RestMethod {
    param($Method, $Uri, $Headers, $ContentType, $Body, $ErrorAction)
    $path = ($Uri -replace '^.*/apis/3\.0/', '')
    $global:CwCalls.Add([pscustomobject]@{ Method = $Method; Path = $path; Body = $Body })
    $b = if ($Body) { $Body | ConvertFrom-Json -AsHashtable -Depth 20 } else { $null }
    $dec = [uri]::UnescapeDataString($path)
    switch -Regex ("$Method $dec") {
        '^GET service/boards\?' { return @([pscustomobject]@{ id = 1; name = 'Managed Services Board' }) }
        '^GET service/boards/1/\w+\?' { return @([pscustomobject]@{ id = 5; name = 'New' }) }
        '^GET service/priorities' { return @([pscustomobject]@{ id = 3; name = 'Priority 3 - Normal' }) }
        '^GET system/departments' { return @([pscustomobject]@{ id = 9; name = 'Service Desk' }) }
        '^GET schedule/statuses' { return @([pscustomobject]@{ id = 2; name = 'Firm' }) }
        '^GET company/companies/100' { return [pscustomobject]@{ id = 100 } }
        '^GET service/templates\?.*templateFlag=true' { return @($global:Masters) }
        '^GET service/templates\?.*templateFlag=false' { return @($global:CoTpls) }
        '^GET service/templates/(\d+)/resources\?' { return @($global:Resources[[int]$Matches[1]]) }
        '^GET service/templates/(\d+)$' { return $global:Tpls[[int]$Matches[1]] }
        '^POST service/templates$' {
            $id = $global:NextId++
            $o = [pscustomobject]@{ id = $id; name = $b.name; templateFlag = $b.templateFlag; department = [pscustomobject]@{ id = 9; name = 'Service Desk' }
                _info = [pscustomobject]@{ lastUpdated = '2026-10-30T00:00:00Z'; updatedBy = 'fn' } }
            $global:Tpls[$id] = $o; return $o
        }
        '^PATCH ' { return [pscustomobject]@{ ok = $true } }
        '^POST service/templates/(\d+)/resources$' {
            $o = [pscustomobject]@{ id = 7000; serviceTemplate = [pscustomobject]@{ id = [int]$Matches[1] }; startDateTime = $b.startDateTime; hours = $b.hours; recurrence = [pscustomobject]@{ id = 8000; RecurType = $b.recurrence.RecurType }; _info = [pscustomobject]@{ lastUpdated = '2026-10-30T00:00:00Z'; updatedBy = 'fn' } }
            $global:ResById[7000] = $o; return $o
        }
        '^GET service/templates/(\d+)/resources/(\d+)$' { return $global:ResById[[int]$Matches[2]] }
        default { throw "Unmocked CW call: $Method $dec" }
    }
}
function Connect-PnPOnline { 'conn' }
function Disconnect-PnPOnline { }
function Get-PnPList { param($Identity, $Connection, $ErrorAction) return [pscustomobject]@{ Title = $Identity } }

# Test hook honoured by Send-CWHttp in run.ps1: route all ConnectWise HTTP through the stub above.
$global:CwHttpOverride = {
    param($Method, $Uri, $BodyJson)
    $r = Invoke-RestMethod -Method $Method -Uri $Uri -Body $BodyJson
    if ($null -eq $r) {
        if ($Uri -match '/\d+$') { return [pscustomobject]@{ Status = 404; Content = '' } }
        return [pscustomobject]@{ Status = 200; Content = '[]' }
    }
    return [pscustomobject]@{ Status = $(if ($Method -eq 'POST') { 201 } else { 200 }); Content = (ConvertTo-Json -InputObject $r -Depth 20 -Compress) }
}
function Get-PnPField { param($List, $Connection) return @($global:Cols | ForEach-Object { [pscustomobject]@{ InternalName = $_ } }) }
function Add-PnPField { param($List, $InternalName, $DisplayName, $Type, $AddToDefaultView, $Connection) $global:Cols += $InternalName; $global:SpWrites.Add("AddField:$InternalName") }
function Get-PnPListItem {
    param($List, $Id, $Query, $Connection)
    if ($List -eq 'CW - Recurring Work Sync Registry') { return @($global:Registry) }
    if ($PSBoundParameters.ContainsKey('Id')) { return $global:Items[$Id] }
    return @()
}
function Set-PnPListItem {
    param($List, $Identity, $Values, $UpdateType, $Connection)
    foreach ($k in $Values.Keys) { $global:Items[$Identity].FieldValues[$k] = $Values[$k] }
    $global:SpWrites.Add("Set:${Identity}:" + (($Values.Keys | Sort-Object) -join ','))
}
function Add-PnPListItem { param($List, $Values, $Connection) $global:SpWrites.Add("AddItem:$($Values.CWKey)") }
function Push-OutputBinding { param($Name, $Value) $global:Out = $Value }

function Reset-World {
    $global:CwCalls = [System.Collections.Generic.List[object]]::new()
    $global:SpWrites = [System.Collections.Generic.List[string]]::new()
    $global:Masters = @(); $global:CoTpls = @(); $global:Registry = @(); $global:Resources = @{}; $global:ResById = @{}; $global:Tpls = @{}; $global:NextId = 5000
    $global:Cols = @('Title', 'CWTemplateId', 'CWBoard', 'CWStatus', 'CompanyId', 'CWDepartment', 'Summary', 'Fingerprint', 'CWLastUpdated', 'CWUpdatedBy', 'CWResourceId', 'ParentTemplateId', 'RecurType', 'RecurInterval', 'DaysOfWeekRaw', 'DateStart',     'RecurrenceId', 'MasterTemplateId', 'CWKey', 'RecordType', 'OwnerDepartment', 'TargetSite', 'TargetList', 'SyncStatus', 'LastSyncUtc')
    $global:Items = @{ 7 = [pscustomobject]@{ Id = 7; FieldValues = @{ Title = 'Weekly Patch Review'; CWBoard = 'Managed Services Board'; CWStatus = 'New'; CompanyId = '100'; CWDepartment = 'Service Desk'; Summary = 'Review patches' } } }
}
function Invoke-Fn {
    param([hashtable]$Body, [bool]$Dry)
    [Environment]::SetEnvironmentVariable('SpTemplateSync_DryRun', $(if ($Dry) { 'true' } else { 'false' }))
    $global:Out = $null
    $req = @{ Body = $Body }
    . $runPath $req @{} 2>&1 | ForEach-Object { if ($env:HARNESS_VERBOSE) { Write-Host "  [log] $_" -ForegroundColor DarkGray } }
    return [pscustomobject]@{ Status = [int]$global:Out.StatusCode; Json = ($global:Out.Body | ConvertFrom-Json) }
}

$script:Fail = 0
function Assert { param([string]$Name, [bool]$Cond) if ($Cond) { Write-Host "PASS  $Name" -ForegroundColor Green } else { Write-Host "FAIL  $Name" -ForegroundColor Red; $script:Fail++ } }

$site = 'https://palaisparc.sharepoint.com/Service/servicedesk'
$tplBody = @{ event = 'created'; siteUrl = $site; listName = 'CW- Service Desk - Templates'; itemId = 7 }

# 1 dry run: no writes anywhere
Reset-World
$r = Invoke-Fn $tplBody $true
Assert 'dry run: outcome DryRun, HTTP 200' ($r.Status -eq 200 -and $r.Json.outcome -eq 'DryRun')
Assert 'dry run: no CW writes' (@($global:CwCalls | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
Assert 'dry run: no SharePoint writes' ($global:SpWrites.Count -eq 0)
Assert 'dry run: plans only the master POST' (@($r.Json.operations | Where-Object { $_ -like 'DRYRUN POST*' }).Count -eq 1 -and $r.Json.warnings.Count -eq 1)

# 2 live create with missing master: master is created, company template is NOT (API cannot), item parked
Reset-World
$r = Invoke-Fn $tplBody $false
Assert 'create: outcome AwaitingCompanyApply, HTTP 200' ($r.Status -eq 200 -and $r.Json.outcome -eq 'AwaitingCompanyApply')
$posts = @($global:CwCalls | Where-Object { $_.Method -eq 'POST' })
Assert 'create: exactly one POST, a master without company' ($posts.Count -eq 1 -and ($posts[0].Body | ConvertFrom-Json).templateFlag -eq $true -and $null -eq ($posts[0].Body | ConvertFrom-Json).company)
Assert 'create: master body has board and department refs' (($posts[0].Body | ConvertFrom-Json).board.id -eq 1 -and ($posts[0].Body | ConvertFrom-Json).department.id -eq 9)
Assert 'create: item parked with MasterTemplateId' ($global:Items[7].FieldValues['PushStatus'] -eq 'AwaitingCompanyApply' -and $global:Items[7].FieldValues['MasterTemplateId'] -eq '5000')
Assert 'create: no registry row or CWTemplateId yet' (-not ($global:SpWrites | Where-Object { $_ -like 'AddItem:*' }) -and -not $global:Items[7].FieldValues.ContainsKey('CWTemplateId'))
Assert 'create: SyncStatus never touched' (-not ($global:SpWrites | Where-Object { $_ -match 'SyncStatus' }))

# 2b the admin applied the master to the company; re-save links it (adopt) and writes back
$global:CoTpls = @([pscustomobject]@{ id = 321; name = 'Weekly Patch Review'; templateFlag = $false; company = [pscustomobject]@{ id = 100 } })
$global:Tpls[321] = [pscustomobject]@{ id = 321; name = 'Weekly Patch Review'; templateFlag = $false; company = [pscustomobject]@{ id = 100 }; department = [pscustomobject]@{ id = 9; name = 'Service Desk' }; _info = [pscustomobject]@{ lastUpdated = '2026-10-30T00:00:00Z'; updatedBy = 'x' } }
$r = Invoke-Fn @{ event = 'modified'; siteUrl = $site; listName = 'CW- Service Desk - Templates'; itemId = 7 } $false
Assert 'adopt: company template linked after manual apply' ($r.Json.outcome -eq 'Adopted' -and $global:Items[7].FieldValues['CWTemplateId'] -eq '321' -and $global:Items[7].FieldValues['PushStatus'] -eq 'Synced')
Assert 'adopt: registry row created, no POST' ($global:SpWrites -contains 'AddItem:T:321' -and @($global:CwCalls | Where-Object { $_.Method -eq 'POST' }).Count -eq 1)
$before = $global:CwCalls.Count
$r2 = Invoke-Fn @{ event = 'modified'; siteUrl = $site; listName = 'CW- Service Desk - Templates'; itemId = 7 } $false
Assert 'echo guard: unchanged inputs -> NoChange, no CW calls' ($r2.Json.outcome -eq 'NoChange' -and $global:CwCalls.Count -eq $before)

# 3 modify -> PATCH only the differing field, as a JSON array even for one operation
$global:Items[7].FieldValues['Summary'] = 'Review patches v2'
foreach ($p in @(@('board', 1), @('status', 5), @('department', 9))) { $global:Tpls[321] | Add-Member -NotePropertyName $p[0] -NotePropertyValue ([pscustomobject]@{ id = $p[1] }) -Force }
$global:Tpls[321] | Add-Member -NotePropertyName summary -NotePropertyValue 'Review patches' -Force
$r3 = Invoke-Fn @{ event = 'modified'; siteUrl = $site; listName = 'CW- Service Desk - Templates'; itemId = 7 } $false
$patch = @($global:CwCalls | Where-Object { $_.Method -eq 'PATCH' })
Assert 'modify: single PATCH, JSON array, only summary' ($r3.Json.outcome -eq 'Updated' -and $patch.Count -eq 1 -and $patch[0].Body.StartsWith('[') -and @($patch[0].Body | ConvertFrom-Json).Count -eq 1 -and (@($patch[0].Body | ConvertFrom-Json))[0].path -eq '/summary')

# 4 existing same-name company template is not adopted without the awaiting state
Reset-World
$global:CoTpls = @([pscustomobject]@{ id = 321; name = 'Weekly Patch Review'; templateFlag = $false; company = [pscustomobject]@{ id = 100 } })
$r = Invoke-Fn $tplBody $false
Assert 'duplicate name: Blocked 422, no POST' ($r.Status -eq 422 -and $r.Json.outcome -eq 'Blocked' -and @($global:CwCalls | Where-Object { $_.Method -eq 'POST' }).Count -eq 0)
Assert 'duplicate name: failure recorded on item' ($global:Items[7].FieldValues['PushStatus'] -eq 'Blocked')

# 5 existing master reused, nothing created
Reset-World
$global:Masters = @([pscustomobject]@{ id = 88; name = 'Weekly Patch Review'; templateFlag = $true })
$r = Invoke-Fn $tplBody $false
Assert 'existing master: no POST, master id reported' (@($global:CwCalls | Where-Object { $_.Method -eq 'POST' }).Count -eq 0 -and $r.Json.masterTemplateId -eq 88 -and $r.Json.outcome -eq 'AwaitingCompanyApply')

# 6 request validation / allowlist
Reset-World
$r = Invoke-Fn @{ event = 'created'; siteUrl = 'https://evil.example.com'; listName = 'CW- Service Desk - Templates'; itemId = 7 } $true
Assert 'bad site: 403' ($r.Status -eq 403)
$r = Invoke-Fn @{ event = 'created'; siteUrl = $site; listName = 'Documents'; itemId = 7 } $true
Assert 'bad list: 403' ($r.Status -eq 403)
$r = Invoke-Fn @{ event = 'created'; siteUrl = $site; listName = 'CW- Service Desk - Templates'; itemId = 'x; drop' } $true
Assert 'bad itemId: 400' ($r.Status -eq 400)
$r = Invoke-Fn @{ event = 'nuke'; siteUrl = $site; listName = 'CW- Service Desk - Templates'; itemId = 7 } $true
Assert 'bad event: 400' ($r.Status -eq 400)

# 7 missing company blocks
Reset-World
$global:Items[7].FieldValues.Remove('CompanyId')
$r = Invoke-Fn $tplBody $true
Assert 'missing CompanyId: Blocked' ($r.Status -eq 422)

# 8 schedule validation
Reset-World
$global:Items[8] = [pscustomobject]@{ Id = 8; FieldValues = @{ ParentTemplateId = '5001'; DateStart = [datetime]'2026-11-01'; RecurType = 'w'; RecurInterval = '1' } }
$sb = @{ event = 'created'; siteUrl = $site; listName = 'CW - Service Desk - Schedules'; itemId = 8 }
$r = Invoke-Fn $sb $true
Assert 'schedule: weekly without days -> Blocked' ($r.Status -eq 422 -and $r.Json.error -like '*DaysOfWeekRaw*')
$global:Items[8].FieldValues.DaysOfWeekRaw = '2'
$global:Tpls[5001] = [pscustomobject]@{ id = 5001; name = 'Weekly Patch Review'; templateFlag = $false; department = [pscustomobject]@{ name = 'Service Desk' } }
$r = Invoke-Fn $sb $true
Assert 'schedule dry run: plans resource POST with recurrence' ($r.Json.outcome -eq 'DryRun' -and ($r.Json.operations -join ' ') -like '*"RecurType":"w"*' -and ($r.Json.operations -join ' ') -like '*"DaysOfWeek":"2"*')
Assert 'schedule dry run: no CW writes' (@($global:CwCalls | Where-Object { $_.Method -ne 'GET' }).Count -eq 0)
Assert 'schedule: DateEnd is never sent' (-not (($r.Json.operations -join ' ') -like '*endDateTime*'))

# 9 delete: never touches ConnectWise
Reset-World
$del = @{ event = 'deleted'; siteUrl = $site; listName = 'CW- Service Desk - Templates'; itemId = 7; cwId = 41 }
$r = Invoke-Fn $del $false
Assert 'delete: NoChange with explanation, zero CW calls' ($r.Status -eq 200 -and $r.Json.outcome -eq 'NoChange' -and $r.Json.warnings.Count -eq 1 -and $global:CwCalls.Count -eq 0 -and $global:SpWrites.Count -eq 0)

# 10 live schedule create against a company template: POST, then write-back
Reset-World
$global:Items[8] = [pscustomobject]@{ Id = 8; FieldValues = @{ ParentTemplateId = '5001'; DateStart = [datetime]'2026-11-01'; RecurType = 'w'; RecurInterval = '1'; DaysOfWeekRaw = '2'; Hours = '1' } }
$global:Tpls[5001] = [pscustomobject]@{ id = 5001; name = 'Weekly Patch Review'; templateFlag = $false; department = [pscustomobject]@{ name = 'Service Desk' } }
$global:Cols += @('DateEnd', 'Hours', 'ScheduleStatus', 'MemberId', 'MemberIdentifier', 'TicketOwner', 'ParentTemplateName', 'AbsDayNbr')
$r = Invoke-Fn $sb $false
Assert 'schedule live: Created, one resource POST, recurrence sent' ($r.Json.outcome -eq 'Created' -and @($global:CwCalls | Where-Object { $_.Method -eq 'POST' }).Count -eq 1 -and ((@($global:CwCalls | Where-Object { $_.Method -eq 'POST' })[0].Body | ConvertFrom-Json).recurrence.DaysOfWeek -eq '2'))
Assert 'schedule live: ids written back, registry row, SyncStatus untouched' ($global:Items[8].FieldValues['CWResourceId'] -eq '7000' -and $global:Items[8].FieldValues['PushStatus'] -eq 'Synced' -and $global:SpWrites -contains 'AddItem:R:7000' -and -not ($global:SpWrites | Where-Object { $_ -match 'SyncStatus' }))

Write-Host ''
if ($script:Fail) { Write-Host "$script:Fail test(s) failed" -ForegroundColor Red; exit 1 } else { Write-Host 'All tests passed' -ForegroundColor Green }
