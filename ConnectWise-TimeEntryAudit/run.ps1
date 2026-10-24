#requires -Version 7.2
<#
.SYNOPSIS
Exports ConnectWise PSA time entries CREATED during the previous day (regardless of the
work date), downloads service/project ticket details and notes, scrubs HR records, and
uploads JSON/CSV to SharePoint under <library>/<yyyy>/<MMMM>/<MM-dd> (the report day).
Output per day: ConnectWise-TimeEntries-<date>.{json,csv} (one row per time entry),
ConnectWise-TicketActivity-<date>.{json,csv} (one row per ticket note, or per ticket with
no notes; joins to time entries on ChargeToType + TicketNumber), and
ConnectWise-TimeEntries-<date>.manifest.json.

.NOTES
Runs as an Azure Functions Timer trigger daily and reports the previous day (default
America/New_York). The timer fires at 04:00 and 05:00 UTC so one fire lands just after
local midnight in both daylight and standard time; the other fire exports the same
previous day again and overwrites the same files. The same applies to manual runs: invoke
the function at any time and it exports the previous day, overwriting earlier output.
File names are per report day.
To re-run an earlier day, set TimeEntryAudit_ReportDate=yyyy-MM-dd (must be before
today), invoke the function, then remove the setting.
The Function App must also have AzureWebJobsStorage configured for timer monitoring.
Configure these Function App settings:
  ConnectWisePsa_ApiBaseUrl, ConnectWisePsa_ApiCompanyId, ConnectWisePsa_ApiPublicKey,
  ConnectWisePsa_ApiPrivateKey, ConnectWisePsa_ApiClientId
  TimeEntryAudit_SharePointSiteUrl, TimeEntryAudit_SharePointFolderSiteRelativeUrl
    (defaults: https://palaisparc.sharepoint.com/Finance and 'Time audits')
SharePoint uses certificate authentication with the shared Ms365 settings:
  Ms365_TenantId, Ms365_AuthAppId, Ms365_CertBase64 (PFX as base64), Ms365_CertPassword.
The app registration needs write access to the Finance site (for example Sites.Selected).
Optional settings: TimeEntryAudit_HRMemberIdentifiers (comma-separated),
TimeEntryAudit_HRDepartmentRegex, TimeEntryAudit_HRWorkRoleRegex,
TimeEntryAudit_KeepLocalFiles (true/false), TimeEntryAudit_LocalStagingPath,
TimeEntryAudit_SkipSharePointUpload (true/false),
TimeEntryAudit_BusinessTimeZoneId, TimeEntryAudit_ReportDate.
#>
param($Timer, $TriggerMetadata)

# The Functions PowerShell worker only allows binding names in param(); all other
# settings come from Function App settings (environment variables).
function Get-Setting {
    param([string]$Name, [string]$Default = $null)
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
    return $value
}
function Test-SettingTrue {
    param([string]$Name)
    return (Get-Setting $Name) -match '^(?i:true|1|yes)$'
}

# Version 1 on purpose: ConnectWise omits empty/optional properties from responses, and
# stricter modes throw when reading those properties.
Set-StrictMode -Version 1
$ErrorActionPreference = 'Stop'

[string]$ConnectWiseApiHost = Get-Setting 'ConnectWisePsa_ApiBaseUrl'
[string]$ConnectWiseCompanyId = Get-Setting 'ConnectWisePsa_ApiCompanyId'
[string]$ConnectWisePublicKey = Get-Setting 'ConnectWisePsa_ApiPublicKey'
[string]$ConnectWisePrivateKey = Get-Setting 'ConnectWisePsa_ApiPrivateKey'
[string]$ConnectWiseClientId = Get-Setting 'ConnectWisePsa_ApiClientId'
[string]$ConnectWiseApiVersion = Get-Setting 'ConnectWisePsa_ApiVersion' '2022.1'

[string]$SharePointSiteUrl = Get-Setting 'TimeEntryAudit_SharePointSiteUrl' 'https://palaisparc.sharepoint.com/Finance'
[string]$SharePointFolderSiteRelativeUrl = Get-Setting 'TimeEntryAudit_SharePointFolderSiteRelativeUrl' 'Time audits'
[string]$TenantId = Get-Setting 'Ms365_TenantId'
[string]$SharePointClientId = Get-Setting 'Ms365_AuthAppId'
[string]$CertificateBase64 = Get-Setting 'Ms365_CertBase64'
[string]$CertificatePassword = Get-Setting 'Ms365_CertPassword'
[string]$BusinessTimeZoneId = Get-Setting 'TimeEntryAudit_BusinessTimeZoneId' 'America/New_York'

# Safest HR classification is an explicit member list. Regexes are also evaluated
# against returned member department/default-department, work role/type, and company.
$HRMemberIdentifiers = @((Get-Setting 'TimeEntryAudit_HRMemberIdentifiers' '') -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
[string]$HRDepartmentRegex = Get-Setting 'TimeEntryAudit_HRDepartmentRegex' '(?i)\b(HR|Human Resources|People Operations|People Ops)\b'
[string]$HRWorkRoleRegex = Get-Setting 'TimeEntryAudit_HRWorkRoleRegex' '(?i)\b(HR|Human Resources|People Operations|People Ops)\b'

[string]$LocalStagingPath = Get-Setting 'TimeEntryAudit_LocalStagingPath' (Join-Path ([IO.Path]::GetTempPath()) 'CloudFirst-CW-TimeExport')
$KeepLocalFiles = Test-SettingTrue 'TimeEntryAudit_KeepLocalFiles'
$SkipSharePointUpload = Test-SettingTrue 'TimeEntryAudit_SkipSharePointUpload'
$script:Warnings = [System.Collections.Generic.List[string]]::new()

function Add-ExportWarning {
    param([string]$Message)
    $script:Warnings.Add($Message)
    Write-Warning $Message
}

function Assert-RequiredValue {
    param([string]$Name, [AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw "Required value '$Name' was not supplied." }
}

function Get-PropertyValue {
    param([object]$Object, [string[]]$Paths)
    foreach ($path in $Paths) {
        $current = $Object
        $found = $true
        foreach ($segment in $path.Split('.')) {
            if ($null -eq $current -or $current.PSObject.Properties.Name -notcontains $segment) {
                $found = $false; break
            }
            $current = $current.$segment
        }
        if ($found -and $null -ne $current -and -not [string]::IsNullOrWhiteSpace([string]$current)) {
            return [string]$current
        }
    }
    return $null
}

function Invoke-CWGet {
    param(
        [Parameter(Mandatory)] [string]$RelativePath,
        [hashtable]$Query = @{},
        [ValidateRange(1,10)] [int]$MaxAttempts = 5,
        [ValidateRange(1,300)] [int]$InitialRetryDelaySeconds = 2
    )
    $apiBase = $ConnectWiseApiHost.Trim()
    if ($apiBase -notmatch '^https://') { $apiBase = "https://$apiBase" }
    $apiBaseUri = $null
    if (-not [Uri]::TryCreate($apiBase, [UriKind]::Absolute, [ref]$apiBaseUri) -or
        $apiBaseUri.Scheme -ne 'https' -or
        -not [string]::IsNullOrWhiteSpace($apiBaseUri.Query) -or
        -not [string]::IsNullOrWhiteSpace($apiBaseUri.Fragment)) {
        throw 'ConnectWiseApiHost must be a valid HTTPS API host or base URL without query or fragment.'
    }
    $base = '{0}/v4_6_release/apis/3.0' -f $apiBaseUri.AbsoluteUri.TrimEnd('/')
    $pairs = @(foreach ($key in $Query.Keys) {
        '{0}={1}' -f [uri]::EscapeDataString([string]$key), [uri]::EscapeDataString([string]$Query[$key])
    })
    $uri = "$base/$($RelativePath.TrimStart('/'))"
    if ($pairs.Count -gt 0) { $uri += '?' + ($pairs -join '&') }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return Invoke-RestMethod -Method Get -Uri $uri -Headers $script:CWHeaders -TimeoutSec 180
        }
        catch {
            $statusCode = $null
            $retryAfter = $null
            if ($null -ne $_.Exception.Response) {
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
                try {
                    $retryValues = $_.Exception.Response.Headers.GetValues('Retry-After')
                    if ($retryValues) {
                        $retryValue = [string]($retryValues | Select-Object -First 1)
                        $seconds = 0
                        if ([int]::TryParse($retryValue, [ref]$seconds)) {
                            $retryAfter = $seconds
                        }
                        else {
                            $retryDate = [DateTimeOffset]::MinValue
                            if ([DateTimeOffset]::TryParse($retryValue, [ref]$retryDate)) {
                                $retryAfter = [int][Math]::Max(0, ($retryDate - [DateTimeOffset]::UtcNow).TotalSeconds)
                            }
                        }
                    }
                }
                catch { }
            }

            $isTransient = ($null -eq $statusCode) -or
                           ($statusCode -eq 408) -or
                           ($statusCode -eq 429) -or
                           ($statusCode -ge 500 -and $statusCode -le 599)

            if (-not $isTransient -or $attempt -eq $MaxAttempts) {
                throw "ConnectWise GET failed after $attempt attempt(s). Path: $RelativePath; Status: $statusCode; Error: $($_.Exception.Message)"
            }

            # Honor Retry-After when ConnectWise supplies it. Otherwise use capped
            # exponential backoff plus jitter to avoid synchronized retries.
            if ($retryAfter -and $retryAfter -gt 0) {
                $delay = [Math]::Min($retryAfter, 300)
            }
            else {
                $backoff = $InitialRetryDelaySeconds * [Math]::Pow(2, $attempt - 1)
                $jitter = Get-Random -Minimum 0 -Maximum 1000
                $delay = [Math]::Min(60, $backoff + ($jitter / 1000.0))
            }
            Write-Warning "Transient ConnectWise error for '$RelativePath' (status $statusCode). Attempt $attempt of $MaxAttempts. Retrying in $([Math]::Round($delay,2)) seconds."
            Start-Sleep -Milliseconds ([int]($delay * 1000))
        }
    }
}

function Get-CWPagedCollection {
    param(
        [Parameter(Mandatory)] [string]$RelativePath,
        [hashtable]$Query = @{},
        [ValidateRange(1,1000)] [int]$PageSize = 1000,
        [ValidateRange(1,10)] [int]$MaxAttemptsPerPage = 5
    )
    $items = [System.Collections.Generic.List[object]]::new()
    $page = 1
    do {
        $q = @{} + $Query
        $q.page = $page
        $q.pageSize = $PageSize
        Write-Verbose "Downloading '$RelativePath' page $page with page size $PageSize."
        $batch = @(Invoke-CWGet -RelativePath $RelativePath -Query $q -MaxAttempts $MaxAttemptsPerPage | ForEach-Object { $_ })
        foreach ($item in $batch) { $items.Add($item) }
        Write-Verbose "Downloaded $($batch.Count) record(s) from '$RelativePath' page $page."
        $page++
    } while ($batch.Count -eq $PageSize)
    return $items
}

function Get-MemberDetail {
    param([object]$TimeEntry)
    if ($null -eq $TimeEntry.member -or $null -eq $TimeEntry.member.id) {
        throw "Time entry '$($TimeEntry.id)' has no member ID. Refusing to export a record that cannot be checked for HR data."
    }
    $id = [string]$TimeEntry.member.id
    if (-not $script:MemberCache.ContainsKey($id)) {
        try { $script:MemberCache[$id] = Invoke-CWGet -RelativePath "system/members/$id" }
        catch { throw "Could not retrieve member ID '$id' for time entry '$($TimeEntry.id)'; refusing to export an unclassified record. $($_.Exception.Message)" }
    }
    if ($null -eq $script:MemberCache[$id]) {
        throw "Member ID '$id' returned no details for time entry '$($TimeEntry.id)'. Refusing to export an unclassified record."
    }
    return $script:MemberCache[$id]
}

function Get-ChargeDetail {
    param([object]$TimeEntry)
    $type = [string]$TimeEntry.chargeToType
    $id = [string]$TimeEntry.chargeToId
    if ([string]::IsNullOrWhiteSpace($id)) { return $null }
    $key = "$type`:$id"
    if (-not $script:ChargeCache.ContainsKey($key)) {
        try {
            $script:ChargeCache[$key] = switch ($type) {
                'ServiceTicket' { Invoke-CWGet -RelativePath "service/tickets/$id" }
                'ProjectTicket' { Invoke-CWGet -RelativePath "project/tickets/$id" }
                'Activity'      { Invoke-CWGet -RelativePath "sales/activities/$id" }
                default         { $null }
            }
        }
        catch { $script:ChargeCache[$key] = $null; Add-ExportWarning "Could not enrich '$key'. $($_.Exception.Message)" }
    }
    return $script:ChargeCache[$key]
}

function Get-TicketNotes {
    param([object]$TimeEntry)
    $type = [string]$TimeEntry.chargeToType
    $id = [string]$TimeEntry.chargeToId
    if ($type -notin @('ServiceTicket','ProjectTicket') -or [string]::IsNullOrWhiteSpace($id)) {
        return @()
    }

    $key = "$type`:$id"
    if (-not $script:TicketNotesCache.ContainsKey($key)) {
        try {
            $notes = $null
            $attempts = if ($type -eq 'ServiceTicket') {
                @(@{ Path = "service/tickets/$id/notes"; Query = @{ orderBy = 'dateCreated asc' } })
            }
            else {
                # Project ticket notes reject some query options with a 400. Retry without
                # ordering, then fall back to the service ticket notes endpoint.
                @(
                    @{ Path = "project/tickets/$id/notes"; Query = @{ orderBy = 'dateCreated asc' } },
                    @{ Path = "project/tickets/$id/notes"; Query = @{} },
                    @{ Path = "service/tickets/$id/notes"; Query = @{ orderBy = 'dateCreated asc' } }
                )
            }
            $lastError = $null
            foreach ($attempt in $attempts) {
                try {
                    $notes = @(Get-CWPagedCollection -RelativePath $attempt.Path -Query $attempt.Query -PageSize 1000 -MaxAttemptsPerPage 5)
                    break
                }
                catch { $lastError = $_ }
            }
            if ($null -eq $notes) { throw $lastError }
            $script:TicketNotesCache[$key] = @($notes | Sort-Object { Get-PropertyValue $_ @('dateCreated','_info.dateEntered','_info.lastUpdated') })
        }
        catch {
            $script:TicketNotesCache[$key] = @()
            Add-ExportWarning "Could not retrieve ticket notes for '$key'. $($_.Exception.Message)"
        }
    }
    return @($script:TicketNotesCache[$key])
}

function Convert-TicketNotesToText {
    param([object[]]$Notes)
    $blocks = foreach ($note in @($Notes)) {
        $created = Get-PropertyValue $note @('dateCreated','_info.dateEntered','_info.lastUpdated')
        $creator = Get-PropertyValue $note @('member.name','member.identifier','contact.name','createdBy','_info.updatedBy')
        $flags = @()
        $internal = Get-PropertyValue $note @('internalAnalysisFlag','internalFlag')
        $discussion = Get-PropertyValue $note @('detailDescriptionFlag')
        $resolution = Get-PropertyValue $note @('resolutionFlag')
        $external = Get-PropertyValue $note @('externalFlag')
        if ($internal -eq 'True') { $flags += 'Internal' }
        if ($discussion -eq 'True') { $flags += 'Discussion' }
        if ($resolution -eq 'True') { $flags += 'Resolution' }
        if ($external -eq 'True') { $flags += 'External' }
        $kind = if ($flags.Count) { $flags -join ',' } else { 'Note' }
        $text = Get-PropertyValue $note @('text','notes','note')
        "[$created] [$kind] [$creator] $text"
    }
    return ($blocks -join "`n---`n")
}

function Test-IsHRRecord {
    param([object]$TimeEntry, [object]$MemberDetail)
    $memberIdentifier = Get-PropertyValue $TimeEntry @('member.identifier')
    if ($memberIdentifier -and $HRMemberIdentifiers -contains $memberIdentifier) { return $true }

    $department = Get-PropertyValue $MemberDetail @(
        'department.name','department.identifier','defaultDepartment.name','defaultDepartment.identifier',
        'businessUnit.name','title'
    )
    $role = Get-PropertyValue $TimeEntry @('workRole.name','workRole.identifier','workType.name','workType.identifier')
    $company = Get-PropertyValue $TimeEntry @('company.name','company.identifier')
    return (($department -and $department -match $HRDepartmentRegex) -or
            ($role -and $role -match $HRWorkRoleRegex) -or
            ($company -and $company -match $HRDepartmentRegex))
}

function New-ExportRecord {
    param([object]$Entry, [object]$Member, [object]$Charge, [object[]]$TicketNotes, [bool]$IsHR)
    $ticketNumber = if ($Entry.chargeToType -in @('ServiceTicket','ProjectTicket')) { $Entry.chargeToId } else { $null }
    $department = Get-PropertyValue $Member @('department.name','department.identifier','defaultDepartment.name','defaultDepartment.identifier','businessUnit.name')

    if ($IsHR) {
        return [ordered]@{
            TimeEntryId = $Entry.id
            HRRedacted = $true
            MemberIdentifier = '[HR-REDACTED]'
            MemberName = '[HR-REDACTED]'
            Department = 'Human Resources'
            DateEnteredUtc = $Entry.dateEntered
            EnteredBy = '[HR-REDACTED]'
            WorkDate = if ($Entry.timeStart) { ([datetimeoffset]$Entry.timeStart).ToString('yyyy-MM-dd') } else { $null }
            TimeStart = $null
            TimeEnd = $null
            ActualHours = $Entry.actualHours
            HoursBilled = $null
            BillableOption = $Entry.billableOption
            Status = $null
            CompanyIdentifier = '[HR-REDACTED]'
            CompanyName = '[HR-REDACTED]'
            ChargeToType = '[HR-REDACTED]'
            ChargeToId = $null
            TicketNumber = $null
            TicketSummary = '[HR-REDACTED]'
            TicketNotesCount = $null
            TicketNotes = '[HR DETAIL AND PII REDACTED]'
            TicketNotesJson = '[HR DETAIL AND PII REDACTED]'
            TicketStatus = $null
            BoardOrProject = '[HR-REDACTED]'
            TicketType = '[HR-REDACTED]'
            TicketSubtype = '[HR-REDACTED]'
            TicketItem = '[HR-REDACTED]'
            TicketPriority = '[HR-REDACTED]'
            TicketInitialDescription = '[HR DETAIL AND PII REDACTED]'
            TicketResolution = '[HR DETAIL AND PII REDACTED]'
            TicketDateEntered = $null
            TicketClosedDate = $null
            WorkRole = '[HR-REDACTED]'
            WorkType = '[HR-REDACTED]'
            Agreement = '[HR-REDACTED]'
            Notes = '[HR DETAIL AND PII REDACTED]'
            InternalNotes = '[HR DETAIL AND PII REDACTED]'
            FullTimeDetail = [ordered]@{
                TimeEntryId = $Entry.id
                HRRedacted = $true
                DateEnteredUtc = $Entry.dateEntered
                WorkDate = if ($Entry.timeStart) { ([datetimeoffset]$Entry.timeStart).ToString('yyyy-MM-dd') } else { $null }
                ActualHours = $Entry.actualHours
                BillableOption = $Entry.billableOption
            }
        }
    }

    return [ordered]@{
        TimeEntryId = $Entry.id
        HRRedacted = $false
        MemberIdentifier = Get-PropertyValue $Entry @('member.identifier')
        MemberName = Get-PropertyValue $Entry @('member.name')
        Department = $department
        DateEnteredUtc = $Entry.dateEntered
        EnteredBy = $Entry.enteredBy
        WorkDate = if ($Entry.timeStart) { ([datetimeoffset]$Entry.timeStart).ToString('yyyy-MM-dd') } else { $null }
        TimeStart = $Entry.timeStart
        TimeEnd = $Entry.timeEnd
        ActualHours = $Entry.actualHours
        HoursBilled = $Entry.hoursBilled
        BillableOption = $Entry.billableOption
        Status = $Entry.status
        CompanyIdentifier = Get-PropertyValue $Entry @('company.identifier')
        CompanyName = Get-PropertyValue $Entry @('company.name')
        ChargeToType = $Entry.chargeToType
        ChargeToId = $Entry.chargeToId
        TicketNumber = $ticketNumber
        TicketSummary = Get-PropertyValue $Charge @('summary','name','subject')
        TicketNotesCount = @($TicketNotes).Count
        TicketNotes = Convert-TicketNotesToText -Notes $TicketNotes
        TicketNotesJson = if (@($TicketNotes).Count) { $TicketNotes | ConvertTo-Json -Depth 20 -Compress } else { '[]' }
        TicketStatus = Get-PropertyValue $Charge @('status.name','status.identifier')
        BoardOrProject = Get-PropertyValue $Charge @('board.name','project.name')
        TicketType = Get-PropertyValue $Charge @('type.name','type.identifier')
        TicketSubtype = Get-PropertyValue $Charge @('subType.name','subType.identifier')
        TicketItem = Get-PropertyValue $Charge @('item.name','item.identifier')
        TicketPriority = Get-PropertyValue $Charge @('priority.name','priority.identifier')
        TicketInitialDescription = Get-PropertyValue $Charge @('initialDescription','description')
        TicketResolution = Get-PropertyValue $Charge @('resolution')
        TicketDateEntered = Get-PropertyValue $Charge @('dateEntered')
        TicketClosedDate = Get-PropertyValue $Charge @('closedDate','dateClosed')
        WorkRole = Get-PropertyValue $Entry @('workRole.name','workRole.identifier')
        WorkType = Get-PropertyValue $Entry @('workType.name','workType.identifier')
        Agreement = Get-PropertyValue $Entry @('agreement.name','agreement.identifier')
        Notes = $Entry.notes
        InternalNotes = $Entry.internalNotes
        FullTimeDetail = $Entry
        ChargeDetail = $Charge
    }
}

Assert-RequiredValue 'ConnectWiseApiHost' $ConnectWiseApiHost
Assert-RequiredValue 'ConnectWiseCompanyId' $ConnectWiseCompanyId
Assert-RequiredValue 'ConnectWisePublicKey' $ConnectWisePublicKey
Assert-RequiredValue 'ConnectWisePrivateKey' $ConnectWisePrivateKey
Assert-RequiredValue 'ConnectWiseClientId' $ConnectWiseClientId

$authPlain = "$ConnectWiseCompanyId+$ConnectWisePublicKey`:$ConnectWisePrivateKey"
$authEncoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($authPlain))
$script:CWHeaders = @{
    Authorization = "Basic $authEncoded"
    ClientId = $ConnectWiseClientId
    Accept = "application/vnd.connectwise.com+json; version=$ConnectWiseApiVersion"
}
$script:MemberCache = @{}
$script:ChargeCache = @{}
$script:TicketNotesCache = @{}

$authModes = @()
if (-not $SkipSharePointUpload) {
    Assert-RequiredValue 'SharePointSiteUrl' $SharePointSiteUrl
    Assert-RequiredValue 'SharePointFolderSiteRelativeUrl' $SharePointFolderSiteRelativeUrl
    Assert-RequiredValue 'Ms365_TenantId' $TenantId
    Assert-RequiredValue 'Ms365_AuthAppId' $SharePointClientId
    Assert-RequiredValue 'Ms365_CertBase64' $CertificateBase64
    Assert-RequiredValue 'Ms365_CertPassword' $CertificatePassword
    $authModes = @('Certificate')
    Import-Module PnP.PowerShell -ErrorAction Stop
}

$tz = [TimeZoneInfo]::FindSystemTimeZoneById($BusinessTimeZoneId)
$nowLocal = [TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, $tz).DateTime

$timerPastDue = $false
if ($null -ne $Timer -and $null -ne $Timer.PSObject.Properties['IsPastDue']) {
    $timerPastDue = [bool]$Timer.IsPastDue
}
if ($timerPastDue) {
    Write-Warning 'The timer invocation is past due; exporting the previous day now.'
}

# Report day: the previous local day by default, or TimeEntryAudit_ReportDate (yyyy-MM-dd)
# to re-run an earlier day. Remove that setting afterwards or the nightly run will
# keep re-exporting it.
$reportDateOverride = Get-Setting 'TimeEntryAudit_ReportDate'
if ($reportDateOverride) {
    $parsedDate = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($reportDateOverride.Trim(), 'yyyy-MM-dd', [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
        throw "TimeEntryAudit_ReportDate '$reportDateOverride' is not a valid yyyy-MM-dd date."
    }
    if ($parsedDate.Date -ge $nowLocal.Date) {
        throw "TimeEntryAudit_ReportDate '$reportDateOverride' must be earlier than today ($($nowLocal.ToString('yyyy-MM-dd')) $BusinessTimeZoneId)."
    }
    $reportDate = $parsedDate.Date
    Write-Host "Using report date override $($reportDate.ToString('yyyy-MM-dd'))."
}
else {
    $reportDate = $nowLocal.Date.AddDays(-1)
}

$startLocal = $reportDate
$endLocal = $reportDate.AddDays(1)
$windowStartUtc = [DateTimeOffset][TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($startLocal, [DateTimeKind]::Unspecified), $tz)
$windowEndUtc = [DateTimeOffset][TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($endLocal, [DateTimeKind]::Unspecified), $tz)
$startApi = $windowStartUtc.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ')
$endApi = $windowEndUtc.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ')

# dateEntered is the creation timestamp. This intentionally does NOT filter timeStart/work date.
$condition = "dateEntered >= [$startApi] and dateEntered < [$endApi]"
Write-Host "Collecting time entries created on $($reportDate.ToString('yyyy-MM-dd')) between $startApi and $endApi (UTC)."
$entries = @(Get-CWPagedCollection -RelativePath 'time/entries' -Query @{ conditions = $condition; orderBy = 'dateEntered asc' })

$records = [System.Collections.Generic.List[object]]::new()
foreach ($entry in $entries) {
    $member = Get-MemberDetail -TimeEntry $entry
    $isHR = Test-IsHRRecord -TimeEntry $entry -MemberDetail $member
    $charge = if ($isHR) { $null } else { Get-ChargeDetail -TimeEntry $entry }
    $ticketNotes = if ($isHR) { @() } else { @(Get-TicketNotes -TimeEntry $entry) }
    $records.Add([pscustomobject](New-ExportRecord -Entry $entry -Member $member -Charge $charge -TicketNotes $ticketNotes -IsHR $isHR))
}

function New-TicketActivityRows {
    # One row per ticket note (or one row for a ticket with no notes). Joins to the time
    # entry dataset on ChargeToType + TicketNumber. HR-redacted entries are never included.
    param([object[]]$Records)
    $tickets = @($Records | Where-Object { -not $_.HRRedacted -and $_.TicketNumber } | Group-Object { "$($_.ChargeToType):$($_.TicketNumber)" })
    foreach ($group in $tickets) {
        $first = $group.Group[0]
        $key = $group.Name
        $charge = $script:ChargeCache[$key]
        $notes = @($script:TicketNotesCache[$key])
        $entryIds = @($group.Group | ForEach-Object { $_.TimeEntryId })
        $totalHours = [double](($group.Group | Measure-Object -Property ActualHours -Sum).Sum)
        $base = [ordered]@{
            ChargeToType = $first.ChargeToType
            TicketNumber = $first.TicketNumber
            TicketSummary = $first.TicketSummary
            CompanyName = $first.CompanyName
            TicketStatus = $first.TicketStatus
            BoardOrProject = $first.BoardOrProject
            TicketType = $first.TicketType
            TicketSubtype = $first.TicketSubtype
            TicketPriority = $first.TicketPriority
            TicketDateEntered = $first.TicketDateEntered
            TicketClosedDate = $first.TicketClosedDate
            TimeEntryIds = ($entryIds -join ';')
            TimeEntryCount = $entryIds.Count
            TimeEntryHoursInReport = $totalHours
        }
        if ($notes.Count -eq 0) {
            [pscustomobject]($base + [ordered]@{
                ActivityType = 'Ticket'; NoteId = $null; NoteDateCreated = $null; NoteCreatedBy = $null
                NoteFlags = $null; NoteText = $null
            })
            continue
        }
        foreach ($note in $notes) {
            $flags = @()
            if ((Get-PropertyValue $note @('internalAnalysisFlag','internalFlag')) -eq 'True') { $flags += 'Internal' }
            if ((Get-PropertyValue $note @('detailDescriptionFlag')) -eq 'True') { $flags += 'Discussion' }
            if ((Get-PropertyValue $note @('resolutionFlag')) -eq 'True') { $flags += 'Resolution' }
            if ((Get-PropertyValue $note @('externalFlag')) -eq 'True') { $flags += 'External' }
            [pscustomobject]($base + [ordered]@{
                ActivityType = 'Note'
                NoteId = Get-PropertyValue $note @('id')
                NoteDateCreated = Get-PropertyValue $note @('dateCreated','_info.dateEntered','_info.lastUpdated')
                NoteCreatedBy = Get-PropertyValue $note @('member.name','member.identifier','contact.name','createdBy','_info.updatedBy')
                NoteFlags = ($flags -join ',')
                NoteText = Get-PropertyValue $note @('text','notes','note')
            })
        }
    }
}

New-Item -ItemType Directory -Path $LocalStagingPath -Force | Out-Null
$reportStamp = $reportDate.ToString('yyyy-MM-dd')
$fileStem = "ConnectWise-TimeEntries-$reportStamp"
$jsonPath = Join-Path $LocalStagingPath "$fileStem.json"
$csvPath = Join-Path $LocalStagingPath "$fileStem.csv"
$manifestPath = Join-Path $LocalStagingPath "$fileStem.manifest.json"
$activityStem = "ConnectWise-TicketActivity-$reportStamp"
$activityJsonPath = Join-Path $LocalStagingPath "$activityStem.json"
$activityCsvPath = Join-Path $LocalStagingPath "$activityStem.csv"
$exportPaths = @($jsonPath, $csvPath, $activityJsonPath, $activityCsvPath, $manifestPath)

$activityRows = [System.Collections.Generic.List[object]]::new()
foreach ($row in @(New-TicketActivityRows -Records @($records.ToArray()))) { $activityRows.Add($row) }

ConvertTo-Json -InputObject @($records.ToArray()) -Depth 100 | Set-Content -Path $jsonPath -Encoding utf8
$records | Select-Object * -ExcludeProperty FullTimeDetail,ChargeDetail | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8
ConvertTo-Json -InputObject @($activityRows.ToArray()) -Depth 20 | Set-Content -Path $activityJsonPath -Encoding utf8
$activityRows | Export-Csv -Path $activityCsvPath -NoTypeInformation -Encoding utf8
$manifest = [ordered]@{
    ExportCreatedUtc = [datetime]::UtcNow.ToString('o')
    ReportDate = $reportDate.ToString('yyyy-MM-dd')
    ReportTimeZone = $BusinessTimeZoneId
    DateEnteredWindowUtc = [ordered]@{ StartInclusive = $startApi; EndExclusive = $endApi }
    Filter = $condition
    RecordCount = $records.Count
    TicketCount = @($activityRows | Select-Object -ExpandProperty TicketNumber -Unique).Count
    TicketActivityRowCount = $activityRows.Count
    HRRecordCount = @($records | Where-Object HRRedacted).Count
    WarningCount = $script:Warnings.Count
    Warnings = @($script:Warnings.ToArray())
    Files = @($exportPaths | ForEach-Object { [IO.Path]::GetFileName($_) })
}
$manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $manifestPath -Encoding utf8

try {
    if (-not $SkipSharePointUpload) {
        $securePassword = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
        Connect-PnPOnline -Url $SharePointSiteUrl -ClientId $SharePointClientId -Tenant $TenantId -CertificateBase64Encoded $CertificateBase64 -CertificatePassword $securePassword

        $targetFolder = [string]::Format([cultureinfo]::InvariantCulture, '{0}/{1:yyyy}/{1:MMMM}/{1:MM-dd}', $SharePointFolderSiteRelativeUrl.Trim('/'), $reportDate)
        Resolve-PnPFolder -SiteRelativePath $targetFolder | Out-Null

        foreach ($path in $exportPaths) {
            Add-PnPFile -Path $path -Folder $targetFolder | Out-Null
            Write-Host "Uploaded $([IO.Path]::GetFileName($path)) to $targetFolder"
        }
    }
}
finally {
    if (-not $SkipSharePointUpload -and $authModes.Count -gt 0) {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    if (-not $KeepLocalFiles) {
        try {
            Remove-Item -Path $exportPaths -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not remove staged export files from '$LocalStagingPath'. Check the Function App temporary storage and access permissions. $($_.Exception.Message)"
        }
    }
}

Write-Host "Complete. Exported $($records.Count) entries; HR-redacted: $(@($records | Where-Object HRRedacted).Count); warnings: $($script:Warnings.Count)."
