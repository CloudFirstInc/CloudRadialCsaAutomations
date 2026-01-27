using namespace System.Net

param($Request, $TriggerMetadata)

function Write-JsonResponse {
    param(
        [Parameter(Mandatory)] [int] $StatusCode,
        [Parameter(Mandatory)]       $BodyObject
    )
    $json = $BodyObject | ConvertTo-Json -Depth 25
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = $StatusCode
        Headers    = @{ "Content-Type" = "application/json" }
        Body       = $json
    })
}

function Write-Log {
    param(
        [ValidateSet('Info','Debug','Warn','Error')] [string] $Level = 'Info',
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('Info','Debug')] [string] $ConfiguredLevel = 'Info'
    )
    if ($Level -eq 'Debug' -and $ConfiguredLevel -ne 'Debug') { return }
    Write-Host "[$Level] $Message"
}

function Get-LoginHost {
    param([ValidateSet('Global','USGov')][string] $GraphCloud = 'Global')
    switch ($GraphCloud.ToLower()) {
        'usgov' { 'login.microsoftonline.us' }
        default { 'login.microsoftonline.com' }
    }
}

function Get-GraphHost {
    param([ValidateSet('Global','USGov')][string] $GraphCloud = 'Global')
    switch ($GraphCloud.ToLower()) {
        'usgov' { 'graph.microsoft.us' }
        default { 'graph.microsoft.com' }
    }
}

function Resolve-TenantId {
    param(
        [Parameter(Mandatory)][string] $CustomerTenant,
        [ValidateSet('Global','USGov')][string] $GraphCloud = 'Global',
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )

    if ($CustomerTenant -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        Write-Log -Level Debug -Message "Resolve-TenantId: '$CustomerTenant' is a GUID." -ConfiguredLevel $LogLevel
        return $CustomerTenant
    }

    $loginHost = Get-LoginHost -GraphCloud $GraphCloud
    $wellKnown = "https://$loginHost/$CustomerTenant/v2.0/.well-known/openid-configuration"
    Write-Log -Level Debug -Message "Resolve-TenantId: GET $wellKnown" -ConfiguredLevel $LogLevel

    try {
        $meta   = Invoke-RestMethod -Method GET -Uri $wellKnown -ErrorAction Stop
        $issuer = [Uri]$meta.issuer
        $segments = $issuer.AbsolutePath.Trim('/').Split('/')
        if ($segments.Length -ge 1 -and $segments[0] -match '^[0-9a-fA-F-]{36}$') {
            $tid = $segments[0]
            Write-Log -Level Debug -Message "Resolve-TenantId: domain '$CustomerTenant' -> tenantId '$tid'." -ConfiguredLevel $LogLevel
            return $tid
        }
        throw "Issuer did not contain a tenant GUID."
    }
    catch {
        throw "Could not resolve tenant from domain '$CustomerTenant'. Details: $($_.Exception.Message)"
    }
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PATCH','DELETE')] [string] $Method,
        [Parameter(Mandatory)][string] $Uri,
        $BodyObject,
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )
    $bodyJson = $null
    if ($null -ne $BodyObject) { $bodyJson = ($BodyObject | ConvertTo-Json -Depth 25) }

    Write-Log -Level Debug -ConfiguredLevel $LogLevel -Message ("Invoke-Graph: $Method $Uri" + ($(if($bodyJson){ "`n$bodyJson" } else { "" })))

    if ($null -ne $BodyObject) {
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $bodyJson -ContentType "application/json" -ErrorAction Stop
    } else {
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -ErrorAction Stop
    }
}

function Get-LatestWindowsSecurityBaselineTemplate {
    param(
        [Parameter(Mandatory)][ValidateSet('Global','USGov')] [string] $GraphCloud,
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )
    $graphHost = Get-GraphHost -GraphCloud $GraphCloud
    $uri = "https://$graphHost/beta/deviceManagement/templates?`$top=999"
    $resp = Invoke-Graph -Method GET -Uri $uri -LogLevel $LogLevel
    $templates = @($resp.value)

    $candidates = $templates | Where-Object {
        $_.templateType -eq 'securityBaseline' -and
        $_.platformType -eq 'windows10AndLater' -and
        ($_.displayName -match 'Security Baseline') -and
        ($_.displayName -match 'Windows')
    }
    if (-not $candidates) { throw "No Windows securityBaseline templates found in /deviceManagement/templates." }

    $notDeprecated = $candidates | Where-Object { $_.isDeprecated -ne $true }
    if ($notDeprecated) { $candidates = $notDeprecated }

    $picked =
        ($candidates | Sort-Object `
            @{ Expression = { if ($_.publishedDateTime) { [datetime]$_.publishedDateTime } else { [datetime]'1900-01-01' } }; Descending = $true }, `
            @{ Expression = { $_.versionInfo }; Descending = $true } `
        | Select-Object -First 1)

    return $picked
}

function Find-IntentByDisplayName {
    param(
        [Parameter(Mandatory)][ValidateSet('Global','USGov')] [string] $GraphCloud,
        [Parameter(Mandatory)][string] $DisplayName,
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )
    $graphHost = Get-GraphHost -GraphCloud $GraphCloud
    $uri = "https://$graphHost/beta/deviceManagement/intents?`$top=999"
    $resp = Invoke-Graph -Method GET -Uri $uri -LogLevel $LogLevel
    $match = @($resp.value) | Where-Object { $_.displayName -eq $DisplayName } | Select-Object -First 1
    return $match
}

function Create-IntentFromTemplate {
    param(
        [Parameter(Mandatory)][ValidateSet('Global','USGov')] [string] $GraphCloud,
        [Parameter(Mandatory)][string] $TemplateId,
        [Parameter(Mandatory)][string] $DisplayName,
        [string] $Description,
        [string[]] $RoleScopeTagIds = @('0'),
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )
    $graphHost = Get-GraphHost -GraphCloud $GraphCloud
    $body = @{
        "@odata.type"    = "#microsoft.graph.deviceManagementIntent"
        displayName      = $DisplayName
        description      = $Description
        templateId       = $TemplateId
        roleScopeTagIds  = $RoleScopeTagIds
    }
    $uri = "https://$graphHost/beta/deviceManagement/intents"
    return Invoke-Graph -Method POST -Uri $uri -BodyObject $body -LogLevel $LogLevel
}

function Assign-IntentToGroups {
    param(
        [Parameter(Mandatory)][ValidateSet('Global','USGov')] [string] $GraphCloud,
        [Parameter(Mandatory)][string] $IntentId,
        [Parameter(Mandatory)][string[]] $GroupIds,
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )
    if (-not $GroupIds -or $GroupIds.Count -eq 0) { return }
    $graphHost = Get-GraphHost -GraphCloud $GraphCloud

    $assignments = @()
    foreach ($gid in $GroupIds) {
        $assignments += @{
            target = @{
                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                groupId       = $gid
            }
        }
    }
    $body = @{ assignments = $assignments }
    $uri  = "https://$graphHost/beta/deviceManagement/intents/$IntentId/assign"
    Invoke-Graph -Method POST -Uri $uri -BodyObject $body -LogLevel $LogLevel | Out-Null
}

# --------------------
# NEW: Group utilities
# --------------------
function Get-DefaultDomainFromTenant {
    param(
        [Parameter(Mandatory)][ValidateSet('Global','USGov')] [string] $GraphCloud,
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )
    $graphHost = Get-GraphHost -GraphCloud $GraphCloud
    $uri = "https://$graphHost/v1.0/domains"
    $resp = Invoke-Graph -Method GET -Uri $uri -LogLevel $LogLevel   # List domains (v1.0) [3](https://learn.microsoft.com/en-us/graph/api/domain-list?view=graph-rest-1.0)
    $domains = @($resp.value)
    if (-not $domains) { throw "No domains returned for tenant." }
    $default = $domains | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1
    if ($default) { return $default.id }
    # Fallback to any verified
    $verified = $domains | Where-Object { $_.isVerified -eq $true } | Select-Object -First 1
    if ($verified) { return $verified.id }
    # Last resort: first
    return ($domains | Select-Object -First 1).id
}

function Get-CustomerPrefix {
    param(
        [Parameter(Mandatory)][string] $CustomerTenant,
        [Parameter(Mandatory)][ValidateSet('Global','USGov')] [string] $GraphCloud,
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )
    $domain = $null
    if ($CustomerTenant -match '\.') {
        $domain = $CustomerTenant
    } else {
        # It's a GUID; fetch default domain
        $domain = Get-DefaultDomainFromTenant -GraphCloud $GraphCloud -LogLevel $LogLevel
    }

    $firstLabel = ($domain -split '\.')[0]
    if ([string]::IsNullOrWhiteSpace($firstLabel)) { $firstLabel = 'Client' }
    # Title-case first letter, keep remainder lower (Ryeny)
    $prefix = ($firstLabel.Substring(0,1).ToUpper() + $(if($firstLabel.Length -gt 1){ $firstLabel.Substring(1).ToLower() } else { "" }))
    return $prefix
}

function New-MailNicknameFromDisplayName {
    param([Parameter(Mandatory)][string] $DisplayName)
    # Safe mailNickname (letters, digits). Remove non-alnum; trim to 64 chars; lower.
    $nick = ($DisplayName -replace '[^a-zA-Z0-9]', '')
    if ([string]::IsNullOrWhiteSpace($nick)) { $nick = "group$(Get-Random -Minimum 1000 -Maximum 9999)" }
    if ($nick.Length -gt 64) { $nick = $nick.Substring(0,64) }
    return $nick.ToLower()
}

function Ensure-Group {
    param(
        [Parameter(Mandatory)][ValidateSet('Global','USGov')] [string] $GraphCloud,
        [Parameter(Mandatory)][string] $DisplayName,
        [Parameter(Mandatory)][string] $Description,
        [ValidateSet('Info','Debug')] [string] $LogLevel = 'Info'
    )
    $graphHost = Get-GraphHost -GraphCloud $GraphCloud
    # Find by exact displayName
    $filterName = [System.Web.HttpUtility]::UrlEncode("displayName eq '$DisplayName'")
    $getUri = "https://$graphHost/v1.0/groups?`$filter=$filterName"
    $found = Invoke-Graph -Method GET -Uri $getUri -LogLevel $LogLevel
    $exists = @($found.value) | Select-Object -First 1
    if ($exists) {
        return @{ id=$exists.id; displayName=$exists.displayName; existed=$true }
    }

    # Create Security group (mailDisabled)
    $mailNick = New-MailNicknameFromDisplayName -DisplayName $DisplayName
    $body = @{
        displayName     = $DisplayName
        description     = $Description
        mailEnabled     = $false
        mailNickname    = $mailNick
        securityEnabled = $true
        groupTypes      = @()  # security group
    }
    $postUri = "https://$graphHost/v1.0/groups"
    $newGrp = Invoke-Graph -Method POST -Uri $postUri -BodyObject $body -LogLevel $LogLevel   # Create group (v1.0) [2](https://learn.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0)
    return @{ id=$newGrp.id; displayName=$newGrp.displayName; existed=$false }
}

# -------------------------
# Main
# -------------------------
$correlationId = [guid]::NewGuid().ToString()

try {
    # Parse body
    $payload = $null
    $rawJson = $null
    if ($Request.PSObject.Properties.Name -contains 'RawBody' -and $Request.RawBody) {
        $rawJson = [string]$Request.RawBody
    } elseif ($Request.Body -is [string]) {
        $rawJson = $Request.Body
    } elseif ($Request.Body -is [System.IO.Stream]) {
        $reader = New-Object IO.StreamReader($Request.Body)
        $rawJson = $reader.ReadToEnd()
    } elseif ($Request.Body -is [System.Collections.IDictionary]) {
        $payload = [hashtable]$Request.Body
    }
    if (-not $payload -and -not [string]::IsNullOrWhiteSpace($rawJson)) {
        $payload = $rawJson | ConvertFrom-Json -ErrorAction Stop
    }
    if (-not $payload) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error="Empty request body."; correlationId=$correlationId }
        return
    }

    # Inputs
    $customerTenant = $payload.CustomerTenant
    if (-not $customerTenant) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error="CustomerTenant is required (GUID or domain)."; correlationId=$correlationId }
        return
    }

    $graphCloud = if ($payload.GraphCloud) { [string]$payload.GraphCloud } else { 'Global' }   # Global | USGov
    $logLevel   = if ($payload.LogLevel)   { [string]$payload.LogLevel }   else { 'Info' }     # Info | Debug
    $corrFromIn = if ($payload.CorrelationId) { [string]$payload.CorrelationId } else { $correlationId }
    $dryRun     = [bool]$payload.DryRun

    # Optional scope tags & explicit group assignments (as before)
    $roleScopeTagIds = @('0')
    if ($payload.RoleScopeTagIds) { $roleScopeTagIds = @($payload.RoleScopeTagIds) }
    $pilotGroupIds = if ($payload.PilotAssignmentGroupIds) { @($payload.PilotAssignmentGroupIds) } else { @() }
    $broadGroupIds = if ($payload.BroadAssignmentGroupIds) { @($payload.BroadAssignmentGroupIds) } else { @() }

    # Default display names
    $pilotName = if ($payload.PilotDisplayName) { [string]$payload.PilotDisplayName } else { "Baseline – Windows 11 – Level 1 (Pilot)" }
    $broadName = if ($payload.BroadDisplayName) { [string]$payload.BroadDisplayName } else { "Baseline – Windows 11 – Level 1 (Broad)" }

    Write-Log -Level Debug -ConfiguredLevel $logLevel -Message ("Inputs: " + (@{
        CustomerTenant = $customerTenant
        GraphCloud     = $graphCloud
        DryRun         = $dryRun
        RoleScopeTags  = $roleScopeTagIds
        PilotName      = $pilotName
        BroadName      = $broadName
        ProvidedPilotAssignments = $pilotGroupIds
        ProvidedBroadAssignments = $broadGroupIds
        CorrelationId  = $corrFromIn
    } | ConvertTo-Json -Depth 8))

    # Resolve tenant GUID
    $tenantId = Resolve-TenantId -CustomerTenant $customerTenant -GraphCloud $graphCloud -LogLevel $logLevel

    # Credentials from App Settings
    $clientId     = $env:Ms365_AuthAppId
    $clientSecret = $env:Ms365_AuthSecretId
    if (-not $clientId -or -not $clientSecret) {
        Write-JsonResponse -StatusCode 500 -BodyObject @{
            error="Missing Graph credentials. Set App Settings Ms365_AuthAppId and Ms365_AuthSecretId."
            correlationId=$corrFromIn
        }
        return
    }

    # Connect (app-only)
    $envName = if ($graphCloud -ieq 'USGov') { 'USGov' } else { 'Global' }
    $secure  = ConvertTo-SecureString $clientSecret -AsPlainText -Force
    $creds   = New-Object System.Management.Automation.PSCredential($clientId, $secure)
    $connectParams = @{
        TenantId               = $tenantId
        ClientSecretCredential = $creds
        Environment            = $envName
        NoWelcome              = $true
    }
    Write-Log -Level Debug -Message "Connecting to Graph. TenantId=$tenantId, Environment=$envName" -ConfiguredLevel $logLevel
    Connect-MgGraph @connectParams | Out-Null

    # -----------------------
    # NEW: Ensure 4 groups
    # -----------------------
    $prefix = Get-CustomerPrefix -CustomerTenant $customerTenant -GraphCloud $graphCloud -LogLevel $logLevel
    $today  = (Get-Date).ToString('yyyy-MM-dd')
    $descSuffix = "Created by CloudFirst on $today"

    $grpNames = [ordered]@{
        ITDevices  = "$prefix-IT-Devices"
        AllDevices = "$prefix-All-Devices"
        ITUsers    = "$prefix-IT-Users"
        AllUser    = "$prefix-All-User"   # per your requested name (singular)
    }

    $grpResults = @{}
    foreach ($k in $grpNames.Keys) {
        $name = $grpNames[$k]
        $desc = switch ($k) {
            'ITDevices'  { "IT-managed device group. $descSuffix" }
            'AllDevices' { "All managed devices group. $descSuffix" }
            'ITUsers'    { "IT-managed users group. $descSuffix" }
            'AllUser'    { "All users group. $descSuffix" }
        }
        if ($dryRun) {
            $grpResults[$k] = @{ id=$null; displayName=$name; existed=$false; dryRun=$true }
        } else {
            $ensured = Ensure-Group -GraphCloud $graphCloud -DisplayName $name -Description $desc -LogLevel $logLevel
            $grpResults[$k] = $ensured
        }
    }

    # If caller didn't provide explicit assignment group IDs, use device groups by default:
    if (-not $pilotGroupIds -or $pilotGroupIds.Count -eq 0)  { if ($grpResults.ITDevices.id)  { $pilotGroupIds += $grpResults.ITDevices.id } }
    if (-not $broadGroupIds -or $broadGroupIds.Count -eq 0)  { if ($grpResults.AllDevices.id) { $broadGroupIds += $grpResults.AllDevices.id } }

    # -----------------------
    # Baseline activity
    # -----------------------
    $template = Get-LatestWindowsSecurityBaselineTemplate -GraphCloud $graphCloud -LogLevel $logLevel

    $result = @{
        TenantId        = $tenantId
        TemplatePicked  = @{
            id               = $template.id
            displayName       = $template.displayName
            versionInfo       = $template.versionInfo
            publishedDateTime = $template.publishedDateTime
            isDeprecated      = $template.isDeprecated
        }
        Groups          = $grpResults
        Created         = @()
        Existing        = @()
        Assignments     = @()
        DryRun          = $dryRun
        CorrelationId   = $corrFromIn
    }

    function Ensure-BaselineIntent {
        param(
            [string] $Name,
            [string] $Desc,
            [string[]] $AssignGroupIds
        )
        $existing = Find-IntentByDisplayName -GraphCloud $graphCloud -DisplayName $Name -LogLevel $logLevel
        if ($existing) {
            $result.Existing += @{ displayName=$Name; id=$existing.id; templateId=$existing.templateId }
            if (-not $dryRun -and $AssignGroupIds -and $AssignGroupIds.Count -gt 0) {
                Assign-IntentToGroups -GraphCloud $graphCloud -IntentId $existing.id -GroupIds $AssignGroupIds -LogLevel $logLevel
                $result.Assignments += @{ intentId=$existing.id; displayName=$Name; groupIds=$AssignGroupIds; existing=$true }
            }
            return $existing
        }

        if ($dryRun) {
            $result.Created += @{ displayName=$Name; id=$null; templateId=$template.id; dryRun=$true }
            return $null
        }

        $created = Create-IntentFromTemplate -GraphCloud $graphCloud -TemplateId $template.id -DisplayName $Name -Description $Desc -RoleScopeTagIds $roleScopeTagIds -LogLevel $logLevel
        $result.Created += @{ displayName=$Name; id=$created.id; templateId=$created.templateId }

        if ($AssignGroupIds -and $AssignGroupIds.Count -gt 0) {
            Assign-IntentToGroups -GraphCloud $graphCloud -IntentId $created.id -GroupIds $AssignGroupIds -LogLevel $logLevel
            $result.Assignments += @{ intentId=$created.id; displayName=$Name; groupIds=$AssignGroupIds }
        }
        return $created
    }

    $pilotDesc = "Windows Security Baseline (Level 1) - Pilot ring. Created via automation."
    $broadDesc = "Windows Security Baseline (Level 1) - Broad ring. Created via automation."

    Ensure-BaselineIntent -Name $pilotName -Desc $pilotDesc -AssignGroupIds $pilotGroupIds | Out-Null
    Ensure-BaselineIntent -Name $broadName -Desc $broadDesc -AssignGroupIds $broadGroupIds | Out-Null

    $message =
        if ($dryRun) { "DryRun enabled – groups/baselines were not created." }
        elseif (($result.Created | Where-Object { $_.id }).Count -gt 0 -or ($grpResults.Values | Where-Object { $_.existed -eq $false }).Count -gt 0) { "Baselines and/or groups created successfully." }
        else { "Everything already existed; no changes required." }

    $result.Message = $message

    Write-JsonResponse -StatusCode 200 -BodyObject $result
}
catch {
    Write-Error $_
    $msg = $_.Exception.Message
    $status = 500
    if ($msg -match 'resolve tenant' -or $msg -match 'not found') { $status = 404 }
    if ($msg -match 'required' -or $msg -match 'Empty request body') { $status = 400 }

    Write-JsonResponse -StatusCode $status -BodyObject @{
        error         = $msg
        correlationId = $correlationId
        stack         = $_.ScriptStackTrace
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}