using namespace System.Net

param($Request, $TriggerMetadata)

# -------------------------
# Helpers
# -------------------------
function Write-JsonResponse {
    param(
        [Parameter(Mandatory)] [int] $StatusCode,
        [Parameter(Mandatory)]       $BodyObject
    )
    $json = $BodyObject | ConvertTo-Json -Depth 10
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = $StatusCode
        Headers    = @{ "Content-Type" = "application/json" }
        Body       = $json
    })
}

function Write-Log {
    param(
        [ValidateSet('Info','Debug')] [string] $Level = 'Info',
        [Parameter(Mandatory)] [string] $Message,
        [string] $ConfiguredLevel = 'Info'
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
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
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

# -------------------------
# Intune policy utilities
# -------------------------
function Get-ExistingWindowsCustomPolicyByName {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $PolicyName,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    # Filter by displayName; then check @odata.type to ensure it's windows10CustomConfiguration.
    $uri = "https://$GraphHost/v1.0/deviceManagement/deviceConfigurations`?`$filter=" + [System.Web.HttpUtility]::UrlEncode("displayName eq '$PolicyName'")
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel

    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    foreach ($item in $resp.value) {
        if ($item.'@odata.type' -eq '#microsoft.graph.windows10CustomConfiguration') {
            return $item
        }
    }
    return $null
}

function New-OrUpdate-WindowsCustomPolicy {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $PolicyName,
        [Parameter(Mandatory)][string] $Description,
        [Parameter(Mandatory)][int]    $Minutes,
        [switch] $DryRun,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    if ($Minutes -lt 1 -or $Minutes -gt 1440) {
        throw "Minutes must be between 1 and 1440. Received: $Minutes"
    }

    # OMA-URI CSP for screen lock inactivity timeout (minutes)
    $omaSetting = @{
        "@odata.type" = "#microsoft.graph.omaSettingInteger"
        displayName   = "MaxInactivityTimeDeviceLock"
        description   = "Maximum inactivity time before device lock (minutes)"
        omaUri        = "./Device/Vendor/MSFT/Policy/Config/DeviceLock/MaxInactivityTimeDeviceLock"
        value         = $Minutes
    }

    $existing = Get-ExistingWindowsCustomPolicyByName -GraphHost $GraphHost -PolicyName $PolicyName -LogLevel $LogLevel

    if ($existing) {
        $policyId = $existing.id
        $patchBody = @{
            description = $Description
            omaSettings = @($omaSetting)
        } | ConvertTo-Json -Depth 6

        Write-Log -Level Debug -Message "PATCH https://$GraphHost/v1.0/deviceManagement/deviceConfigurations/$policyId`n$patchBody" -ConfiguredLevel $LogLevel
        if (-not $DryRun) {
            Invoke-MgGraphRequest -Method PATCH -Uri "https://$GraphHost/v1.0/deviceManagement/deviceConfigurations/$policyId" -Body $patchBody -ContentType "application/json" -ErrorAction Stop | Out-Null
            # Re-read to confirm
            $updated = Invoke-MgGraphRequest -Method GET -Uri "https://$GraphHost/v1.0/deviceManagement/deviceConfigurations/$policyId" -ErrorAction Stop
            return @{ action="updated"; id=$policyId; policy=$updated }
        }
        else {
            return @{ action="would_update"; id=$policyId; policy=$existing }
        }
    }
    else {
        $postBody = @{
            "@odata.type" = "#microsoft.graph.windows10CustomConfiguration"
            displayName   = $PolicyName
            description   = $Description
            omaSettings   = @($omaSetting)
        } | ConvertTo-Json -Depth 6

        Write-Log -Level Debug -Message "POST https://$GraphHost/v1.0/deviceManagement/deviceConfigurations`n$postBody" -ConfiguredLevel $LogLevel
        if (-not $DryRun) {
            $created = Invoke-MgGraphRequest -Method POST -Uri "https://$GraphHost/v1.0/deviceManagement/deviceConfigurations" -Body $postBody -ContentType "application/json" -ErrorAction Stop
            return @{ action="created"; id=$created.id; policy=$created }
        }
        else {
            return @{ action="would_create"; id=$null; policy=$null }
        }
    }
}

function Assign-DeviceConfigurationToGroup {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $PolicyId,
        [Parameter(Mandatory)][string] $GroupId,
        [switch] $DryRun,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    if (-not ($GroupId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')) {
        throw "GroupId must be a GUID."
    }

    $assignBody = @{
        assignments = @(
            @{
                "@odata.type" = "#microsoft.graph.deviceConfigurationAssignment"
                target        = @{
                    "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                    groupId       = $GroupId
                }
            }
        )
    } | ConvertTo-Json -Depth 6

    $uri = "https://$GraphHost/v1.0/deviceManagement/deviceConfigurations/$PolicyId/assign"
    Write-Log -Level Debug -Message "POST $uri`n$assignBody" -ConfiguredLevel $LogLevel

    if (-not $DryRun) {
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $assignBody -ContentType "application/json" -ErrorAction Stop | Out-Null
        return @{ assigned=$true; groupId=$GroupId }
    }
    else {
        return @{ assigned=$false; wouldAssignTo=$GroupId }
    }
}

# -------------------------
# Main
# -------------------------
$correlationId = [guid]::NewGuid().ToString()

try {
    # -------- Parse request body (robust) --------
    $payload   = $null
    $rawJson   = $null

    if ($Request.PSObject.Properties.Name -contains 'RawBody' -and $Request.RawBody) {
        $rawJson = [string]$Request.RawBody
    }
    elseif ($Request.Body -is [string]) {
        $rawJson = $Request.Body
    }
    elseif ($Request.Body -is [System.IO.Stream]) {
        $reader  = New-Object IO.StreamReader($Request.Body)
        $rawJson = $reader.ReadToEnd()
    }
    elseif ($Request.Body -is [System.Collections.IDictionary]) {
        $payload = [hashtable]$Request.Body
    }

    if (-not $payload) {
        if (-not [string]::IsNullOrWhiteSpace($rawJson)) {
            $payload = $rawJson | ConvertFrom-Json -ErrorAction Stop
        }
    }

    if (-not $payload) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error = "Empty request body."; correlationId = $correlationId }
        return
    }

    # -------- Inputs & defaults --------
    $customerTenant = $payload.CustomerTenant
    if (-not $customerTenant) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error = "CustomerTenant is required (GUID or domain)."; correlationId = $correlationId }
        return
    }

    $graphCloud = if ($payload.GraphCloud) { [string]$payload.GraphCloud } else { 'Global' }   # Global | USGov
    $logLevel   = if ($payload.LogLevel)   { [string]$payload.LogLevel }   else { 'Info' }     # Info | Debug
    $corrFromIn = if ($payload.CorrelationId) { [string]$payload.CorrelationId } else { $correlationId }
    $dryRun     = [bool]$payload.DryRun

    $policyName = if ($payload.PolicyName) { [string]$payload.PolicyName } else { 'Screen lock after 20 minutes (CSP DeviceLock)' }
    $description = if ($payload.Description) { [string]$payload.Description } else { 'Sets DeviceLock/MaxInactivityTimeDeviceLock to 20 minutes via OMA-URI.' }
    $minutes = if ($payload.Minutes) { [int]$payload.Minutes } else { 20 }
    $groupId = $payload.GroupId  # optional Entra ID group GUID for assignment

    Write-Log -Level Debug -ConfiguredLevel $logLevel -Message ("Inputs: " + (@{
        CustomerTenant = $customerTenant
        GraphCloud     = $graphCloud
        PolicyName     = $policyName
        Minutes        = $minutes
        GroupId        = $groupId
        DryRun         = $dryRun
        CorrelationId  = $corrFromIn
    } | ConvertTo-Json -Depth 5))

    if ($minutes -lt 1 -or $minutes -gt 1440) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{
            error         = "Minutes must be between 1 and 1440. Received: $minutes"
            correlationId = $corrFromIn
        }
        return
    }

    # -------- Resolve tenant GUID --------
    $tenantId = Resolve-TenantId -CustomerTenant $customerTenant -GraphCloud $graphCloud -LogLevel $logLevel

    # -------- Credentials from App Settings --------
    $clientId     = $env:Ms365_AuthAppId
    $clientSecret = $env:Ms365_AuthSecretId
    if (-not $clientId -or -not $clientSecret) {
        Write-JsonResponse -StatusCode 500 -BodyObject @{
            error         = "Missing Graph credentials. Set App Settings Ms365_AuthAppId and Ms365_AuthSecretId."
            correlationId = $corrFromIn
        }
        return
    }

    # -------- Connect (app-only) --------
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

    $graphHost = Get-GraphHost -GraphCloud $graphCloud

    # -------- Create or Update the policy --------
    $result = New-OrUpdate-WindowsCustomPolicy -GraphHost $graphHost -PolicyName $policyName -Description $description -Minutes $minutes -DryRun:($dryRun) -LogLevel $logLevel

    # -------- Optional assignment to a group --------
    $assignment = $null
    if ($groupId) {
        if ($result.id) {
            $assignment = Assign-DeviceConfigurationToGroup -GraphHost $graphHost -PolicyId $result.id -GroupId $groupId -DryRun:($dryRun) -LogLevel $logLevel
        }
        else {
            # On DryRun and 'would_create', there is no ID yet; report intent only.
            $assignment = @{ assigned = $false; wouldAssignTo = $groupId; note = "DryRun or policy does not yet have an ID." }
        }
    }

    # -------- Respond --------
    Write-JsonResponse -StatusCode 200 -BodyObject @{
        TenantId      = $tenantId
        Action        = $result.action
        PolicyId      = $result.id
        PolicyName    = $policyName
        Minutes       = $minutes
        Assigned      = if ($assignment) { [bool]$assignment.assigned } else { $false }
        Assignment    = $assignment
        DryRun        = $dryRun
        CorrelationId = $corrFromIn
    }
}
catch {
    Write-Error $_
    $msg = $_.Exception.Message
    $status = 500
    if ($msg -match 'resolve tenant' -or $msg -match 'not found') { $status = 404 }
    if ($msg -match 'required' -or $msg -match 'must be between') { $status = 400 }

    Write-JsonResponse -StatusCode $status -BodyObject @{
        error         = $msg
        correlationId = $correlationId
        stack         = $_.ScriptStackTrace
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}