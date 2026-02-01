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
# Intune app (Microsoft 365 Apps) utilities
# -------------------------
function Get-ExistingOfficeSuiteAppByName {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $AppName,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    $filter = [System.Web.HttpUtility]::UrlEncode("displayName eq '$AppName'")
    $uri = "https://$GraphHost/v1.0/deviceAppManagement/mobileApps?`$filter=$filter"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel

    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    foreach ($item in $resp.value) {
        if ($item.'@odata.type' -eq '#microsoft.graph.officeSuiteApp') {
            return $item
        }
    }
    return $null
}

function Build-OfficeSuiteAppBody {
    param(
        [Parameter(Mandatory)][string] $AppName,
        [Parameter(Mandatory)][string] $Description,
        [Parameter(Mandatory)][ValidateSet('x86','x64')] [string] $Architecture,
        [Parameter(Mandatory)][ValidateSet('current','monthlyEnterprise','semiAnnual','semiAnnualPreview','betaChannel')] [string] $UpdateChannel,
        [Parameter(Mandatory)][bool] $SharedComputerActivation,
        [Parameter(Mandatory)][bool] $AutoAcceptEula,
        [Parameter(Mandatory)][ValidateSet('full','none')] [string] $InstallProgressDisplayLevel,
        [Parameter(Mandatory)][bool] $UninstallOlderOffice,
        [Parameter(Mandatory)][hashtable] $ExcludedApps
    )

    # Ensure excludedApps only contains expected keys (Graph will reject unknown properties)
    $allowed = @('access','excel','oneDrive','oneNote','outlook','powerPoint','publisher','teams','word','project','visio','skypeForBusiness')
    $cleanExcluded = @{}
    foreach ($k in $allowed) {
        if ($ExcludedApps.ContainsKey($k)) { $cleanExcluded[$k] = [bool]$ExcludedApps[$k] } else { $cleanExcluded[$k] = $false }
    }

    return @{
        "@odata.type" = "#microsoft.graph.officeSuiteApp"
        displayName   = $AppName
        description   = $Description
        publisher     = "Microsoft"
        isFeatured    = $false
        officePlatformArchitecture = $Architecture    # x86 | x64
        updateChannel              = $UpdateChannel   # current | monthlyEnterprise | semiAnnual | semiAnnualPreview | betaChannel
        excludedApps               = $cleanExcluded
        useSharedComputerActivation          = $SharedComputerActivation
        autoAcceptEula                       = $AutoAcceptEula
        installProgressDisplayLevel          = $InstallProgressDisplayLevel  # full | none
        shouldUninstallOlderVersionsOfOffice = $UninstallOlderOffice
        # Optional: officeConfigurationXml = "<Configuration>...</Configuration>"
    }
}

function New-OrUpdate-OfficeSuiteApp {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][hashtable] $DesiredBody,
        [switch] $DryRun,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    $existing = Get-ExistingOfficeSuiteAppByName -GraphHost $GraphHost -AppName $DesiredBody.displayName -LogLevel $LogLevel

    if ($existing) {
        $id = $existing.id
        # Build PATCH body: Only include mutable properties
        $patch = @{
            description                        = $DesiredBody.description
            officePlatformArchitecture         = $DesiredBody.officePlatformArchitecture
            updateChannel                      = $DesiredBody.updateChannel
            excludedApps                       = $DesiredBody.excludedApps
            useSharedComputerActivation        = $DesiredBody.useSharedComputerActivation
            autoAcceptEula                     = $DesiredBody.autoAcceptEula
            installProgressDisplayLevel        = $DesiredBody.installProgressDisplayLevel
            shouldUninstallOlderVersionsOfOffice = $DesiredBody.shouldUninstallOlderVersionsOfOffice
        } | ConvertTo-Json -Depth 6

        Write-Log -Level Debug -Message "PATCH https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$id`n$patch" -ConfiguredLevel $LogLevel
        if (-not $DryRun) {
            Invoke-MgGraphRequest -Method PATCH -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$id" -Body $patch -ContentType "application/json" -ErrorAction Stop | Out-Null
            $updated = Invoke-MgGraphRequest -Method GET -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$id" -ErrorAction Stop
            return @{ action="updated"; id=$id; app=$updated }
        }
        else {
            return @{ action="would_update"; id=$id; app=$existing }
        }
    }
    else {
        $postBody = $DesiredBody | ConvertTo-Json -Depth 8
        Write-Log -Level Debug -Message "POST https://$GraphHost/v1.0/deviceAppManagement/mobileApps`n$postBody" -ConfiguredLevel $LogLevel

        if (-not $DryRun) {
            $created = Invoke-MgGraphRequest -Method POST -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps" -Body $postBody -ContentType "application/json" -ErrorAction Stop
            return @{ action="created"; id=$created.id; app=$created }
        }
        else {
            return @{ action="would_create"; id=$null; app=$null }
        }
    }
}

function Get-MobileAppAssignments {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $AppId,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    $uri = "https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$AppId/assignments"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel
    (Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop).value
}

function Assign-MobileAppRequiredToGroup {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $AppId,
        [Parameter(Mandatory)][string] $GroupId,
        [switch] $DryRun,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    if (-not ($GroupId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')) {
        throw "GroupId must be a GUID."
    }

    # Build the desired assignment object
    $newAssignment = @{
        "@odata.type" = "#microsoft.graph.mobileAppAssignment"
        intent        = "required"
        target        = @{
            "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
            groupId       = $GroupId
        }
        # settings can be omitted for OfficeSuiteApp; defaults are fine for Required install
    }

    # Get existing assignments and keep them; ensure we don't duplicate the same group/intent
    $existing = Get-MobileAppAssignments -GraphHost $GraphHost -AppId $AppId -LogLevel $LogLevel
    $assignmentsToSend = @()

    $alreadyThere = $false
    foreach ($a in $existing) {
        $assignmentsToSend += @{
            "@odata.type" = "#microsoft.graph.mobileAppAssignment"
            intent        = $a.intent
            target        = $a.target
            settings      = $a.settings
        }
        if ($a.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and
            $a.target.groupId -eq $GroupId -and
            $a.intent -eq 'required') {
            $alreadyThere = $true
        }
    }

    if (-not $alreadyThere) { $assignmentsToSend += $newAssignment }

    $body = @{ assignments = $assignmentsToSend } | ConvertTo-Json -Depth 10
    $uri  = "https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$AppId/assign"

    Write-Log -Level Debug -Message "POST $uri`n$body" -ConfiguredLevel $LogLevel
    if (-not $DryRun) {
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
        return @{ assigned=$true; groupId=$GroupId; mergedExisting=$true }
    }
    else {
        return @{ assigned=$false; wouldAssignTo=$GroupId; mergedExisting=$true }
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

    $appName        = if ($payload.AppName) { [string]$payload.AppName } else { 'Microsoft 365 Apps for Windows (Latest, Uninstall older)' }
    $description    = if ($payload.Description) { [string]$payload.Description } else { 'Deploys latest Microsoft 365 Apps (Office) for Windows and removes older versions if found.' }
    $architecture   = if ($payload.Architecture) { [string]$payload.Architecture } else { 'x64' }  # x86 | x64
    $updateChannel  = if ($payload.UpdateChannel) { [string]$payload.UpdateChannel } else { 'current' } # current | monthlyEnterprise | semiAnnual | semiAnnualPreview | betaChannel
    $sca            = if ($null -ne $payload.SharedComputerActivation) { [bool]$payload.SharedComputerActivation } else { $false }
    $autoEula       = if ($null -ne $payload.AutoAcceptEula) { [bool]$payload.AutoAcceptEula } else { $true }
    $displayLevel   = if ($payload.InstallProgressDisplayLevel) { [string]$payload.InstallProgressDisplayLevel } else { 'full' } # full | none
    $uninstallOlder = if ($null -ne $payload.UninstallOlderOffice) { [bool]$payload.UninstallOlderOffice } else { $true }
    $excludedApps   = if ($payload.ExcludedApps) { [hashtable]$payload.ExcludedApps } else { @{} }

    $groupId        = $payload.GroupId  # Optional Entra ID group GUID to assign as Required

    Write-Log -Level Debug -ConfiguredLevel $logLevel -Message ("Inputs: " + (@{
        CustomerTenant = $customerTenant
        GraphCloud     = $graphCloud
        AppName        = $appName
        Architecture   = $architecture
        UpdateChannel  = $updateChannel
        SCA            = $sca
        AutoEULA       = $autoEula
        DisplayLevel   = $displayLevel
        UninstallOld   = $uninstallOlder
        ExcludedApps   = $excludedApps
        GroupId        = $groupId
        DryRun         = $dryRun
        CorrelationId  = $corrFromIn
    } | ConvertTo-Json -Depth 6))

    if ($architecture -notin @('x86','x64')) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{
            error         = "Architecture must be 'x86' or 'x64'. Received: $architecture"
            correlationId = $corrFromIn
        }
        return
    }
    if ($updateChannel -notin @('current','monthlyEnterprise','semiAnnual','semiAnnualPreview','betaChannel')) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{
            error         = "UpdateChannel invalid. Use one of: current, monthlyEnterprise, semiAnnual, semiAnnualPreview, betaChannel."
            correlationId = $corrFromIn
        }
        return
    }
    if ($displayLevel -notin @('full','none')) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{
            error         = "InstallProgressDisplayLevel must be 'full' or 'none'."
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

    # -------- Build desired OfficeSuite App body --------
    $desired = Build-OfficeSuiteAppBody `
        -AppName $appName `
        -Description $description `
        -Architecture $architecture `
        -UpdateChannel $updateChannel `
        -SharedComputerActivation $sca `
        -AutoAcceptEula $autoEula `
        -InstallProgressDisplayLevel $displayLevel `
        -UninstallOlderOffice $uninstallOlder `
        -ExcludedApps $excludedApps

    # -------- Create or update --------
    $result = New-OrUpdate-OfficeSuiteApp -GraphHost $graphHost -DesiredBody $desired -DryRun:($dryRun) -LogLevel $logLevel

    # -------- Optional assignment (merge with existing) --------
    $assignment = $null
    if ($groupId) {
        if ($result.id) {
            $assignment = Assign-MobileAppRequiredToGroup -GraphHost $graphHost -AppId $result.id -GroupId $groupId -DryRun:($dryRun) -LogLevel $logLevel
        }
        else {
            $assignment = @{ assigned=$false; wouldAssignTo=$groupId; note="DryRun or app not yet created; no AppId." }
        }
    }

    # -------- Respond --------
    Write-JsonResponse -StatusCode 200 -BodyObject @{
        TenantId      = $tenantId
        Action        = $result.action
        AppId         = $result.id
        AppName       = $appName
        Architecture  = $architecture
        UpdateChannel = $updateChannel
        UninstallOlder= $uninstallOlder
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
    if ($msg -match 'required' -or $msg -match 'invalid' -or $msg -match 'must be') { $status = 400 }

    Write-JsonResponse -StatusCode $status -BodyObject @{
        error         = $msg
        correlationId = $correlationId
        stack         = $_.ScriptStackTrace
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}