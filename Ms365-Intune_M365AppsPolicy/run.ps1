<#
.SYNOPSIS
    Azure Functions HTTP-trigger PowerShell script to create or update an Intune "Microsoft 365 Apps for Windows"
    (officeSuiteApp) deployment that installs the latest M365 Apps and removes older Office versions if found.

.DESCRIPTION
    - Resolves tenant (GUID or domain) and selects Global/USGov Graph cloud
    - Connects to Microsoft Graph using app-only credentials (ClientId/ClientSecret)
    - Checks if an officeSuiteApp with the provided displayName exists
    - Creates or updates the app with desired properties (architecture, channel, excluded apps, SCA, uninstall older)
    - Optionally assigns the app as Required to an Entra ID group (merging with existing assignments)
    - Supports DryRun and emits detailed error messages from Graph on failure
    - RecreateIfImmutable flag — on PATCH immutability errors, optionally delete/recreate and re-apply existing assignments

.PREREQUISITES
    Graph Application Permissions (with Admin Consent):
      - DeviceManagementApps.ReadWrite.All
    Intune RBAC:
      - Assign the app’s service principal to an Intune role with App management rights
      - Ensure Scope (groups) covers target groups you’ll assign to
    Function App Settings:
      - Ms365_AuthAppId    = <Client ID GUID>
      - Ms365_AuthSecretId = <Client Secret value>
    PowerShell Modules:
      - requirements.psd1: Microsoft.Graph = '2.*'

.HTTP
    POST /api/intune/m365apps
#>

using namespace System.Net

param($Request, $TriggerMetadata)

# -------------------------
# Helpers
# -------------------------

function Write-JsonResponse {
    <#
    .SYNOPSIS
        Emits a standardized JSON HTTP response back to the Azure Functions runtime.
    #>
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
    <#
    .SYNOPSIS
        Minimal log helper with Info/Debug filtering.
    #>
    param(
        [ValidateSet('Info','Debug')] [string] $Level = 'Info',
        [Parameter(Mandatory)] [string] $Message,
        [string] $ConfiguredLevel = 'Info'
    )
    if ($Level -eq 'Debug' -and $ConfiguredLevel -ne 'Debug') { return }
    Write-Host "[$Level] $Message"
}

function Get-LoginHost {
    <#
    .SYNOPSIS
        Returns the correct AAD login host for the selected cloud.
    #>
    param([string] $GraphCloud = 'Global')
    switch ($GraphCloud.ToLower()) {
        'usgov' { 'login.microsoftonline.us' }
        default { 'login.microsoftonline.com' }
    }
}

function Get-GraphHost {
    <#
    .SYNOPSIS
        Returns the correct Microsoft Graph host for the selected cloud.
    #>
    param([string] $GraphCloud = 'Global')
    switch ($GraphCloud.ToLower()) {
        'usgov' { 'graph.microsoft.us' }
        # Optional alias support if you ever pass "gcc"
        'gcc'   { 'graph.microsoft.us' }
        default { 'graph.microsoft.com' }
    }
}

function Resolve-TenantId {
    <#
    .SYNOPSIS
        Resolves a tenant GUID from an input domain (or passes through a GUID).
    #>
    param(
        [Parameter(Mandatory)][string] $CustomerTenant,
        [string] $GraphCloud = 'Global',
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

function Get-GraphErrorText {
    <#
    .SYNOPSIS
        Extracts the underlying Microsoft Graph error message from a PowerShell ErrorRecord.
    #>
    param([Parameter(Mandatory=$true)] $ErrorRecord)
    try {
        if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            $json = $ErrorRecord.ErrorDetails.Message
            $obj  = $json | ConvertFrom-Json -ErrorAction Stop
            if ($obj.error.message) { return $obj.error.message }
        }
    } catch {}
    return $ErrorRecord.Exception.Message
}

function ConvertTo-Hashtable {
    <#
    .SYNOPSIS
        Converts PSCustomObject/objects into a Hashtable (one level), for safe downstream use.
    .NOTES
        Used especially for ExcludedApps which often deserializes as PSCustomObject.
    #>
    param([Parameter(ValueFromPipeline=$true)] $InputObject)
    if ($null -eq $InputObject) { return @{} }
    if ($InputObject -is [hashtable]) { return $InputObject }
    if ($InputObject -is [pscustomobject]) {
        $ht = @{}
        foreach ($p in $InputObject.PSObject.Properties) {
            $ht[$p.Name] = $p.Value
        }
        return $ht
    }
    try {
        return ($InputObject | ConvertTo-Json -Depth 10 | ConvertFrom-Json -AsHashtable)
    } catch {
        return @{}
    }
}

# -------------------------
# Intune app (Microsoft 365 Apps) utilities
# -------------------------

function Get-ExistingOfficeSuiteAppByName {
    <#
    .SYNOPSIS
        Returns an existing OfficeSuiteApp (Microsoft 365 Apps for Windows) by displayName.
    #>
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
    <#
    .SYNOPSIS
        Builds a compliant hashtable body for creating/updating an officeSuiteApp via Graph.
    .PARAMETER ExcludedApps
        Supported keys: access, excel, oneDrive, oneNote, outlook, powerPoint, publisher, teams, word, project, visio, skypeForBusiness
    #>
    param(
        [Parameter(Mandatory)][string] $AppName,
        [Parameter(Mandatory)][string] $Description,
        [Parameter(Mandatory)][ValidateSet('x86','x64')] [string] $Architecture,
        [Parameter(Mandatory)][ValidateSet('current','monthlyEnterprise','semiAnnual','semiAnnualPreview')] [string] $UpdateChannel,
        [Parameter(Mandatory)][bool] $SharedComputerActivation,
        [Parameter(Mandatory)][bool] $AutoAcceptEula,
        [Parameter(Mandatory)][ValidateSet('full','none')] [string] $InstallProgressDisplayLevel,
        [Parameter(Mandatory)][bool] $UninstallOlderOffice,
        [Parameter(Mandatory)][hashtable] $ExcludedApps
    )

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
        officePlatformArchitecture = $Architecture
        updateChannel              = $UpdateChannel
        excludedApps               = $cleanExcluded
        useSharedComputerActivation          = $SharedComputerActivation
        autoAcceptEula                       = $AutoAcceptEula
        installProgressDisplayLevel          = $InstallProgressDisplayLevel
        shouldUninstallOlderVersionsOfOffice = $UninstallOlderOffice
        # Optional: officeConfigurationXml = "<Configuration>...</Configuration>"
    }
}

function New-OrUpdate-OfficeSuiteApp {
    <#
    .SYNOPSIS
        Creates a new officeSuiteApp or updates an existing one by displayName (idempotent by name).
    .NOTES
        Avoid PATCHing immutable fields (e.g., officePlatformArchitecture). If needed, re-create the app (see RecreateIfImmutable).
    #>
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][hashtable] $DesiredBody,
        [switch] $DryRun,
        [switch] $RecreateIfImmutable,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    $existing = Get-ExistingOfficeSuiteAppByName -GraphHost $GraphHost -AppName $DesiredBody.displayName -LogLevel $LogLevel

    if ($existing) {
        $id = $existing.id

        # PATCH: exclude architecture (commonly immutable). Keep updateChannel, but remove if your tenant 400s on it.
        $patchMap = @{
            description                         = $DesiredBody.description
            # officePlatformArchitecture        = $DesiredBody.officePlatformArchitecture   # often immutable
            updateChannel                       = $DesiredBody.updateChannel                 # remove if 400 persists
            excludedApps                        = $DesiredBody.excludedApps
            useSharedComputerActivation         = $DesiredBody.useSharedComputerActivation
            autoAcceptEula                      = $DesiredBody.autoAcceptEula
            installProgressDisplayLevel         = $DesiredBody.installProgressDisplayLevel
            shouldUninstallOlderVersionsOfOffice= $DesiredBody.shouldUninstallOlderVersionsOfOffice
        }
        $patch = $patchMap | ConvertTo-Json -Depth 6

        Write-Log -Level Debug -Message "PATCH https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$id`n$patch" -ConfiguredLevel $LogLevel
        if (-not $DryRun) {
            try {
                Invoke-MgGraphRequest -Method PATCH -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$id" -Body $patch -ContentType "application/json" -ErrorAction Stop | Out-Null
                $updated = Invoke-MgGraphRequest -Method GET -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$id" -ErrorAction Stop
                return @{ action="updated"; id=$id; app=$updated }
            } catch {
                $why = Get-GraphErrorText -ErrorRecord $_
                Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "PATCH failed: $why"

                if (-not $RecreateIfImmutable) {
                    throw "PATCH officeSuiteApp failed: $why"
                }

                # Attempt delete + recreate, preserving assignments
                Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "RecreateIfImmutable = true. Attempting delete + recreate with preserved assignments."

                # Read existing assignments
                $oldAssignments = @()
                try {
                    $oldAssignments = Get-MobileAppAssignments -GraphHost $GraphHost -AppId $id -LogLevel $LogLevel
                } catch {
                    Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "Could not read old assignments before delete: $(Get-GraphErrorText -ErrorRecord $_)"
                    $oldAssignments = @()
                }

                # DELETE old app
                try {
                    Invoke-MgGraphRequest -Method DELETE -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$id" -ErrorAction Stop | Out-Null
                } catch {
                    $dwhy = Get-GraphErrorText -ErrorRecord $_
                    throw "DELETE old officeSuiteApp ($id) failed: $dwhy"
                }

                # POST new app
                $postBody = $DesiredBody | ConvertTo-Json -Depth 8
                $created = $null
                try {
                    $created = Invoke-MgGraphRequest -Method POST -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps" -Body $postBody -ContentType "application/json" -ErrorAction Stop
                } catch {
                    $pwhy = Get-GraphErrorText -ErrorRecord $_
                    throw "POST new officeSuiteApp after delete failed: $pwhy"
                }

                # Re-apply old assignments to the new app (best-effort)
                $preservedCount = 0
                if ($created -and $created.id -and $oldAssignments -and $oldAssignments.Count -gt 0) {
                    $assignmentsToSend = @()
                    foreach ($a in $oldAssignments) {
                        $assignmentsToSend += @{
                            "@odata.type" = "#microsoft.graph.mobileAppAssignment"
                            intent        = $a.intent
                            target        = $a.target
                            settings      = $a.settings
                        }
                    }
                    $assignBody = @{ assignments = $assignmentsToSend } | ConvertTo-Json -Depth 10
                    try {
                        Invoke-MgGraphRequest -Method POST -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps/$($created.id)/assign" -Body $assignBody -ContentType "application/json" -ErrorAction Stop | Out-Null
                        $preservedCount = $oldAssignments.Count
                    } catch {
                        Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "Re-assign to new app failed (continuing): $(Get-GraphErrorText -ErrorRecord $_)"
                    }
                }

                return @{
                    action               = "recreated"
                    id                   = $created.id
                    app                  = $created
                    recreatedFrom        = $id
                    preservedAssignments = $preservedCount
                }
            }
        }
        else {
            return @{ action="would_update"; id=$id; app=$existing }
        }
    }
    else {
        $postBody = $DesiredBody | ConvertTo-Json -Depth 8
        Write-Log -Level Debug -Message "POST https://$GraphHost/v1.0/deviceAppManagement/mobileApps`n$postBody" -ConfiguredLevel $LogLevel

        if (-not $DryRun) {
            try {
                $created = Invoke-MgGraphRequest -Method POST -Uri "https://$GraphHost/v1.0/deviceAppManagement/mobileApps" -Body $postBody -ContentType "application/json" -ErrorAction Stop
                return @{ action="created"; id=$created.id; app=$created }
            } catch {
                $why = Get-GraphErrorText -ErrorRecord $_
                throw "POST officeSuiteApp failed: $why"
            }
        }
        else {
            return @{ action="would_create"; id=$null; app=$null }
        }
    }
}

function Get-MobileAppAssignments {
    <#
    .SYNOPSIS
        Returns existing assignments for a given mobile app.
    #>
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
    <#
    .SYNOPSIS
        Assigns the app to an Entra ID group with 'required' intent, merging with existing assignments.
    #>
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

    $newAssignment = @{
        "@odata.type" = "#microsoft.graph.mobileAppAssignment"
        intent        = "required"
        target        = @{
            "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
            groupId       = $GroupId
        }
    }

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
        try {
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
        } catch {
            $why = Get-GraphErrorText -ErrorRecord $_
            throw "ASSIGN officeSuiteApp failed: $why"
        }
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
    # -------- Parse request body (robust for Portal, cURL, SDK) --------
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
            try {
                # Prefer hashtable in PS 7+
                $payload = $rawJson | ConvertFrom-Json -AsHashtable -ErrorAction Stop
            } catch {
                # Fallback to PSCustomObject then normalize case-by-case
                $payload = $rawJson | ConvertFrom-Json -ErrorAction Stop
            }
        }
    }

    if (-not $payload) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error = "Empty request body."; correlationId = $correlationId }
        return
    }

    # -------- Inputs & defaults --------
    # Safely pull values from either Hashtable or PSCustomObject
    $customerTenant = if ($payload.CustomerTenant) { [string]$payload.CustomerTenant } elseif ($payload['CustomerTenant']) { [string]$payload['CustomerTenant'] } else { $null }
    if (-not $customerTenant) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error = "CustomerTenant is required (GUID or domain)."; correlationId = $correlationId }
        return
    }

    # Normalize GraphCloud; ensure it never yields an empty host
    $graphCloud = if ($payload.GraphCloud) { ([string]$payload.GraphCloud).Trim() } elseif ($payload['GraphCloud']) { ([string]$payload['GraphCloud']).Trim() } else { 'Global' }
    if ([string]::IsNullOrWhiteSpace($graphCloud)) { $graphCloud = 'Global' }
    switch ($graphCloud.ToLower()) {
        'usgov' { $graphCloud = 'USGov' }
        'gcc'   { $graphCloud = 'USGov' }  # optional alias
        default { $graphCloud = 'Global' }
    }

    $logLevel   = if ($payload.LogLevel)   { [string]$payload.LogLevel }   elseif ($payload['LogLevel']) { [string]$payload['LogLevel'] } else { 'Info' }
    $corrFromIn = if ($payload.CorrelationId) { [string]$payload.CorrelationId } elseif ($payload['CorrelationId']) { [string]$payload['CorrelationId'] } else { $correlationId }
    $dryRun     = if ($null -ne $payload.DryRun) { [bool]$payload.DryRun } elseif ($payload.ContainsKey('DryRun')) { [bool]$payload['DryRun'] } else { $false }
    $recreateIfImmutable = if ($null -ne $payload.RecreateIfImmutable) { [bool]$payload.RecreateIfImmutable } elseif ($payload.ContainsKey('RecreateIfImmutable')) { [bool]$payload['RecreateIfImmutable'] } else { $false }

    $appName        = if ($payload.AppName) { [string]$payload.AppName } elseif ($payload['AppName']) { [string]$payload['AppName'] } else { 'Microsoft 365 Apps for Windows (Latest, Uninstall older)' }
    $description    = if ($payload.Description) { [string]$payload.Description } elseif ($payload['Description']) { [string]$payload['Description'] } else { 'Deploys latest Microsoft 365 Apps (Office) for Windows and removes older versions if found.' }
    $architecture   = if ($payload.Architecture) { [string]$payload.Architecture } elseif ($payload['Architecture']) { [string]$payload['Architecture'] } else { 'x64' }  # x86 | x64
    $updateChannel  = if ($payload.UpdateChannel) { [string]$payload.UpdateChannel } elseif ($payload['UpdateChannel']) { [string]$payload['UpdateChannel'] } else { 'current' } # current | monthlyEnterprise | semiAnnual | semiAnnualPreview
    $sca            = if ($null -ne $payload.SharedComputerActivation) { [bool]$payload.SharedComputerActivation } elseif ($payload.ContainsKey('SharedComputerActivation')) { [bool]$payload['SharedComputerActivation'] } else { $false }
    $autoEula       = if ($null -ne $payload.AutoAcceptEula) { [bool]$payload.AutoAcceptEula } elseif ($payload.ContainsKey('AutoAcceptEula')) { [bool]$payload['AutoAcceptEula'] } else { $true }
    $displayLevel   = if ($payload.InstallProgressDisplayLevel) { [string]$payload.InstallProgressDisplayLevel } elseif ($payload['InstallProgressDisplayLevel']) { [string]$payload['InstallProgressDisplayLevel'] } else { 'full' } # full | none
    $uninstallOlder = if ($null -ne $payload.UninstallOlderOffice) { [bool]$payload.UninstallOlderOffice } elseif ($payload.ContainsKey('UninstallOlderOffice')) { [bool]$payload['UninstallOlderOffice'] } else { $true }

    # Normalize ExcludedApps to an actual hashtable
    $excludedAppsInput = if ($payload.ExcludedApps) { $payload.ExcludedApps } elseif ($payload['ExcludedApps']) { $payload['ExcludedApps'] } else { $null }
    $excludedApps = if ($excludedAppsInput) { ConvertTo-Hashtable $excludedAppsInput } else { @{} }

    $groupId        = if ($payload.GroupId) { [string]$payload.GroupId } elseif ($payload['GroupId']) { [string]$payload['GroupId'] } else { $null }

    Write-Log -Level Debug -ConfiguredLevel $logLevel -Message ("Inputs: " + (@{
        CustomerTenant       = $customerTenant
        GraphCloud           = $graphCloud
        AppName              = $appName
        Architecture         = $architecture
        UpdateChannel        = $updateChannel
        SCA                  = $sca
        AutoEULA             = $autoEula
        DisplayLevel         = $displayLevel
        UninstallOld         = $uninstallOlder
        ExcludedApps         = $excludedApps
        GroupId              = $groupId
        DryRun               = $dryRun
        RecreateIfImmutable  = $recreateIfImmutable
        CorrelationId        = $corrFromIn
    } | ConvertTo-Json -Depth 6))

    # -------- Input validation --------
    if ($architecture -notin @('x86','x64')) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{
            error         = "Architecture must be 'x86' or 'x64'. Received: $architecture"
            correlationId = $corrFromIn
        }
        return
    }
    if ($updateChannel -notin @('current','monthlyEnterprise','semiAnnual','semiAnnualPreview')) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{
            error         = "UpdateChannel invalid. Use one of: current, monthlyEnterprise, semiAnnual, semiAnnualPreview."
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

    Write-Log -Level Debug -Message "Connecting to Graph. TenantId=$tenantId, Environment=$envName" -ConfiguredLevel $logLevel
    Connect-MgGraph -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret -Environment $envName -NoWelcome | Out-Null

    # Optional: confirm context
    try {
        $ctx = Get-MgContext
        Write-Log -Level Info -ConfiguredLevel $logLevel -Message "Graph connected: AppId=$($ctx.ClientId); TenantId=$($ctx.TenantId); Env=$($ctx.Environment)"
    } catch {
        Write-Log -Level Info -ConfiguredLevel $logLevel -Message "Get-MgContext failed: $($_.Exception.Message)"
    }

    # -------- Probe read to surface RBAC/permission problems early --------
    $graphHost = Get-GraphHost -GraphCloud $graphCloud
    try {
        $probe = Invoke-MgGraphRequest -Method GET -Uri "https://$graphHost/v1.0/deviceAppManagement/mobileApps?`$top=1" -ErrorAction Stop
        Write-Log -Level Info -ConfiguredLevel $logLevel -Message ("Graph probe OK. Returned: " + ($probe.value.Count))
    } catch {
        $why = Get-GraphErrorText -ErrorRecord $_
        throw "Graph probe (read mobileApps) failed: $why"
    }

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

    # -------- Create or update (idempotent by displayName) --------
    $result = New-OrUpdate-OfficeSuiteApp -GraphHost $graphHost -DesiredBody $desired -DryRun:($dryRun) -RecreateIfImmutable:($recreateIfImmutable) -LogLevel $logLevel

    # -------- Optional assignment (merge with existing) --------
    $assignment = $null

    # Normalize/trim GroupId and ensure it's a non-empty string before assignment
    $normalizedGroupId = $null
    if ($groupId) {
        $normalizedGroupId = ([string]$groupId).Trim()
        if ([string]::IsNullOrWhiteSpace($normalizedGroupId)) {
            $normalizedGroupId = $null
        }
    }

    if ($normalizedGroupId) {
        # Only attempt assignment if we have an AppId (DryRun 'would_create' has no id)
        $hasAppId = ($null -ne $result) -and ($null -ne $result.id) -and (-not [string]::IsNullOrWhiteSpace([string]$result.id))

        if ($hasAppId) {
            try {
                $assignment = Assign-MobileAppRequiredToGroup `
                    -GraphHost $graphHost `
                    -AppId $result.id `
                    -GroupId $normalizedGroupId `
                    -DryRun:($dryRun) `
                    -LogLevel $logLevel
            }
            catch {
                $msg = $_.Exception.Message
                throw "Assignment step failed for GroupId '$normalizedGroupId': $msg"
            }
        }
        else {
            # Report intent for DryRun/new create without id
            $assignment = @{
                assigned      = $false
                wouldAssignTo = $normalizedGroupId
                note          = "DryRun enabled or app not yet created; no AppId available for /assign."
            }
        }
    }

    # Safe boolean for top-level Assigned field
    $assignedFlag = $false
    if ($assignment -and ($assignment -is [hashtable]) -and $assignment.ContainsKey('assigned')) {
        $assignedFlag = [bool]$assignment['assigned']
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
        Assigned      = $assignedFlag
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
    if ($msg -match 'required' -or $msg -match 'invalid' -or $msg -match 'must be' -or $msg -match 'failed:') { $status = 400 }

    Write-JsonResponse -StatusCode $status -BodyObject @{
        error         = $msg
        correlationId = $correlationId
        stack         = $_.ScriptStackTrace
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}