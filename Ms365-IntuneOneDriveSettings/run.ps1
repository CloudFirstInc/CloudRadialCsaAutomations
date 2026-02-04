using namespace System.Net

param($Request, $TriggerMetadata)

# ============================================================
# Helpers: JSON, logging, cloud endpoints, tenant & domain
# ============================================================
function Write-JsonResponse {
    param(
        [Parameter(Mandatory)] [int] $StatusCode,
        [Parameter(Mandatory)]       $BodyObject
    )
    $json = $BodyObject | ConvertTo-Json -Depth 15
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
            Write-Log -Level Debug -Message "Resolve-TenantId: '$CustomerTenant' -> tenantId '$tid'." -ConfiguredLevel $LogLevel
            return $tid
        }
        throw "Issuer did not contain a tenant GUID."
    }
    catch {
        throw "Could not resolve tenant from domain '$CustomerTenant'. Details: $($_.Exception.Message)"
    }
}

function To-NameCase {
    param([Parameter(Mandatory)][string] $Text)
    $t = $Text.ToLower()
    return (Get-Culture).TextInfo.ToTitleCase($t)
}

function Get-DomainPrefixFromDomain {
    param([Parameter(Mandatory)][string] $Domain)
    $d = $Domain.ToLower()
    $prefix = ($d -split '\.')[0]
    return $prefix
}

# ============================================================
# Graph connection (App-only) via Client Credentials
# ============================================================
function Connect-GraphAppOnly {
    param(
        [Parameter(Mandatory)][string] $TenantId,
        [Parameter(Mandatory)][string] $GraphCloud,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    $clientId     = $env:Ms365_AuthAppId
    $clientSecret = $env:Ms365_AuthSecretId
    if (-not $clientId -or -not $clientSecret) {
        throw "Missing Graph credentials. Set App Settings Ms365_AuthAppId and Ms365_AuthSecretId."
    }

    $envName = if ($GraphCloud -ieq 'USGov') { 'USGov' } else { 'Global' }
    $secure  = ConvertTo-SecureString $clientSecret -AsPlainText -Force
    $creds   = New-Object System.Management.Automation.PSCredential($clientId, $secure)

    Write-Log -Level Debug -Message "Connecting to Graph. TenantId=$TenantId, Environment=$envName" -ConfiguredLevel $LogLevel
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $creds -Environment $envName -NoWelcome | Out-Null
}

# ============================================================
# Domain helpers (find default verified domain when only GUID is provided)
# ============================================================
function Get-DefaultDomainForTenant {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    # Try default verified domain first; fall back to any verified non-initial; else first domain
    $uri = "https://$GraphHost/v1.0/domains"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

    $default = $resp.value | Where-Object { $_.isDefault -eq $true -and $_.isVerified -eq $true }
    if ($default) { return $default[0].id }

    $verified = $resp.value | Where-Object { $_.isVerified -eq $true -and $_.isInitial -ne $true }
    if ($verified) { return $verified[0].id }

    if ($resp.value.Count -gt 0) { return $resp.value[0].id }

    throw "No domains found in tenant."
}

# ============================================================
# Group helpers
# ============================================================
function Resolve-GroupIdByDisplayName {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $DisplayName,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    $encoded = [System.Web.HttpUtility]::UrlEncode("displayName eq '$DisplayName'")
    $uri = "https://$GraphHost/v1.0/groups?`$filter=$encoded&`$select=id,displayName"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    if ($resp.value.Count -gt 0) { return $resp.value[0].id }
    return $null
}

# ============================================================
# Administrative Templates (ADMX-backed) via Graph
# ============================================================
function Get-ExistingGpConfigByName {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $Name,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    $encoded = [System.Web.HttpUtility]::UrlEncode("displayName eq '$Name'")
    $uri = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations?`$filter=$encoded"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    if ($resp.value.Count -gt 0) { return $resp.value[0] }
    return $null
}

function New-OrUpdate-GpConfig {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $Name,
        [Parameter(Mandatory)][string] $Description,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    $existing = Get-ExistingGpConfigByName -GraphHost $GraphHost -Name $Name -LogLevel $LogLevel
    if ($existing) {
        $id = $existing.id
        $body = @{ description = $Description } | ConvertTo-Json -Depth 5
        $uri  = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$id"
        Write-Log -Level Debug -Message "PATCH $uri`n$body" -ConfiguredLevel $LogLevel
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
        return @{ action="updated"; id=$id }
    }
    else {
        $body = @{
            displayName = $Name
            description = $Description
        } | ConvertTo-Json -Depth 5
        $uri = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations"
        Write-Log -Level Debug -Message "POST $uri`n$body" -ConfiguredLevel $LogLevel
        $created = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop
        return @{ action="created"; id=$created.id }
    }
}

function Find-GpDefinitionByDisplayName {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $DisplayName,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    # Try exact, then contains; prefer OneDrive categories if multiple
    $encodedEq = [System.Web.HttpUtility]::UrlEncode("displayName eq '$DisplayName'")
    $uri = "https://$GraphHost/v1.0/deviceManagement/groupPolicyDefinitions?`$filter=$encodedEq"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue

    if (-not $resp.value -or $resp.value.Count -eq 0) {
        $contains = [System.Web.HttpUtility]::UrlEncode("contains(displayName,'$DisplayName')")
        $uri2 = "https://$GraphHost/v1.0/deviceManagement/groupPolicyDefinitions?`$filter=$contains"
        Write-Log -Level Debug -Message "GET $uri2" -ConfiguredLevel $LogLevel
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri2 -ErrorAction Stop
    }

    $candidates = @()
    foreach ($d in $resp.value) {
        if ($d.categoryPath -and ($d.categoryPath -like "*OneDrive*" -or $d.categoryPath -like "*Microsoft OneDrive*")) {
            $candidates += $d
        }
    }
    if ($candidates.Count -eq 0 -and $resp.value.Count -gt 0) {
        $candidates = $resp.value
    }
    if ($candidates.Count -gt 0) { return $candidates[0] }
    return $null
}

function Get-ExistingDefinitionValue {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $ConfigId,
        [Parameter(Mandatory)][string] $DefinitionId,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    $uri = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$ConfigId/definitionValues?`$filter=definitionId eq '$DefinitionId'"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    if ($resp.value.Count -gt 0) { return $resp.value[0] }
    return $null
}

function Ensure-DefinitionValue {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $ConfigId,
        [Parameter(Mandatory)][object] $Definition,
        [Parameter(Mandatory)][bool]   $Enabled,
        [hashtable] $PresentationValues,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    $defId = $Definition.id
    $existing = Get-ExistingDefinitionValue -GraphHost $GraphHost -ConfigId $ConfigId -DefinitionId $defId -LogLevel $LogLevel

    if (-not $existing) {
        $body = @{
            definition     = @{ id = $defId }
            enabled        = $Enabled
        } | ConvertTo-Json -Depth 6
        $uriCreate = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$ConfigId/definitionValues"
        Write-Log -Level Debug -Message "POST $uriCreate`n$body" -ConfiguredLevel $LogLevel
        $existing = Invoke-MgGraphRequest -Method POST -Uri $uriCreate -Body $body -ContentType "application/json" -ErrorAction Stop
    }
    else {
        if ($null -ne $Enabled -and $existing.enabled -ne $Enabled) {
            $patchBody = @{ enabled = $Enabled } | ConvertTo-Json -Depth 4
            $uriPatch  = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$ConfigId/definitionValues/$($existing.id)"
            Write-Log -Level Debug -Message "PATCH $uriPatch`n$patchBody" -ConfiguredLevel $LogLevel
            Invoke-MgGraphRequest -Method PATCH -Uri $uriPatch -Body $patchBody -ContentType "application/json" -ErrorAction Stop | Out-Null
            $existing = Invoke-MgGraphRequest -Method GET -Uri $uriPatch -ErrorAction Stop
        }
    }

    if ($PresentationValues -and $PresentationValues.Keys.Count -gt 0) {
        $uriPres = "https://$GraphHost/v1.0/deviceManagement/groupPolicyDefinitions/$defId/presentations"
        Write-Log -Level Debug -Message "GET $uriPres" -ConfiguredLevel $LogLevel
        $pres = Invoke-MgGraphRequest -Method GET -Uri $uriPres -ErrorAction Stop

        $uriPV = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$ConfigId/definitionValues/$($existing.id)/presentationValues"
        Write-Log -Level Debug -Message "GET $uriPV" -ConfiguredLevel $LogLevel
        $existingPV = Invoke-MgGraphRequest -Method GET -Uri $uriPV -ErrorAction Stop

        foreach ($k in $PresentationValues.Keys) {
            $target = $pres.value | Where-Object { $_.displayName -eq $k }
            if (-not $target) {
                Write-Log -Level Debug -Message "Presentation '$k' not found on definition '$($Definition.displayName)'. Skipping." -ConfiguredLevel $LogLevel
                continue
            }

            $val = $PresentationValues[$k]
            $existingMatch = $existingPV.value | Where-Object { $_.presentationId -eq $target.id }

            if ($target.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationText') {
                $payload = @{
                    "@odata.type" = "#microsoft.graph.groupPolicyPresentationValueText"
                    presentationValue = @{ "@odata.type" = "#microsoft.graph.groupPolicyPresentationValueText" }
                    presentationId = $target.id
                    value          = [string]$val
                }
            }
            elseif ($target.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationCheckBox') {
                $payload = @{
                    "@odata.type" = "#microsoft.graph.groupPolicyPresentationValueBoolean"
                    presentationValue = @{ "@odata.type" = "#microsoft.graph.groupPolicyPresentationValueBoolean" }
                    presentationId = $target.id
                    value          = [bool]$val
                }
            }
            elseif ($target.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationDecimalTextBox' -or
                    $target.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationComboBoxNumber') {
                $payload = @{
                    "@odata.type" = "#microsoft.graph.groupPolicyPresentationValueDecimal"
                    presentationValue = @{ "@odata.type" = "#microsoft.graph.groupPolicyPresentationValueDecimal" }
                    presentationId = $target.id
                    value          = [int]$val
                }
            }
            else {
                Write-Log -Level Debug -Message "Unhandled presentation type '$($target.'@odata.type')' on '$k'." -ConfiguredLevel $LogLevel
                continue
            }

            if ($existingMatch) {
                $pvUri = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$ConfigId/definitionValues/$($existing.id)/presentationValues/$($existingMatch.id)"
                $pvBody = $payload | ConvertTo-Json -Depth 10
                Write-Log -Level Debug -Message "PATCH $pvUri`n$pvBody" -ConfiguredLevel $LogLevel
                Invoke-MgGraphRequest -Method PATCH -Uri $pvUri -Body $pvBody -ContentType "application/json" -ErrorAction Stop | Out-Null
            }
            else {
                $pvUri = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$ConfigId/definitionValues/$($existing.id)/presentationValues"
                $pvBody = $payload | ConvertTo-Json -Depth 10
                Write-Log -Level Debug -Message "POST $pvUri`n$pvBody" -ConfiguredLevel $LogLevel
                Invoke-MgGraphRequest -Method POST -Uri $pvUri -Body $pvBody -ContentType "application/json" -ErrorAction Stop | Out-Null
            }
        }
    }

    return @{ definitionId = $defId; definitionValueId = $existing.id; enabled = $Enabled }
}

# ---- Assignments (merge friendly) ----
function Get-GpConfigAssignmentGroupIds {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $ConfigId,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    $uri = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$ConfigId/assignments"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    $ids = @()
    foreach ($a in $resp.value) {
        if ($a.target.'@odata.type' -eq '#microsoft.graph.groupAssignmentTarget' -and $a.target.groupId) {
            $ids += $a.target.groupId
        }
    }
    return ($ids | Select-Object -Unique)
}

function Set-GpConfigAssignments {
    param(
        [Parameter(Mandatory)][string] $GraphHost,
        [Parameter(Mandatory)][string] $ConfigId,
        [Parameter(Mandatory)][string[]] $GroupIds,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    $assignments = @()
    foreach ($gid in ($GroupIds | Select-Object -Unique)) {
        $assignments += @{
            target = @{
                "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                groupId       = $gid
            }
        }
    }
    $body = @{ assignments = $assignments } | ConvertTo-Json -Depth 10
    $uri  = "https://$GraphHost/v1.0/deviceManagement/groupPolicyConfigurations/$ConfigId/assign"
    Write-Log -Level Debug -Message "POST $uri`n$body" -ConfiguredLevel $LogLevel
    Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
    return @{ assigned = $true; groups = ($GroupIds | Select-Object -Unique) }
}

# ============================================================
# Main
# ============================================================
$correlationId = [guid]::NewGuid().ToString()

try {
    # -------- Parse request body --------
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
    $dryRun     = [bool]$payload.DryRun

    $policyName = if ($payload.PolicyName) { [string]$payload.PolicyName } else { 'Sharepoint and OneDrive Settings' }
    $description = if ($payload.Description) { [string]$payload.Description } else { 'OneDrive/SharePoint Administrative Templates baseline.' }

    $kfmTenantOverride = if ($payload.KfmTenantIdOverride) { [string]$payload.KfmTenantIdOverride } else { '' }

    # -------- Resolve tenant GUID & connect --------
    $tenantId = Resolve-TenantId -CustomerTenant $customerTenant -GraphCloud $graphCloud -LogLevel $logLevel
    $graphHost = Get-GraphHost -GraphCloud $graphCloud

    if ($dryRun) {
        Write-JsonResponse -StatusCode 200 -BodyObject @{
            TenantId      = $tenantId
            Action        = "dryrun"
            PolicyName    = $policyName
            CorrelationId = $correlationId
        }
        return
    }

    Connect-GraphAppOnly -TenantId $tenantId -GraphCloud $graphCloud -LogLevel $logLevel | Out-Null

    # -------- Determine domain prefix for group names --------
    $domainForPrefix = $null
    if ($customerTenant -match '\.') {
        # Input looked like a domain
        $domainForPrefix = $customerTenant
    }
    else {
        # Input was a GUID; fetch default/verified domain from Graph
        $domainForPrefix = Get-DefaultDomainForTenant -GraphHost $graphHost -LogLevel $logLevel
    }

    $rawPrefix = Get-DomainPrefixFromDomain -Domain $domainForPrefix
    $prefixNameCase = To-NameCase -Text $rawPrefix     # e.g., 'ryeny' -> 'Ryeny'

    $userGroupName = "$prefixNameCase-All-User"
    $compGroupName = "$prefixNameCase-All-Computer"

    Write-Log -Level Debug -Message "Derived group names: '$userGroupName', '$compGroupName'." -ConfiguredLevel $logLevel

    # -------- Create/Update GP config --------
    $cfg = New-OrUpdate-GpConfig -GraphHost $graphHost -Name $policyName -Description $description -LogLevel $logLevel
    $configId = $cfg.id

    # -------- Settings to enforce (Admin Templates: OneDrive) --------
    $settingsMatrix = @(
        @{ Name = 'Allow users to choose how to handle Office file sync conflicts'; Enabled = $true;  PV = @{} }
        @{ Name = 'Prevent users from syncing personal OneDrive accounts';          Enabled = $true;  PV = @{} }
        @{ Name = 'Require users to confirm large delete operations';               Enabled = $true;  PV = @{} }
        @{ Name = 'Silently sign in users to the OneDrive sync app with their Windows credentials'; Enabled = $true; PV = @{} }
        @{ Name = 'Start OneDrive automatically when signing in to Windows';       Enabled = $true;  PV = @{} }
        @{ Name = 'Use OneDrive Files On-Demand';                                  Enabled = $true;  PV = @{} }
        @{ Name = 'Warn users who are low on disk space';                          Enabled = $false; PV = @{} }
        @{
            Name    = 'Silently move Windows known folders to OneDrive';
            Enabled = $true;
            PV      = @{
                'Desktop'  = $true
                'Documents'= $true
                'Pictures' = $true
                'Show notification to users after folders have been redirected' = $false
                'Tenant ID' = $( if ([string]::IsNullOrWhiteSpace($kfmTenantOverride)) { $tenantId } else { $kfmTenantOverride } )
            }
        }
        # NOTE: "Configure team site libraries to sync automatically" -> INTENTIONALLY NOT CONFIGURED
    )

    $applied = @()
    foreach ($s in $settingsMatrix) {
        $def = Find-GpDefinitionByDisplayName -GraphHost $graphHost -DisplayName $s.Name -LogLevel $logLevel
        if (-not $def) {
            $applied += @{ name=$s.Name; status="definition_not_found" }
            continue
        }
        $res = Ensure-DefinitionValue -GraphHost $graphHost -ConfigId $configId -Definition $def -Enabled $s.Enabled -PresentationValues $s.PV -LogLevel $logLevel
        $applied += @{ name=$s.Name; status="ok"; definitionId=$res.definitionId; definitionValueId=$res.definitionValueId; enabled=$res.enabled }
    }

    # -------- Resolve both assignment groups --------
    $userGroupId = Resolve-GroupIdByDisplayName -GraphHost $graphHost -DisplayName $userGroupName -LogLevel $logLevel
    $compGroupId = Resolve-GroupIdByDisplayName -GraphHost $graphHost -DisplayName $compGroupName -LogLevel $logLevel

    $assignment = $null
    $notFound   = @()
    $toAssign   = @()

    if ($userGroupId) { $toAssign += $userGroupId } else { $notFound += $userGroupName }
    if ($compGroupId) { $toAssign += $compGroupId } else { $notFound += $compGroupName }

    if ($toAssign.Count -gt 0) {
        # Merge with existing assignments (avoid overwriting others)
        $existingIds = Get-GpConfigAssignmentGroupIds -GraphHost $graphHost -ConfigId $configId -LogLevel $logLevel
        $merged = @($existingIds + $toAssign) | Select-Object -Unique
        $assignment = Set-GpConfigAssignments -GraphHost $graphHost -ConfigId $configId -GroupIds $merged -LogLevel $logLevel
    }
    else {
        $assignment = @{ assigned=$false; message="No target groups found by derived names."; attempted = @($userGroupName, $compGroupName) }
    }

    # -------- Respond --------
    Write-JsonResponse -StatusCode 200 -BodyObject @{
        TenantId        = $tenantId
        GraphCloud      = $graphCloud
        PolicyAction    = $cfg.action
        PolicyId        = $configId
        PolicyName      = $policyName
        Settings        = $applied
        DerivedGroups   = @{
            UserGroupDisplayName = $userGroupName
            ComputerGroupDisplayName = $compGroupName
            UserGroupId = $userGroupId
            ComputerGroupId = $compGroupId
            NotFound = $notFound
        }
        Assignment      = $assignment
        CorrelationId   = $correlationId
    }
}
catch {
    Write-Error $_
    $msg = $_.Exception.Message
    $status = 500
    if ($msg -match 'resolve tenant' -or $msg -match 'not found') { $status = 404 }

    Write-JsonResponse -StatusCode $status -BodyObject @{
        error         = $msg
        correlationId = $correlationId
        stack         = $_.ScriptStackTrace
    }
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}