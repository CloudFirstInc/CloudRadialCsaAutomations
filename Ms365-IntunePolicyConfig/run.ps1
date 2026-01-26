
using namespace System.Net

# Azure Functions entry point
param($Request, $TriggerMetadata)

# ---------------------------
# Helper: JSON HTTP response
# ---------------------------
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

# --------------------------------------
# Helper: minimal logging with log level
# --------------------------------------
function Write-Log {
    param(
        [ValidateSet('Info','Debug')] [string] $Level = 'Info',
        [Parameter(Mandatory)] [string] $Message,
        [string] $ConfiguredLevel = 'Info'
    )
    # Only emit Debug when requested
    if ($Level -eq 'Debug' -and $ConfiguredLevel -ne 'Debug') { return }
    Write-Host "[$Level] $Message"
}

# -----------------------------------------------------------
# Helper: Cloud-specific login host for OpenID metadata fetch
# -----------------------------------------------------------
function Get-LoginHost {
    param([ValidateSet('Global','USGov')][string] $GraphCloud = 'Global')
    switch ($GraphCloud.ToLower()) {
        'usgov' { 'login.microsoftonline.us' }
        default { 'login.microsoftonline.com' }
    }
}

# ------------------------------------------------------------------
# Helper: Resolve tenant GUID from either GUID or verified domain
# ------------------------------------------------------------------
function Resolve-TenantId {
    param(
        [Parameter(Mandatory)][string] $CustomerTenant,  # GUID or domain
        [ValidateSet('Global','USGov')][string] $GraphCloud = 'Global',
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    # GUID path -> return as-is
    if ($CustomerTenant -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        Write-Log -Level Debug -Message "Resolve-TenantId: '$CustomerTenant' recognized as GUID." -ConfiguredLevel $LogLevel
        return $CustomerTenant
    }

    # Domain path -> discover via OpenID metadata
    $loginHost = Get-LoginHost -GraphCloud $GraphCloud
    $wellKnown = "https://$loginHost/$CustomerTenant/v2.0/.well-known/openid-configuration"
    Write-Log -Level Debug -Message "Resolve-TenantId: Fetching $wellKnown" -ConfiguredLevel $LogLevel

    try {
        $meta   = Invoke-RestMethod -Method GET -Uri $wellKnown -ErrorAction Stop
        $issuer = [Uri]$meta.issuer  # e.g., https://login.microsoftonline.com/<tenantId>/v2.0
        $segments = $issuer.AbsolutePath.Trim('/').Split('/')
        if ($segments.Length -ge 1 -and $segments[0] -match '^[0-9a-fA-F-]{36}$') {
            $tid = $segments[0]
            Write-Log -Level Debug -Message "Resolve-TenantId: Resolved domain '$CustomerTenant' to tenantId '$tid'." -ConfiguredLevel $LogLevel
            return $tid
        }
        throw "Issuer did not contain a tenant GUID."
    }
    catch {
        throw "Could not resolve tenant from domain '$CustomerTenant'. Details: $($_.Exception.Message)"
    }
}

# ---------------------------------------------
# Helper: Fetch current Intune settings (v1.0)
# ---------------------------------------------
function Get-CurrentComplianceSettings {
    param(
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/settings"
    Write-Log -Level Debug -Message "GET $uri" -ConfiguredLevel $LogLevel
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    return @{
        secureByDefault                      = [bool]$resp.secureByDefault
        deviceComplianceCheckinThresholdDays = [int] $resp.deviceComplianceCheckinThresholdDays
    }
}

# -----------------------------------------------------------
# Helper: PATCH new Intune tenant compliance settings (v1.0)
# -----------------------------------------------------------
function Set-ComplianceSettings {
    param(
        [Parameter(Mandatory)][bool] $SecureByDefault,
        [Parameter(Mandatory)][int]  $ValidityDays,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement"
    $body = @{
        settings = @{
            secureByDefault                      = $SecureByDefault
            deviceComplianceCheckinThresholdDays = $ValidityDays
        }
    } | ConvertTo-Json -Depth 5

    Write-Log -Level Debug -Message "PATCH $uri`n$body" -ConfiguredLevel $LogLevel
    Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
}

# -------------------------
# Main function execution
# -------------------------
$correlationId = [guid]::NewGuid().ToString()

try {
    # ----- Parse JSON body -----
    $payload = $null
    if ($Request.Body) {
        $reader  = New-Object IO.StreamReader($Request.Body)
        $rawJson = $reader.ReadToEnd()
        if ($rawJson) { $payload = $rawJson | ConvertFrom-Json }
    }

    if (-not $payload) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error = "Empty request body."; correlationId = $correlationId }
        return
    }

    # ----- Extract inputs & defaults -----
    $customerTenant = $payload.CustomerTenant
    if (-not $customerTenant) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error = "CustomerTenant is required (GUID or domain)."; correlationId = $correlationId }
        return
    }

    $graphCloud = if ($payload.GraphCloud) { [string]$payload.GraphCloud } else { 'Global' }   # Global | USGov
    $logLevel   = if ($payload.LogLevel)   { [string]$payload.LogLevel }   else { 'Info' }     # Info | Debug
    $corrFromIn = if ($payload.CorrelationId) { [string]$payload.CorrelationId } else { $correlationId }
    $dryRun     = $payload.DryRun

    $days = if ($payload.ComplianceValidityDays) { [int]$payload.ComplianceValidityDays } else { 14 }
    if ($days -lt 1 -or $days -gt 120) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{
            error         = "ComplianceValidityDays must be between 1 and 120. Received: $days"
            correlationId = $corrFromIn
        }
        return
    }

    $markNotCompliant = $true
    if ($null -ne $payload.MarkDevicesWithoutPolicyNotCompliant) {
        $markNotCompliant = [bool]$payload.MarkDevicesWithoutPolicyNotCompliant
    }

    Write-Log -Level Debug -ConfiguredLevel $logLevel -Message ("Inputs: " + (@{
        CustomerTenant = $customerTenant
        GraphCloud     = $graphCloud
        Days           = $days
        MarkNotComp    = $markNotCompliant
        DryRun         = $dryRun
        CorrelationId  = $corrFromIn
    } | ConvertTo-Json -Depth 5))

    # ----- Resolve tenant GUID -----
    $tenantId = Resolve-TenantId -CustomerTenant $customerTenant -GraphCloud $graphCloud -LogLevel $logLevel

    # ----- Read creds from App Settings -----
    $clientId     = $env:Ms365_AuthAppId
    $clientSecret = $env:Ms365_AuthSecretId
    if (-not $clientId -or -not $clientSecret) {
        Write-JsonResponse -StatusCode 500 -BodyObject @{
            error         = "Missing Graph credentials. Set App Settings Ms365_AuthAppId and Ms365_AuthSecretId."
            correlationId = $corrFromIn
        }
        return
    }

    # ----- Connect to Graph (App-only) -----
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

    # ----- Read current settings -----
    $before = Get-CurrentComplianceSettings -LogLevel $logLevel
    $after  = @{
        secureByDefault                      = [bool]$markNotCompliant
        deviceComplianceCheckinThresholdDays = [int] $days
    }

    $needsUpdate = ($before.secureByDefault -ne $after.secureByDefault) -or
                   ($before.deviceComplianceCheckinThresholdDays -ne $after.deviceComplianceCheckinThresholdDays)

    if ($dryRun -or -not $needsUpdate) {
        $message = if ($dryRun) { "DryRun enabled – no changes posted." } else { "Already aligned. No changes required." }
        Write-JsonResponse -StatusCode 200 -BodyObject @{
            TenantId      = $tenantId
            Updated       = $           Before        = $before
            After         = $after
            Message       = $message
            CorrelationId = $corrFromIn
        }
        return
    }

    # ----- Apply update -----
    Set-ComplianceSettings -SecureByDefault $after.secureByDefault -ValidityDays $after.deviceComplianceCheckinThresholdDays -LogLevel $logLevel

    # ----- Confirm and return -----
    $final = Get-CurrentComplianceSettings -LogLevel $logLevel

    Write-JsonResponse -StatusCode 200 -BodyObject @{
        TenantId      = $tenantId
        Updated       = $true
        Before        = $before
        After         = $final
        Message       = "Tenant-wide Intune compliance policy settings are now aligned."
        CorrelationId = $corrFromIn
    }
}
catch {
    Write-Error $_
    $msg = $_.Exception.Message
    # Try to detect obvious bad input / not found and set smarter status when possible
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
