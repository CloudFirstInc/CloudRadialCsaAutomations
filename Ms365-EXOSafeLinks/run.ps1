using namespace System.Net

param($Request, $TriggerMetadata)

# --- Early module probe (as requested) ---
try {
    Import-Module ExchangeOnlineManagement -MinimumVersion 3.4.0 -ErrorAction Stop
    Write-Host "[Info] ExchangeOnlineManagement module imported."
} catch {
    Write-Host "[Error] Could not import ExchangeOnlineManagement: $($_.Exception.Message)"
    throw
}

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

function Get-ExchangeEnvironmentName {
    param([ValidateSet('Global','USGov')][string] $GraphCloud = 'Global')
    switch ($GraphCloud.ToLower()) {
        'usgov' { 'O365USGovGCCHigh' }  # use O365USGovDoD for DoD tenants if needed
        default { 'O365Default' }
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
        $meta     = Invoke-RestMethod -Method GET -Uri $wellKnown -ErrorAction Stop
        $issuer   = [Uri]$meta.issuer
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

function Ensure-Module {
    param(
        [Parameter(Mandatory)][string] $Name,
        [string] $MinVersion
    )
    try {
        if ($MinVersion) {
            Import-Module $Name -MinimumVersion $MinVersion -ErrorAction Stop | Out-Null
        } else {
            Import-Module $Name -ErrorAction Stop | Out-Null
        }
    }
    catch {
        throw "Required module '$Name' (min $MinVersion) not found. Add it to requirements.psd1 at the Function App root and enable managedDependency in host.json. Example: @{ 'ExchangeOnlineManagement' = '3.4.0' }"
    }
}

function Get-CertificateFromEnv {
    <#
      Returns [System.Security.Cryptography.X509Certificates.X509Certificate2]
      Sources, in order:
        1) Ms365_AuthCertThumbprint (LocalMachine/My or CurrentUser/My)
        2) Ms365_CertBase64 + Ms365_CertPassword (PFX as base64 in App Settings)
    #>

    $thumb = $env:Ms365_AuthCertThumbprint
    $b64   = $env:Ms365_CertBase64
    $pwd   = $env:Ms365_CertPassword

    if ($thumb) {
        $stores = @('Cert:\LocalMachine\My','Cert:\CurrentUser\My')
        foreach ($s in $stores) {
            $cert = Get-ChildItem -Path $s -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -ieq $thumb }
            if ($cert) { return $cert }
        }
        throw "Certificate with thumbprint '$thumb' not found in LocalMachine/CurrentUser 'My' stores."
    }

    if ($b64) {
        if (-not $pwd) { throw "Ms365_CertBase64 is set but Ms365_CertPassword is missing." }

        try {
            # Normalize Base64 (remove whitespace/newlines)
            $cleanB64 = ($b64 -replace '\s', '')
            $bytes    = [Convert]::FromBase64String($cleanB64)

            # Detect OS (don't assign to built-in $IsWindows constant)
            $onWindows = $false
            try {
                $onWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
                    [System.Runtime.InteropServices.OSPlatform]::Windows
                )
            } catch {
                $onWindows = ($PSVersionTable.Platform -eq 'Win32NT')
            }

            # Use EphemeralKeySet to avoid writing to cert stores in sandboxed environments
            $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable `
                   -bor [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet

            # If you specifically want persisted keys on Windows, you *may* switch to MachineKeySet:
            # if ($onWindows) {
            #     $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable `
            #            -bor [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet
            # }

            # Constructor path (fixes: "immutable on this platform")
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($bytes, $pwd, $flags)

            if (-not $cert.HasPrivateKey) {
                throw "The reconstructed certificate does not contain a private key."
            }

            return $cert
        }
        catch {
            throw "Failed to construct certificate from Ms365_CertBase64. $($_.Exception.Message)"
        }
    }

    throw "No certificate source found. Set either Ms365_AuthCertThumbprint or (Ms365_CertBase64 + Ms365_CertPassword)."
}

function Connect-ExchangeAsApp {
    param(
        [Parameter(Mandatory)][string] $TenantForConnection,     # tenant GUID (from Resolve-TenantId) – for logging/trace
        [Parameter(Mandatory)][string] $OrganizationDomain,      # VERIFIED domain for EXO (-Organization), e.g. contoso.com or contoso.onmicrosoft.com
        [Parameter(Mandatory)][string] $AppId,
        [ValidateSet('Global','USGov')][string] $GraphCloud = 'Global',
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )
    Ensure-Module -Name 'ExchangeOnlineManagement' -MinVersion '3.4.0'
    $exoEnv = Get-ExchangeEnvironmentName -GraphCloud $GraphCloud
    $cert   = Get-CertificateFromEnv

    if ([string]::IsNullOrWhiteSpace($OrganizationDomain) -or
        $OrganizationDomain -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        throw "OrganizationDomain must be a verified domain name (e.g., contoso.com or contoso.onmicrosoft.com), not a GUID."
    }

    $connParams = @{
        AppId                   = $AppId
        Certificate             = $cert
        Organization            = $OrganizationDomain   # EXO expects a domain here, not GUID
        ShowBanner              = $false
        ExchangeEnvironmentName = $exoEnv
        CommandName             = @()  # lazy-load
    }

    Write-Log -Level Debug -Message "Connecting to Exchange Online (Environment=$exoEnv, Org=$OrganizationDomain, TenantId=$TenantForConnection)" -ConfiguredLevel $LogLevel
    Connect-ExchangeOnline @connParams | Out-Null
}

function Get-PolicySnapshot {
    param([string] $PolicyName)
    $p = Get-SafeLinksPolicy -Identity $PolicyName -ErrorAction SilentlyContinue
    if (-not $p) { return $null }
    return @{
        Name                     = $p.Name
        EnableSafeLinksForEmail  = [bool]$p.EnableSafeLinksForEmail
        EnableSafeLinksForOffice = [bool]$p.EnableSafeLinksForOffice
        EnableSafeLinksForTeams  = [bool]$p.EnableSafeLinksForTeams
        ScanUrls                 = [bool]$p.ScanUrls
        EnableForInternalSenders = [bool]$p.EnableForInternalSenders
        DeliverMessageAfterScan  = [bool]$p.DeliverMessageAfterScan
        AllowClickThrough        = [bool]$p.AllowClickThrough
        TrackClicks              = [bool]$p.TrackClicks
    }
}

function Get-RuleSnapshot {
    param([string] $RuleName)
    $r = Get-SafeLinksRule -Identity $RuleName -ErrorAction SilentlyContinue
    if (-not $r) { return $null }
    return @{
        Name              = $r.Name
        SafeLinksPolicy   = $r.SafeLinksPolicy
        RecipientDomainIs = @($r.RecipientDomainIs)
        Priority          = [int]$r.Priority
        Enabled           = [bool]$r.Enabled
    }
}

function Ensure-SafeLinksPolicy {
    param(
        [Parameter(Mandatory)][string] $PolicyName,
        [Parameter(Mandatory)][hashtable] $Desired,
        [switch] $DryRun,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    $before = Get-PolicySnapshot -PolicyName $PolicyName

    if (-not $before) {
        if ($DryRun) {
            Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "[DryRun] Would create Safe Links policy '$PolicyName'."
        } else {
            Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "Creating Safe Links policy '$PolicyName'..."
            New-SafeLinksPolicy -Name $PolicyName `
                -EnableSafeLinksForEmail $Desired.EnableSafeLinksForEmail `
                -EnableSafeLinksForOffice $Desired.EnableSafeLinksForOffice `
                -EnableSafeLinksForTeams $Desired.EnableSafeLinksForTeams `
                -ScanUrls $Desired.ScanUrls `
                -EnableForInternalSenders $Desired.EnableForInternalSenders `
                -DeliverMessageAfterScan $Desired.DeliverMessageAfterScan `
                -AllowClickThrough $Desired.AllowClickThrough `
                -TrackClicks $Desired.TrackClicks `
                -ErrorAction Stop | Out-Null
        }
        $after = $DryRun ? $Desired : (Get-PolicySnapshot -PolicyName $PolicyName)
        return @{ Created = $true; Updated = $false; Before = $before; After = $after }
    }

    $needsUpdate = $false
    foreach ($k in $Desired.Keys) {
        if ($before[$k] -ne $Desired[$k]) { $needsUpdate = $true; break }
    }

    if (-not $needsUpdate) {
        return @{ Created = $false; Updated = $false; Before = $before; After = $before }
    }

    if ($DryRun) {
        Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "[DryRun] Would update Safe Links policy '$PolicyName'."
        $after = $Desired
    } else {
        Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "Updating Safe Links policy '$PolicyName'..."
        Set-SafeLinksPolicy -Identity $PolicyName `
            -EnableSafeLinksForEmail $Desired.EnableSafeLinksForEmail `
            -EnableSafeLinksForOffice $Desired.EnableSafeLinksForOffice `
            -EnableSafeLinksForTeams $Desired.EnableSafeLinksForTeams `
            -ScanUrls $Desired.ScanUrls `
            -EnableForInternalSenders $Desired.EnableForInternalSenders `
            -DeliverMessageAfterScan $Desired.DeliverMessageAfterScan `
            -AllowClickThrough $Desired.AllowClickThrough `
            -TrackClicks $Desired.TrackClicks `
            -ErrorAction Stop | Out-Null
        $after = Get-PolicySnapshot -PolicyName $PolicyName
    }

    return @{ Created = $false; Updated = $true; Before = $before; After = $after }
}

function Ensure-SafeLinksRule {
    param(
        [Parameter(Mandatory)][string] $RuleName,
        [Parameter(Mandatory)][string] $PolicyName,
        [Parameter(Mandatory)][string[]] $Domains,
        [int] $Priority = 0,   # default; not Mandatory
        [switch] $DryRun,
        [ValidateSet('Info','Debug')][string] $LogLevel = 'Info'
    )

    $before = Get-RuleSnapshot -RuleName $RuleName
    if (-not $before) {
        if ($DryRun) {
            Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "[DryRun] Would create Safe Links rule '$RuleName' -> policy '$PolicyName' (Priority $Priority)."
        } else {
            Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "Creating Safe Links rule '$RuleName' -> policy '$PolicyName' (Priority $Priority)..."
            New-SafeLinksRule -Name $RuleName `
                -SafeLinksPolicy $PolicyName `
                -RecipientDomainIs ([string[]]$Domains) `
                -Priority $Priority `
                -ErrorAction Stop | Out-Null
        }
        $after = $DryRun ? @{ Name=$RuleName; SafeLinksPolicy=$PolicyName; RecipientDomainIs=$Domains; Priority=$Priority; Enabled=$true } : (Get-RuleSnapshot -RuleName $RuleName)
        return @{ Created = $true; Updated = $false; Before = $before; After = $after }
    }

    $beforeDomains = @($before.RecipientDomainIs) | Sort-Object -Unique
    $newDomains    = @($Domains) | Sort-Object -Unique

    $needsUpdate = $false
    if ($before.SafeLinksPolicy -ne $PolicyName) { $needsUpdate = $true }
    elseif ($beforeDomains -join ',' -ne $newDomains -join ',') { $needsUpdate = $true }
    elseif ($before.Priority -ne $Priority) { $needsUpdate = $true }

    if (-not $needsUpdate) {
        return @{ Created = $false; Updated = $false; Before = $before; After = $before }
    }

    if ($DryRun) {
        Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "[DryRun] Would update Safe Links rule '$RuleName'."
        $after = @{ Name=$RuleName; SafeLinksPolicy=$PolicyName; RecipientDomainIs=$newDomains; Priority=$Priority; Enabled=$before.Enabled }
        return @{ Created = $false; Updated = $true; Before = $before; After = $after }
    }

    Write-Log -Level Info -ConfiguredLevel $LogLevel -Message "Updating Safe Links rule '$RuleName'..."
    try {
        Set-SafeLinksRule -Identity $RuleName -SafeLinksPolicy $PolicyName -RecipientDomainIs ([string[]]$newDomains) -Priority $Priority -ErrorAction Stop | Out-Null
    }
    catch {
        # If priority 0 is already reserved elsewhere, retry without touching priority
        Write-Log -Level Debug -ConfiguredLevel $LogLevel -Message "Priority update failed: $($_.Exception.Message). Retrying without priority change..."
        Set-SafeLinksRule -Identity $RuleName -SafeLinksPolicy $PolicyName -RecipientDomainIs ([string[]]$newDomains) -ErrorAction Stop | Out-Null
    }
    $after = Get-RuleSnapshot -RuleName $RuleName
    return @{ Created = $false; Updated = $true; Before = $before; After = $after }
}

# -------------------------
# Main
# -------------------------
$correlationId = [guid]::NewGuid().ToString()

try {
    # Parse request body robustly (Portal, cURL, SDK)
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

    # Inputs & defaults
    $customerTenant = [string]$payload.CustomerTenant
    if (-not $customerTenant) {
        Write-JsonResponse -StatusCode 400 -BodyObject @{ error = "CustomerTenant is required (GUID or domain)."; correlationId = $correlationId }
        return
    }

    $graphCloud = if ($payload.GraphCloud) { [string]$payload.GraphCloud } else { 'Global' }   # Global | USGov
    $logLevel   = if ($payload.LogLevel)   { [string]$payload.LogLevel }   else { 'Info' }     # Info | Debug
    $corrFromIn = if ($payload.CorrelationId) { [string]$payload.CorrelationId } else { $correlationId }
    $dryRun     = [bool]$payload.DryRun

    $policyName = if ($payload.PolicyName) { [string]$payload.PolicyName } else { 'CF Tenant-Wide Safe Links' }
    $ruleName   = $policyName
    $priority   = if ($payload.Priority -ne $null) { [int]$payload.Priority } else { 0 }

    # Decide OrganizationDomain for EXO (-Organization requires a domain)
    $organizationDomain = $null
    if ($customerTenant -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        # CustomerTenant looks like a domain (e.g., ryeny.gov or contoso.onmicrosoft.com)
        $organizationDomain = $customerTenant
    } else {
        # GUID provided; allow explicit override in payload
        if ($payload.OrganizationDomain) {
            $organizationDomain = [string]$payload.OrganizationDomain
        } else {
            throw "CustomerTenant is a GUID, but Exchange Online requires a domain for -Organization. Add 'OrganizationDomain' to the payload (e.g., 'contoso.onmicrosoft.com' or a verified domain)."
        }
    }

    # Desired settings per your spec
    $desired = @{
        EnableSafeLinksForEmail   = $true
        EnableSafeLinksForOffice  = $true
        EnableSafeLinksForTeams   = $true
        ScanUrls                  = $true
        EnableForInternalSenders  = $true
        DeliverMessageAfterScan   = $true
        AllowClickThrough         = $false
        TrackClicks               = $true
    }

    Write-Log -Level Debug -ConfiguredLevel $logLevel -Message ("Inputs: " + (@{
        CustomerTenant = $customerTenant
        OrganizationDomain = $organizationDomain
        GraphCloud     = $graphCloud
        PolicyName     = $policyName
        Priority       = $priority
        DryRun         = $dryRun
        CorrelationId  = $corrFromIn
    } | ConvertTo-Json -Depth 5))

    # Resolve tenant GUID for traceability/logging
    $tenantId = Resolve-TenantId -CustomerTenant $customerTenant -GraphCloud $graphCloud -LogLevel $logLevel

    # Credentials from App Settings
    $clientId = $env:Ms365_AuthAppId
    if (-not $clientId) {
        Write-JsonResponse -StatusCode 500 -BodyObject @{
            error         = "Missing AppId. Set App Setting 'Ms365_AuthAppId'. Also provide cert via 'Ms365_AuthCertThumbprint' OR ('Ms365_CertBase64' + 'Ms365_CertPassword')."
            correlationId = $corrFromIn
        }
        return
    }

    # Connect (Exchange app-only) with a domain for -Organization
    Connect-ExchangeAsApp `
        -TenantForConnection $tenantId `
        -OrganizationDomain  $organizationDomain `
        -AppId               $clientId `
        -GraphCloud          $graphCloud `
        -LogLevel            $logLevel

    # Gather current state - accepted domains
    [string[]]$acceptedDomains = @()
    try {
        $acceptedDomains = Get-AcceptedDomain -ErrorAction Stop | ForEach-Object {
            if ($_.PSObject.Properties.Name -contains 'DomainName' -and $_.DomainName) { $_.DomainName } else { $_.Name }
        }
        $acceptedDomains = $acceptedDomains | Sort-Object -Unique
    }
    catch {
        throw "Failed to read accepted domains. Ensure the app has Exchange RBAC permissions. $($_.Exception.Message)"
    }

    if (-not $acceptedDomains -or $acceptedDomains.Count -eq 0) {
        throw "No accepted domains found in the target tenant."
    }

    # Ensure policy
    $policyResult = Ensure-SafeLinksPolicy -PolicyName $policyName -Desired $desired -DryRun:($dryRun) -LogLevel $logLevel

    # Ensure rule (scope = all accepted domains)
    $ruleResult = Ensure-SafeLinksRule -RuleName $ruleName -PolicyName $policyName -Domains $acceptedDomains -Priority $priority -DryRun:($dryRun) -LogLevel $logLevel

    # Summarize
    $updated = ($policyResult.Created -or $policyResult.Updated -or $ruleResult.Created -or $ruleResult.Updated)
    $msg = if ($dryRun) {
        "DryRun enabled – no changes posted."
    } elseif ($updated) {
        "Safe Links policy/rule are now aligned."
    } else {
        "Already aligned. No changes required."
    }

    Write-JsonResponse -StatusCode 200 -BodyObject @{
        TenantId        = $tenantId
        Organization    = $organizationDomain
        PolicyName      = $policyName
        RuleName        = $ruleName
        DomainsCount    = $acceptedDomains.Count
        PolicyResult    = $policyResult
        RuleResult      = $ruleResult
        Updated         = $updated
        Message         = $msg
        CorrelationId   = $corrFromIn
    }
}
catch {
    Write-Error $_
   tatus = 500
    if ($msg -match 'resolve tenant' -or $msg -match 'not found') { $status = 404 }
    if ($msg -match 'Missing' -or $msg -match 'No certificate source found' -or $msg -match 'required') { $status = 400 }
    if ($msg -match 'OrganizationDomain') { $status = 400 }

    Write-JsonResponse -StatusCode $status -BodyObject @{
        error         = $msg
        correlationId = $correlationId
        exception     = $_.Exception.ToString()
        stack         = $_.ScriptStackTrace
    }
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
}