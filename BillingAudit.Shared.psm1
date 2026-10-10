function Get-BillingAuditManagedIdentityToken {
    param([string]$Resource)
    if (-not $env:IDENTITY_ENDPOINT -or -not $env:IDENTITY_HEADER) { return $null }
    $uri = '{0}?resource={1}&api-version=2019-08-01' -f $env:IDENTITY_ENDPOINT,[Uri]::EscapeDataString($Resource)
    return (Invoke-RestMethod -Method Get -Uri $uri -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER } -TimeoutSec 30).access_token
}

function Get-BillingAuditGraphAccessToken {
    $token = Get-BillingAuditManagedIdentityToken -Resource 'https://graph.microsoft.com/'
    if ($token) { return $token }
    $tenantId = if ($env:Graph_TenantId) { $env:Graph_TenantId } else { 'palaisparc.onmicrosoft.com' }
    $clientId = if ($env:Graph_ClientId) { $env:Graph_ClientId } else { 'c0d6f1f4-f988-47d3-874c-1ac5fc855463' }
    if ([string]::IsNullOrWhiteSpace($env:Graph_ClientSecret)) { throw 'Managed identity is unavailable and Graph_ClientSecret is not configured.' }
    $body = @{ client_id=$clientId; client_secret=$env:Graph_ClientSecret; scope='https://graph.microsoft.com/.default'; grant_type='client_credentials' }
    return (Invoke-RestMethod -Method Post -Uri ('https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $tenantId) -Body $body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 30).access_token
}

function Invoke-BillingAuditGraphApi {
    param([string]$Uri,[string]$Method='Get',[string]$AccessToken,[object]$Body=$null,[int]$MaxRetries=5)
    $headers = @{ Authorization='Bearer {0}' -f $AccessToken; Accept='application/json'; 'Content-Type'='application/json' }
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $parameters = @{ Uri=$Uri; Method=$Method; Headers=$headers; ErrorAction='Stop'; TimeoutSec=60 }
            if ($null -ne $Body) { $parameters.Body = $Body | ConvertTo-Json -Depth 40 -Compress }
            return Invoke-RestMethod @parameters
        }
        catch {
            $status = $null; $retryAfter = $null
            try { $status = [int]$_.Exception.Response.StatusCode; $retryAfter = $_.Exception.Response.Headers['Retry-After'] } catch {}
            if ($attempt -eq $MaxRetries -or $status -notin @(429,500,502,503,504)) { throw }
            $seconds = 0
            if ($retryAfter) { [void][int]::TryParse([string]$retryAfter,[ref]$seconds) }
            if ($seconds -le 0) { $seconds = [Math]::Min(30,[int][Math]::Pow(2,$attempt)) + (Get-Random -Minimum 0 -Maximum 3) }
            Start-Sleep -Seconds $seconds
        }
    }
}

function Get-BillingAuditSharePointResources {
    param([string]$AccessToken)
    $siteId = $env:SharePoint_SiteId
    if (-not $siteId) { $siteId = (Invoke-BillingAuditGraphApi -Uri 'https://graph.microsoft.com/v1.0/sites/palaisparc.sharepoint.com:/finance' -AccessToken $AccessToken).id }
    function Resolve-ListId([string]$Name,[string]$ConfiguredId) {
        if ($ConfiguredId) { return $ConfiguredId }
        $filter = [Uri]::EscapeDataString("displayName eq '$Name'")
        $result = Invoke-BillingAuditGraphApi -Uri ('https://graph.microsoft.com/v1.0/sites/{0}/lists?$filter={1}' -f $siteId,$filter) -AccessToken $AccessToken
        $listMatches = @($result.value | Where-Object { $_.displayName -eq $Name })
        if ($listMatches.Count -ne 1) { throw ('Expected exactly one SharePoint list named "{0}".' -f $Name) }
        return $listMatches[0].id
    }
    return [pscustomobject]@{ SiteId=$siteId; RunsListId=(Resolve-ListId 'Billing Audit Runs' $env:SharePoint_BillingAuditRunsListId); DetailsListId=(Resolve-ListId 'Billing Audit Details' $env:SharePoint_BillingAuditDetailsListId) }
}

function New-BillingAuditSharePointItem {
    param([object]$Resources,[string]$ListId,[hashtable]$Fields,[string]$AccessToken)
    return Invoke-BillingAuditGraphApi -Uri ('https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items' -f $Resources.SiteId,$ListId) -Method Post -AccessToken $AccessToken -Body @{ fields=$Fields }
}

function Get-BillingAuditSharePointItemByRunId {
    param([object]$Resources,[string]$AuditRunId,[string]$AccessToken)
    $filter = [Uri]::EscapeDataString("fields/AuditRunId eq '$($AuditRunId.Replace("'","''"))'")
    $result = Invoke-BillingAuditGraphApi -Uri ('https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items?expand=fields&$filter={2}' -f $Resources.SiteId,$Resources.RunsListId,$filter) -AccessToken $AccessToken
    $items = @($result.value)
    if ($items.Count -ne 1) { throw ('Audit run "{0}" was not found exactly once.' -f $AuditRunId) }
    return $items[0]
}

function Update-BillingAuditSharePointItem {
    param([object]$Resources,[string]$ListId,[string]$ItemId,[hashtable]$Fields,[string]$AccessToken)
    return Invoke-BillingAuditGraphApi -Uri ('https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items/{2}/fields' -f $Resources.SiteId,$ListId,$ItemId) -Method Patch -AccessToken $AccessToken -Body $Fields
}

Export-ModuleMember -Function *
