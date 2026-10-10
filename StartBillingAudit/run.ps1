using namespace System.Net
param($Request, $TriggerMetadata)

Import-Module (Join-Path $PSScriptRoot '..\BillingAudit.Shared.psm1') -Force
$correlationId = [guid]::NewGuid().ToString()
function Send-Response([HttpStatusCode]$StatusCode,[object]$Body) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ StatusCode=$StatusCode; ContentType='application/json'; Body=($Body | ConvertTo-Json -Depth 10 -Compress) })
}
function Get-Value([object]$Body,[string]$Name) {
    $property = $Body.PSObject.Properties[$Name]
    if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return [string]$property.Value }
    return $null
}

try {
    $body = if ($Request.Body -is [string]) { $Request.Body | ConvertFrom-Json -ErrorAction Stop } else { $Request.Body }
    $customerName = Get-Value $body 'CustomerName'
    $monthText = Get-Value $body 'AuditMonth'
    $yearText = Get-Value $body 'AuditYear'
    $requestedBy = Get-Value $body 'RequestedBy'
    $month = 0; $year = 0
    if ([string]::IsNullOrWhiteSpace($customerName)) { throw [ArgumentException]::new('CustomerName is required.') }
    if (-not [int]::TryParse($monthText,[ref]$month) -or $month -notin 1..12) { throw [ArgumentException]::new('AuditMonth must be an integer from 1 through 12.') }
    if (-not [int]::TryParse($yearText,[ref]$year) -or $year -lt 2000 -or $year -gt 2100) { throw [ArgumentException]::new('AuditYear must be a four-digit year.') }
    if (-not $requestedBy) { $requestedBy = 'Copilot Finance Agent' }

    $periodStart = [datetime]::SpecifyKind([datetime]::new($year,$month,1),[DateTimeKind]::Utc)
    $periodEnd = $periodStart.AddMonths(1).AddSeconds(-1)
    $prefix = ($customerName -replace '[^A-Za-z0-9]','').Substring(0,[Math]::Min(8,($customerName -replace '[^A-Za-z0-9]','').Length)).ToUpperInvariant()
    $auditRunId = '{0}-{1:0000}{2:00}-{3}' -f $prefix,$year,$month,$correlationId.Substring(0,8)
    $token = Get-BillingAuditGraphAccessToken
    $resources = Get-BillingAuditSharePointResources -AccessToken $token
    Write-Host "[$correlationId] [INFO] SharePoint create. AuditRunId=$auditRunId; CustomerName=$customerName; TicketId="
    $runItem = New-BillingAuditSharePointItem -Resources $resources -ListId $resources.RunsListId -AccessToken $token -Fields @{
        Title=('Billing Audit {0}' -f $auditRunId); AuditRunId=$auditRunId; CustomerName=$customerName; AuditMonth=$month; AuditYear=$year
        PeriodStart=$periodStart.ToString('o'); PeriodEnd=$periodEnd.ToString('o'); RequestedBy=$requestedBy; StartedOn=(Get-Date).ToUniversalTime().ToString('o')
        RunStatus='Running'; AgreementsFound=0; ProjectsFound=0; TicketsFound=0; CorrelationId=$correlationId
    }
    Push-OutputBinding -Name AuditQueue -Value (@{ auditRunId=$auditRunId; runSharePointItemId=[string]$runItem.id; customerName=$customerName; auditMonth=$month; auditYear=$year; requestedBy=$requestedBy; correlationId=$correlationId } | ConvertTo-Json -Compress)
    Send-Response ([HttpStatusCode]::Accepted) @{ success=$true; auditRunId=$auditRunId; runSharePointItemId=[string]$runItem.id; runStatus='Running' }
}
catch {
    $status = if ($_.Exception -is [ArgumentException]) { [HttpStatusCode]::BadRequest } else { [HttpStatusCode]::InternalServerError }
    Send-Response $status @{ success=$false; correlationId=$correlationId; error=$_.Exception.Message }
}
