using namespace System.Net
param($Request, $TriggerMetadata)

Import-Module (Join-Path $PSScriptRoot '..\BillingAudit.Shared.psm1') -Force
$correlationId = [guid]::NewGuid().ToString()
function Send-Response([HttpStatusCode]$StatusCode,[object]$Body) {
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ StatusCode=$StatusCode; ContentType='application/json'; Body=($Body | ConvertTo-Json -Depth 10 -Compress) })
}
try {
    $auditRunId = [string]$Request.Params.auditRunId
    if ([string]::IsNullOrWhiteSpace($auditRunId)) { throw [ArgumentException]::new('auditRunId is required.') }
    $token = Get-BillingAuditGraphAccessToken
    $resources = Get-BillingAuditSharePointResources -AccessToken $token
    $item = Get-BillingAuditSharePointItemByRunId -Resources $resources -AuditRunId $auditRunId -AccessToken $token
    $fields = $item.fields
    Send-Response ([HttpStatusCode]::OK) @{
        success=$true; auditRunId=$auditRunId; runSharePointItemId=[string]$item.id; runStatus=[string]$fields.RunStatus
        ticketsReviewed=[int]$fields.TicketsFound; ticketsClassified=[int]$fields.TicketsClassified; covered=[int]$fields.CoveredCount
        project=[int]$fields.ExistingProjectCount; billSeparately=[int]$fields.BillSeparatelyCount; manualReview=[int]$fields.ManualReviewCount; failed=[int]$fields.FailedTicketCount
    }
}
catch {
    $status = if ($_.Exception -is [ArgumentException]) { [HttpStatusCode]::BadRequest } else { [HttpStatusCode]::InternalServerError }
    Send-Response $status @{ success=$false; correlationId=$correlationId; error=$_.Exception.Message }
}
