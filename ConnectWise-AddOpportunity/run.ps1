using namespace System.Net

param($Request, $TriggerMetadata)

function Create-ConnectWiseOpportunity {
    param (
        [string]$ConnectWiseUrl,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId,
        [string]$CompanyId,
        [string]$OpportunityName,
        [string]$SalesStage = "Lead",
        [datetime]$ExpectedCloseDate = (Get-Date).AddDays(30),
        [decimal]$ExpectedRevenue = 1000.00
    )

    $apiUrl = "$ConnectWiseUrl/v4_6_release/apis/3.0/sales/opportunities"

    $opportunityPayload = @{
        name = $OpportunityName
        company = @{
            id = $CompanyId
        }
        expectedCloseDate = $ExpectedCloseDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        salesStage = $SalesStage
        expectedRevenue = $ExpectedRevenue
    } | ConvertTo-Json -Depth 3

    $headers = @{
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${PublicKey}:${PrivateKey}"))
        "Content-Type" = "application/json"
        "clientId" = $ClientId
    }

    $result = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $opportunityPayload -AllowInsecureRedirect
    return $result
}

# Extract inputs
$CompanyId = $Request.Body.CompanyId
$OpportunityName = $Request.Body.OpportunityName
$SalesStage = $Request.Body.SalesStage
$ExpectedCloseDate = $Request.Body.ExpectedCloseDate
$ExpectedRevenue = $Request.Body.ExpectedRevenue
$SecurityKey = $env:SecurityKey

# Security check
if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break;
}

# Validation
if (-Not $CompanyId) {
    Write-Host "Missing company ID"
    break;
}
if (-Not $OpportunityName) {
    Write-Host "Missing opportunity name"
    break;
}

# Call the function
$result = Create-ConnectWiseOpportunity -ConnectWiseUrl $env:ConnectWisePsa_ApiBaseUrl `
    -PublicKey "$env:ConnectWisePsa_ApiCompanyId+$env:ConnectWisePsa_ApiPublicKey" `
    -PrivateKey $env:ConnectWisePsa_ApiPrivateKey `
    -ClientId $env:ConnectWisePsa_ApiClientId `
    -CompanyId $CompanyId `
    -OpportunityName $OpportunityName `
    -SalesStage $SalesStage `
    -ExpectedCloseDate $ExpectedCloseDate `
    -ExpectedRevenue $ExpectedRevenue

# Return response
$body = @{
    response = ($result | ConvertTo-Json -Depth 3);
} 

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})
