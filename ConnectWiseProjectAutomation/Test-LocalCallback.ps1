# Posts the sample payload to the locally running function (start it first with `func start` or F5).
[CmdletBinding()]
param(
    [string]$Uri = 'http://localhost:7071/api/ConnectWiseProjectCallback'
)

$payload = @{
    test      = $true
    source    = 'ManualTest'
    projectId = 12345
    status    = 'SharePoint Update'
} | ConvertTo-Json

Invoke-RestMethod -Uri $Uri -Method Post -ContentType 'application/json' -Body $payload
