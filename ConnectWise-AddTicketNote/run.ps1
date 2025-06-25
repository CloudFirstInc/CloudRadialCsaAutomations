<#
.SYNOPSIS
    This function is used to add a note to a ConnectWise ticket.
.DESCRIPTION
    Requires environment variables:
        - ConnectWisePsa_ApiBaseUrl
        - ConnectWisePsa_ApiCompanyId
        - ConnectWisePsa_ApiPublicKey
        - ConnectWisePsa_ApiPrivateKey
        - ConnectWisePsa_ApiClientId
        - SecurityKey (optional)
.INPUTS
    JSON Structure:
    {
        "TicketId": "123456",
        "Message": "This is a note",
        "Internal": true,
        "SecurityKey": "optional"
    }
.OUTPUTS
    JSON structure of the response from the ConnectWise API
#>

using namespace System.Net

param($Request, $TriggerMetadata)

function Add-ConnectWiseTicketNote {
    param (
        [string]$ConnectWiseUrl,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId,
        [string]$TicketId,
        [string]$Text,
        [boolean]$Internal = $false
    )

    $apiUrl = "$ConnectWiseUrl/v4_6_release/apis/3.0/service/tickets/$TicketId/notes"

    $notePayload = @{
        ticketId = $TicketId
        text = $Text
        detailDescriptionFlag = $true
        internalAnalysisFlag = $Internal
    } | ConvertTo-Json -Depth 3

    $headers = @{
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${PublicKey}:${PrivateKey}"))
        "Content-Type" = "application/json"
        "clientId" = $ClientId
    }

    try {
        $result = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $notePayload -AllowInsecureRedirect
        if (-not $result) {
            $result = @{ Message = "Note added successfully, but no content returned." }
        }
    }
    catch {
        Write-Host "❌ API call failed: $_"
        $result = @{ Message = "API call failed: $($_.Exception.Message)" }
    }

    return $result
}


Write-Host "📥 Raw Request Body:"
Write-Host ($Request.Body | ConvertTo-Json -Depth 5)
Write-Host "🔍 Processing request to add a note to a ConnectWise ticket..."

# Extract request values
$TicketId = $Request.Body.TicketId
$Text = $Request.Body.Message
$Internal = $Request.Body.Internal
$SecurityKey = $env:SecurityKey

# Security check
if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "❌ Invalid security key"
    $body = @{ Result = @{ Message = "Invalid security key" } }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::Forbidden
        Body = $body
        ContentType = "application/json"
    })
    return
}

# Input validation
if (-Not $TicketId) {
    Write-Host "❌ Missing ticket number"
    $body = @{ Result = @{ Message = "Missing ticket number" } }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = $body
        ContentType = "application/json"
    })
    return
}

if (-Not $Text) {
    Write-Host "❌ Missing ticket text"
    $body = @{ Result = @{ Message = "Missing ticket text" } }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = $body
        ContentType = "application/json"
    })
    return
}

if (-Not $Internal) {
    $Internal = $false
}

Write-Host "📨 TicketId: $TicketId"
Write-Host "📝 Text: $Text"
Write-Host "🔒 Internal: $Internal"

# Call the function
$result = Add-ConnectWiseTicketNote -ConnectWiseUrl $env:ConnectWisePsa_ApiBaseUrl `
    -PublicKey "$env:ConnectWisePsa_ApiCompanyId+$env:ConnectWisePsa_ApiPublicKey" `
    -PrivateKey $env:ConnectWisePsa_ApiPrivateKey `
    -ClientId $env:ConnectWisePsa_ApiClientId `
    -TicketId $TicketId `
    -Text $Text `
    -Internal $Internal

Write-Host "📦 Final result: $($result | ConvertTo-Json -Depth 5)"

# Return response
$body = @{
    Result = $result
}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})
