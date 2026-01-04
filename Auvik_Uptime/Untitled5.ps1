
# Ensure TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$region = "us5"
$user   = "dgentles@cloudfirstinc.com"
$apiKey = "i9sh4CMEMMZQRgguk73fJ+eXUsnX2Mo4GWoestxWeg6UnVni"

# Build Basic token: base64("user:apiKey")
$credentialPair = "$($user):$apiKey"      # <-- key fix here
$basicToken = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($credentialPair))

$headers = @{
    Authorization = "Basic $basicToken"
    Accept        = "application/json"
}

$uri = "https://auvikapi.$region.my.auvik.com/v1/tenants?page[first]=200"

try {
    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
    $response | ConvertTo-Json -Depth 6
}
catch {
    Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__
    Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
    if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream()) {
        $reader = New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())
        Write-Host "Body:" $reader.ReadToEnd()
    }
    throw
}
