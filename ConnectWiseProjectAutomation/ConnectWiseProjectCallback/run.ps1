using namespace System.Net

# Diagnostic receiver for ConnectWise PSA callbacks.
# Logs what arrives (secrets redacted) and returns 200. It makes no outbound calls and changes nothing.
param($Request, $TriggerMetadata)

$MaxBodyLogChars = 32768
$SensitivePattern = '(?i)auth|key|secret|token|password|pwd|credential|cookie|signature|^code$|^sig$'

function Get-SafeMap {
    # Flattens a header/query collection to an ordered map with sensitive values redacted.
    param($Source)
    $map = [ordered]@{}
    if ($null -eq $Source) { return $map }
    foreach ($key in @($Source.Keys | Sort-Object)) {
        $map[[string]$key] = if ($key -match $SensitivePattern) { '[REDACTED]' } else { [string]$Source[$key] }
    }
    $map
}

function Write-JsonLog {
    param([string]$Label, $Value)
    Write-Host ("{0}: {1}" -f $Label, ($Value | ConvertTo-Json -Depth 20))
}

$timestampUtc = [DateTime]::UtcNow.ToString('o')

try {
    $invocationId = $TriggerMetadata.InvocationId
    if (-not $invocationId) { $invocationId = $TriggerMetadata.sys.InvocationId }
    if (-not $invocationId) { $invocationId = "generated-$([guid]::NewGuid())" }

    $headers = Get-SafeMap $Request.Headers
    $query   = Get-SafeMap $Request.Query

    # RawBody preserves exactly what ConnectWise sent, regardless of how the worker auto-parsed it.
    $rawBody = if ($null -ne $Request.RawBody) { [string]$Request.RawBody }
               elseif ($null -ne $Request.Body) { [string]$Request.Body }
               else { '' }

    Write-Host "=== ConnectWiseProjectCallback received ==="
    Write-Host "TimestampUtc : $timestampUtc"
    Write-Host "InvocationId : $invocationId"
    Write-Host "Method       : $($Request.Method)"
    Write-Host "Content-Type : $($Request.Headers['content-type'])"
    Write-Host "User-Agent   : $($Request.Headers['user-agent'])"
    Write-JsonLog 'Query' $query
    Write-JsonLog 'Headers (sensitive values redacted)' $headers

    if ([string]::IsNullOrWhiteSpace($rawBody)) {
        Write-Host 'Body: <empty>'
    }
    else {
        $logBody = $rawBody
        if ($logBody.Length -gt $MaxBodyLogChars) {
            $logBody = $logBody.Substring(0, $MaxBodyLogChars) + "...[truncated, $($rawBody.Length) chars total]"
        }
        Write-Host "Body (raw, $($rawBody.Length) chars): $logBody"

        try {
            $parsed = $rawBody | ConvertFrom-Json -ErrorAction Stop
            Write-Host "Body (parsed JSON):`n$($parsed | ConvertTo-Json -Depth 20)"
        }
        catch {
            Write-Host 'Body is not valid JSON; raw body logged above.'
        }
    }

    $response = [ordered]@{
        received     = $true
        handler      = 'ConnectWiseProjectCallback'
        timestampUtc = $timestampUtc
    }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        ContentType = 'application/json'
        Body        = ($response | ConvertTo-Json -Compress)
    })
}
catch {
    # Details go to the log only; the caller gets a generic message.
    Write-Error "ConnectWiseProjectCallback failed: $($_.Exception.Message)"
    $response = [ordered]@{
        received     = $false
        handler      = 'ConnectWiseProjectCallback'
        timestampUtc = $timestampUtc
        error        = 'Internal error while processing callback.'
    }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::InternalServerError
        ContentType = 'application/json'
        Body        = ($response | ConvertTo-Json -Compress)
    })
}
