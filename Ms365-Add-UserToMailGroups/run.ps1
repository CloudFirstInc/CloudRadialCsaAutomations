<#
.SYNOPSIS
    Adds a user to mail-enabled security groups and distribution lists using Microsoft Graph.
#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "📨 Function triggered: Adding user to mail-enabled groups..."

# Extract input
$userId = $Request.Body.UserId
$groupIds = $Request.Body.GroupIds
$tenantId = $Request.Body.TenantId
$ticketId = $Request.Body.TicketId
$securityKey = $Request.Headers.SecurityKey

# Validate input
if (-not $userId -or -not $groupIds -or $groupIds.Count -eq 0) {
    Write-Host "❌ Missing required parameters."
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::BadRequest
        Body = @{
            Message = "UserId and GroupIds are required."
            TicketId = $ticketId
        }
    })
    return
}

# Security check
if ($env:SecurityKey -and $securityKey -ne $env:SecurityKey) {
    Write-Host "❌ Invalid security key."
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::Forbidden
        Body = @{
            Message = "Invalid security key."
            TicketId = $ticketId
        }
    })
    return
}

# Connect to Microsoft Graph
Write-Host "🔐 Connecting to Microsoft Graph..."
$securePassword = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $securePassword)
Connect-MgGraph -ClientSecretCredential $credential -TenantId $tenantId
Write-Host "✅ Connected to Microsoft Graph."

# Add user to each group
$addedGroups = @()
$failedGroups = @()

foreach ($groupId in $groupIds) {
    try {
        $group = Get-MgGroup -GroupId $groupId
        if ($group.MailEnabled -eq $true) {
            New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId
            Write-Host "➕ Added to group: $($group.DisplayName)"
            $addedGroups += $group.DisplayName
        } else {
            Write-Host "⚠️ Skipped non-mail-enabled group: $($group.DisplayName)"
        }
    }
    catch {
        Write-Host "❌ Failed to add to group ID $groupId: $_"
        $failedGroups += $groupId
    }
}

# Return response
$response = @{
    Message = "Group assignment completed."
    TicketId = $ticketId
    AddedGroups = $addedGroups
    FailedGroups = $failedGroups
}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $response
    ContentType = "application/json"
})
Write-Host "✅ Function completed for TicketId: $ticketId"
