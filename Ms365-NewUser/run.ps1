<#
.SYNOPSIS
    This function creates a new Microsoft 365 user account and optionally clones group memberships and permissions from a model user.
#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "🔄 Function triggered: Starting user creation process..."
Import-Module Microsoft.Graph.Groups

# Initialize response
$resultCode = 200
$message = ""

# Extract input


$FirstName = $Request.Body.FirstName
$LastName = $Request.Body.LastName
$MiddleName = $Request.Body.MiddleName
$Department = $Request.Body.Department
$JobTitle = $Request.Body.JobTitle
$StartDate = $Request.Body.StartDate
$OfficeLocation = $Request.Body.OfficeLocation
$SoftwareNeeded = $Request.Body.SoftwareNeeded
$TypeofComputer = $Request.Body.TypeofComputer
$EmployeeType = $Request.Body.EmployeeType
$ModelUser = $Request.Body.ModelUser
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey
if (-not $FirstName -or -not $LastName) {
    $message = "FirstName and LastName are required."
    $resultCode = 400
    Write-Host "❌ Missing required fields: FirstName or LastName."
    return
}

Write-Host "📥 Input received: FirstName=${FirstName}, LastName=${LastName}, ModelUser=${ModelUser}, TicketId=${TicketId}, StartDate=${StartDate}"

# Use environment variable if TenantId is not provided
if (-not $TenantId) {
    $TenantId = $env:Ms365_TenantId
    Write-Host "ℹ️ TenantId not provided. Using default from environment."
} else {
    Write-Host "✅ TenantId provided: ${TenantId}"
}

# Validate TenantId format
if (-not $TenantId -or $TenantId -notmatch '^[0-9a-fA-F\-]{36}$') {
    $message = "Invalid or missing TenantId. Please provide a valid GUID."
    $resultCode = 400
    Write-Host "❌ Invalid TenantId format: ${TenantId}"
    return
}

# Security check
if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    $message = "Invalid security key"
    $resultCode = 403
    Write-Host "❌ Security key validation failed."
    return
}

# Validate required fields
if (-not $FirstName -or -not $LastName) {
    $message = "FirstName and LastName are required."
    $resultCode = 400
    Write-Host "❌ Missing required fields: FirstName or LastName."
    return
}

# Connect to Microsoft Graph
Write-Host "🔐 Connecting to Microsoft Graph..."
$securePassword = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $securePassword)
Connect-MgGraph -ClientSecretCredential $credential -TenantId $TenantId
Write-Host "✅ Connected to Microsoft Graph."

# Get default domain
Write-Host "🌐 Retrieving default domain for tenant..."
$domains = Get-MgDomain
$defaultDomain = $domains | Where-Object { $_.IsDefault -eq $true }

if (-not $defaultDomain) {
    $message = "Could not retrieve default domain for tenant."
    $resultCode = 500
    Write-Host "❌ Failed to retrieve default domain."
    return
}

$domainName = $defaultDomain.Id
$upn = "${FirstName}.${LastName}@${domainName}"
$mailNickName = "${FirstName}${LastName}"
$displayName = "${FirstName} ${MiddleName} ${LastName}"
Write-Host "✅ Default domain resolved: ${domainName}"
Write-Host "🛠️ Creating user: ${displayName} (${upn})..."

# Use hardcoded password
$randomPassword = "TempP@ssw0rd!"

# Create user using splatting
try {
    $userParams = @{
        AccountEnabled    = $true
        DisplayName       = $displayName
        MailNickname      = $mailNickName
        UserPrincipalName = $upn
        PasswordProfile   = @{
            ForceChangePasswordNextSignIn = $true
            Password = $randomPassword
        }
        GivenName         = $FirstName
        Surname           = $LastName
        Department        = $Department
        JobTitle          = $JobTitle
        OfficeLocation    = $OfficeLocation
    }

    $newUser = New-MgUser @userParams
    $message = "User ${upn} created successfully."
    Write-Host "✅ User created: ${upn}"
}
catch {
    $message = "Failed to create user: $_"
    $resultCode = 500
    Write-Host "❌ Error creating user: $_"
    return
}

# Clone model user permissions and groups
if ($ModelUser) {
    Write-Host "🔄 Cloning group memberships from model user: ${ModelUser}"
    try {
        $modelUserObj = Get-MgUser -Filter "userPrincipalName eq '${ModelUser}'"
        if (-not $modelUserObj) {
            throw "Model user not found."
        }

        $groups = Get-MgUserMemberOf -UserId $modelUserObj.Id -All
        $addedGroups = @()
        $skippedGroups = @()

        foreach ($group in $groups) {
            if ($group.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group") {
                $groupId = $group.Id
                $groupName = $group.DisplayName
                $mailEnabled = $group.AdditionalProperties["mailEnabled"]
                $securityEnabled = $group.AdditionalProperties["securityEnabled"]

                if ($mailEnabled -eq $true -and $securityEnabled -eq $true) {
                    Write-Host "⚠️ Skipping mail-enabled security group: $groupName"
                    $skippedGroups += $groupName
                    continue
                }

                if ($groupId) {
                    try {
                        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $newUser.Id
                        Write-Host "➕ Added to group: ${groupName}"
                        $addedGroups += $groupName
                    }
                    catch {
                        Write-Host "⚠️ Failed to add to group ${groupName}: $_"
                    }
                }
            }
        }

        if ($addedGroups.Count -gt 0) {
            $message += " Added to groups: " + ($addedGroups -join ", ") + "."
        }

        if ($skippedGroups.Count -gt 0) {
            $message += " Skipped mail-enabled groups: " + ($skippedGroups -join ", ") + "."
        }

        Write-Host "✅ Group memberships cloned."
    }
    catch {
        $message += " Failed to clone permissions from ${ModelUser}: $_"
        Write-Host "⚠️ Error cloning permissions: $_"
    }
}

# Return response
Write-Host "📤 Returning response..."
$body = @{
    Message = $message
    TicketId = $TicketId
    ResultCode = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
}

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
    ContentType = "application/json"
})
