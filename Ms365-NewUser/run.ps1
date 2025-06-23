
<#

.SYNOPSIS

    This function creates a new Microsoft 365 user account and optionally clones group memberships and permissions from a model user.

.DESCRIPTION

    This function provisions a new Microsoft 365 user using the Microsoft Graph API. It supports copying group memberships and settings from a model user if specified.

    The function requires the following environment variables to be set:

    Ms365_AuthAppId     - Application Id of the service principal
    Ms365_AuthSecretId  - Secret Id of the service principal
    Ms365_TenantId      - Default Tenant Id of the Microsoft 365 tenant
    SecurityKey         - Optional, used as an additional step to secure the function

    The function requires the following module to be installed:

    Microsoft.Graph

.INPUTS

    FirstName     - First name of the new user
    LastName      - Last name of the new user
    MiddleName    - Middle name of the new user (optional)
    Department    - Department name
    JobTitle      - Job title
    StartDate     - Start date of the user (optional)
    OfficeLocation- Office location
    ModelUser     - UPN of an existing user to model group memberships and permissions after (optional)
    TenantId      - Tenant Id to use for the request; if blank, uses the environment variable Ms365_TenantId
    TicketId      - Optional tracking ID for the request
    SecurityKey   - Optional security key for validating the request

    JSON Structure:

    {
        "FirstName": "@FirstName",
        "LastName": "@LastName",
        "MiddleName": "@MiddleName",
        "Department": "@Department",
        "JobTitle": "@JobTitle",
        "StartDate": "@StartDate",
        "OfficeLocation": "@OfficeLocation",
        "ModelUser": "@ModelUser",
        "TenantId": "@TenantId",
        "TicketId": "@TicketId",
        "SecurityKey": "@SecurityKey"
    }

.OUTPUTS

    JSON response with the following fields:

    Message       - Descriptive string of the result
    TicketId      - TicketId passed in parameters
    ResultCode    - 200 for success, 400/403/500 for various failure conditions
    ResultStatus  - "Success" or "Failure"

#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "🔄 Function triggered: Starting user creation process..."

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
$ModelUser = $Request.Body.ModelUser
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey

Write-Host "📥 Input received: FirstName=${FirstName}, LastName=${LastName}, ModelUser=${ModelUser}"

# Use environment variable if TenantId is not provided
if (-not $TenantId) {
    $TenantId = $env:Ms365_TenantId
    Write-Host "ℹ️ TenantId not provided. Using default from environment."
} else {
    Write-Host "✅ TenantId provided: $TenantId"
}

# Validate TenantId format
if (-not $TenantId -or $TenantId -notmatch '^[0-9a-fA-F\-]{36}$') {
    $message = "Invalid or missing TenantId. Please provide a valid GUID."
    $resultCode = 400
    Write-Host "❌ Invalid TenantId format: $TenantId"
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
$upn = "${FirstName}.${LastName}@$domainName"
$mailNickName = "${FirstName}${LastName}"
$displayName = "${FirstName} ${MiddleName} ${LastName}"
Write-Host "✅ Default domain resolved: $domainName"
Write-Host "🛠️ Creating user: $displayName ($upn)..."

# Create user using splatting
try {
    $userParams = @{
        AccountEnabled    = $true
        DisplayName       = $displayName
        MailNickname      = $mailNickName
        UserPrincipalName = $upn
        PasswordProfile   = @{
            ForceChangePasswordNextSignIn = $true
            Password = "TempP@ssw0rd!"
        }
        GivenName         = $FirstName
        Surname           = $LastName
        Department        = $Department
        JobTitle          = $JobTitle
        OfficeLocation    = $OfficeLocation
    }

    $newUser = New-MgUser @userParams
    $message = "User ${upn} created successfully."
    Write-Host "✅ User created: $upn"
}
catch {
    $message = "Failed to create user: $_"
    $resultCode = 500
    Write-Host "❌ Error creating user: $_"
    return
}

# Clone model user permissions and groups
if ($ModelUser) {
    Write-Host "🔄 Cloning group memberships from model user: $ModelUser"
    try {
        $modelUserObj = Get-MgUser -Filter "userPrincipalName eq '$ModelUser'"
        if (-not $modelUserObj) {
            throw "Model user not found."
        }

        $groups = Get-MgUserMemberOf -UserId $modelUserObj.Id -All

        foreach ($group in $groups) {
            if ($group.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group") {
                $groupId = $group.Id
                if ($groupId) {
                    try {
                        Add-MgGroupMember -GroupId $groupId -DirectoryObjectId $newUser.Id
                        Write-Host "➕ Added to group: $groupId"
                    }
                    catch {
                        Write-Host "⚠️ Failed to add to group $groupId: $_"
                    }
                }
            }
        }

        $message += " Permissions cloned from ${ModelUser}."
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
