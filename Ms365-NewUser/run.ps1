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

Write-Host "Create M365 User function triggered."

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

# Use environment variable if TenantId is not provided
if (-not $TenantId) {
    $TenantId = $env:Ms365_TenantId
}

# Security check
if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    $message = "Invalid security key"
    $resultCode = 403
    goto End
}

# Validate required fields
if (-not $FirstName -or -not $LastName) {
    $message = "FirstName and LastName are required."
    $resultCode = 400
    goto End
}

# Connect to Microsoft Graph
$securePassword = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $securePassword)
Connect-MgGraph -ClientSecretCredential $credential -TenantId $TenantId

# Generate UPN and MailNickName
$upn = "${FirstName}.${LastName}@yourdomain.com"
$mailNickName = "${FirstName}${LastName}"
$displayName = "${FirstName} ${MiddleName} ${LastName}"

# Create user
try {
    $newUser = New-MgUser -AccountEnabled $true `
        -DisplayName $displayName `
        -MailNickname $mailNickName `
        -UserPrincipalName $upn `
        -PasswordProfile @{ ForceChangePasswordNextSignIn = $true; Password = "TempP@ssw0rd!" } `
        -GivenName $FirstName `
        -Surname $LastName `
        -Department $Department `
        -JobTitle $JobTitle `
        -OfficeLocation $OfficeLocation
    $message = "User ${upn} created successfully."
}
catch {
    $message = "Failed to create user: $_"
    $resultCode = 500
    goto End
}

# Clone model user permissions and groups
if ($ModelUser) {
    try {
        $modelUserObj = Get-MgUser -Filter "userPrincipalName eq '$ModelUser'"
        $groups = Get-MgUserMemberOf -UserId $modelUserObj.Id
        foreach ($group in $groups) {
            if ($group.'@odata.type' -eq "#microsoft.graph.group") {
                Add-MgGroupMember -GroupId $group.Id -DirectoryObjectId $newUser.Id
            }
        }
        $message += " Permissions cloned from ${ModelUser}."
    }
    catch {
        $message += " Failed to clone permissions from ${ModelUser}: $_"
    }
}

End:
# Return response
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
