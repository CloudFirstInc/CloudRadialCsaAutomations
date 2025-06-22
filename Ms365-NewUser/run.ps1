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
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey

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
Connect-MgGraph -ClientSecretCredential $credential -TenantId $env:Ms365_TenantId

# Generate UPN and MailNickName
$upn = "$($FirstName.ToLower()).$($LastName.ToLower())@yourdomain.com"
$mailNickName = "$($FirstName.ToLower())$($LastName.ToLower())"

# Create user
try {
    $newUser = New-MgUser -AccountEnabled $true `
        -DisplayName "$FirstName $MiddleName $LastName" `
        -MailNickname $mailNickName `
        -UserPrincipalName $upn `
        -PasswordProfile @{ ForceChangePasswordNextSignIn = $true; Password = "TempP@ssw0rd!" } `
        -GivenName $FirstName `
        -Surname $LastName `
        -Department $Department `
        -JobTitle $JobTitle `
        -OfficeLocation $OfficeLocation
    $message = "User $upn created successfully."
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
        $message += " Permissions cloned from $ModelUser."
    }
    catch {
        $message += " Failed to clone permissions from $ModelUser: $_"
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
