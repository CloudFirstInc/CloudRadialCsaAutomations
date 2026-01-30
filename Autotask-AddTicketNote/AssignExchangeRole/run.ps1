<#

.SYNOPSIS
    
    This function is to create a new Security group in Microsoft 365.

.DESCRIPTION
    
    This function is to create a new security group in Microsoft 365.
    
    The function requires the following environment variables to be set:
    
    Ms365_AuthAppId - Application Id of the service principal
    Ms365_AuthSecretId - Secret Id of the service principal
    Ms365_TenantId - Tenant Id of the Microsoft 365 tenant
    SecurityKey - Optional, use this as an additional step to secure the function
 
    The function requires the following modules to be installed:
    
    Microsoft.Graph
    
.INPUTS


    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    TicketId - optional - string value of the ticket id used for transaction tracking
    SecurityKey - Optional, use this as an additional step to secure the function

    JSON Structure

    {

        "TenantId": "12345678-1234-1234-123456789012",
        "TicketId": "123456,
        "SecurityKey", "optional"
    }

.OUTPUTS

    JSON response with the following fields:

    Message - Descriptive string of result
    TicketId - TicketId passed in Parameters
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"

#>
using namespace System.Net

param($Request, $TriggerMetadata)

# Import the required module
# Import-Module Microsoft.Graph

# Define the variables
$AppID = $env:Ms365_AuthAppId
$RoleName = "Exchange Administrator"
$TenantID = $Request.Body.Company.CompanyTenantId
$TicketId = $Request.Body.Ticket.TicketId
$SecurityKey = $env:SecurityKey

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break
}

try {
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)
    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId
    
    # Use Connect-MgGraph with ClientSecretCredential
    #$clientSecretCredential = [Microsoft.Graph.PowerShell.Authentication.ClientSecretCredential]::new($env:Ms365_TenantId, $env:Ms365_AuthAppId, $env:Ms365_AuthSecretId)
    #Connect-MgGraph -ClientSecretCredential $clientSecretCredential -Scopes "RoleManagement.ReadWrite.Directory", "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

    # Get the role definition ID for the Exchange Administrator role
    $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq '$RoleName'"
    if ($null -eq $roleDefinitions -or $roleDefinitions.Count -eq 0) {
        throw "Role definition not found for $RoleName"
    }
    $roleDefinition = $roleDefinitions[0]
    $roleDefinitionId = $roleDefinition.Id

    # Get the service principal ID for the application
    $servicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$AppID'"
    $servicePrincipalId = $servicePrincipal.Id

    # Assign the role to the service principal in the remote tenant
    New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $servicePrincipalId -RoleDefinitionId $roleDefinitionId -DirectoryScopeId "/" -DirectoryScopeTenantId $TenantID

    # Output bindings are passed out via return value.
    $Response = @{
        status = 200
        body = "Role assignment successful"
        TicketId = $TicketId
        message = "The Exchange Administrator role has been assigned to the CloudFirstAutomation Enterprise Application service principal"
    } 

    # Send the response to the webhook
    $webhookUrl = "https://cloudfirstfunctions.azurewebsites.net/ConnectWise-AddTicketNote?code=rPhgABesk0e55iUT-v6Upe9AhWgySs86q6Fm0lTeG7FtAzFuohJekg%3D%3D"
    Invoke-RestMethod -Uri $webhookUrl -Method Post -Body ($Response | ConvertTo-Json) -ContentType "application/json"
} catch {
    Write-Host "An error occurred: $_"
    $Response = @{
        status = 500
        body = "Internal Server Error"
    }
}
return $Response

return $Response
