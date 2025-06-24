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
    Write-Host "🔎 Raw Request Body: $($Request.Body | ConvertTo-Json -Depth 5)"


    # Helper function to extract values from Ticket.Questions
    function Get-QuestionValue($questions, $id) {
        return ($questions | Where-Object { $_.Id -eq $id }).Value
    }

    # Extract input from nested structure
    $questions = $Request.Body.Ticket.Questions

    $FirstName = Get-QuestionValue $questions "FirstName"
    $LastName = Get-QuestionValue $questions "LastName"
    $MiddleName = Get-QuestionValue $questions "MiddleName"
    $Department = Get-QuestionValue $questions "Department"
    $JobTitle = Get-QuestionValue $questions "Title"
    $StartDate = Get-QuestionValue $questions "StartDate"
    $OfficeLocation = Get-QuestionValue $questions "OfficeLocation"
    $SoftwareNeeded = Get-QuestionValue $questions "SoftwareNeeded"
    $TypeofComputer = Get-QuestionValue $questions "TypeofComputer"
    $EmployeeType = Get-QuestionValue $questions "EmployeeType"
    $ModelUser = Get-QuestionValue $questions "ModelUser"
    $TenantId = Get-QuestionValue $questions "TenantId"
    $TicketId = $Request.Body.Ticket.TicketString
    $SecurityKey = $env:SecurityKey

    # Validate required fields
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

    # Connect to Microsoft Graph
    Write-Host "🔐 Connecting to Microsoft Graph..."
    $securePassword = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $securePassword)
    Connect-MgGraph -ClientSecretCredential $credential -TenantId $TenantId
    Write-Host "✅ Connected to Microsoft Graph."

    # 🌐 Retrieve default domain with null check
    Write-Host "🌐 Retrieving default domain for tenant..."
    $domains = Get-MgDomain
    if (-not $domains) {
        $message = "Could not retrieve domains for tenant."
        $resultCode = 500
        Write-Host "❌ No domains returned from Microsoft Graph."
        return
    }

    $defaultDomain = $domains | Where-Object { $_.IsDefault -eq $true }
    if (-not $defaultDomain) {
        $message = "Could not retrieve default domain for tenant."
        $resultCode = 500
        Write-Host "❌ Failed to retrieve default domain."
        return
    }

    $domainName = $defaultDomain.Id
    $firstInitial = $FirstName.Substring(0,1)
    $upn = "${firstInitial}${LastName}@${domainName}".ToLower()
    $mailNickName = "${firstInitial}${LastName}".ToLower()

    # 🧠 Display name formatting with optional middle name
    if ($MiddleName) {
        $displayName = "$FirstName $MiddleName $LastName"
    } else {
        $displayName = "$FirstName $LastName"
    }

    Write-Host "✅ Default domain resolved: ${domainName}"
    Write-Host "🛠️ Creating user: ${displayName} (${upn})..."

    # Use hardcoded password (consider replacing with secure generation)
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

    # 🔄 Clone group memberships from model user
    if ($ModelUser) {
        Write-Host "🔄 Cloning group memberships from model user: ${ModelUser}"
        try {
            $modelUserObj = Get-MgUser -Filter "userPrincipalName eq '${ModelUser}'"
            if (-not $modelUserObj) {
                throw "Model user not found."
            }

            $groupRefs = Get-MgUserMemberOf -UserId $modelUserObj.Id -All
            $groups = foreach ($groupRef in $groupRefs) {
                if ($groupRef.'@odata.type' -eq "#microsoft.graph.group") {
                    Get-MgGroup -GroupId $groupRef.Id
                }
            }

            $addedGroups = @()
            $skippedGroups = @()
Write-Host "🔍Here are All Groups $groups"
            foreach ($group in $groups) {
                $groupName = $group.DisplayName
                $mailEnabled = $group.MailEnabled
                $securityEnabled = $group.SecurityEnabled
                Write-Host "🔍 Processing group: ${groupName} (MailEnabled: $mailEnabled, SecurityEnabled: $securityEnabled)"

                if ($mailEnabled -eq $true -and $securityEnabled -eq $true) {
                    Write-Host "⚠️ Skipping mail-enabled security group: $groupName"
                    $skippedGroups += $groupName
                    continue
                }

                try {
                    New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $newUser.Id
                    Write-Host "➕ Added to group: ${groupName}"
                    $addedGroups += $groupName
                }
                catch {
                    Write-Host "⚠️ Failed to add to group ${groupName}: $_"
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
