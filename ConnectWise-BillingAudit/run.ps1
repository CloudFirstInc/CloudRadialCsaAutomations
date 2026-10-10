<#
.SYNOPSIS
    Starts a monthly CloudFirst managed-services billing audit.

.DESCRIPTION
    HTTP-triggered PowerShell Azure Function. Accepts CustomerName, AuditMonth, and AuditYear.
    The function resolves the ConnectWise company, retrieves approved managed-services agreements,
    open projects and project tickets, retrieves service tickets and time entries for the audit
    period, gathers ticket notes, calls Azure OpenAI for a recommendation, and saves one result per
    ticket to a SharePoint list through Microsoft Graph.

    This function is recommendation-only. It does not update ConnectWise tickets.

.REQUIRED ENVIRONMENT VARIABLES
    ConnectWisePsa_ApiBaseUrl
    ConnectWisePsa_ApiCompanyId
    ConnectWisePsa_ApiPublicKey
    ConnectWisePsa_ApiPrivateKey
    ConnectWisePsa_ApiClientId

    AzureOpenAI_Endpoint              Example: https://resource.openai.azure.com
    AzureOpenAI_Deployment            Deployment name, not base model name

    SharePoint_SiteId                 Microsoft Graph site ID (optional; resolved from the finance site when absent)
    SharePoint_BillingAuditRunsListId Optional Billing Audit Runs list ID
    SharePoint_BillingAuditDetailsListId Optional Billing Audit Details list ID

.AUTHENTICATION ENVIRONMENT VARIABLES
    Azure OpenAI, choose one:
      AzureOpenAI_ApiKey
      -or- enable managed identity and grant Cognitive Services OpenAI User

    Microsoft Graph, choose one:
      Managed identity with Sites.Selected/Sites.ReadWrite.All as appropriate
      -or-
      Graph_TenantId
      Graph_ClientId
      Graph_ClientSecret

.OPTIONAL ENVIRONMENT VARIABLES
    SecurityKey
    DebugLogging                      1 enables verbose logs
    AzureOpenAI_ApiVersion            Default: 2024-10-21
    BillingAudit_MaxTickets           Default: 250
    BillingAudit_MaxNotesCharacters   Default: 24000
    BillingAudit_DryRun               Default: 0. Set 1 to skip GPT and SharePoint writes.
    SharePoint_ListItemTitlePrefix    Default: Billing Audit

.INPUTS
    POST JSON:
    {
      "CustomerName": "City of Rye",
      "AuditMonth": 9,
      "AuditYear": 2026,
      "RequestedBy": "Dennis Gentles",
      "SecurityKey": "optional"
    }

.OUTPUTS
    JSON summary containing the audit run ID, customer, date range, agreements, projects,
    counts, SharePoint item IDs, warnings, and per-ticket processing errors.

.NOTES
    Confirm ConnectWise endpoint availability and field names in your PSA version before production.
    Confirm SharePoint internal column names in New-SharePointAuditFields.
#>

using namespace System.Net
param($QueueItem, $TriggerMetadata)

$CorrelationId = [guid]::NewGuid().ToString()
$DebugEnabled = ($env:DebugLogging -eq '1')

$AllowedAgreementTypes = @(
    'Essentials CoManaged Services',
    'Essentials Managed Services',
    'Gold CoManaged Services',
    'Gold Managed Services',
    'Platinum CoManaged Services',
    'Platinum Managed Services',
    'Video Surveillance Managed Services'
)

function Write-Info { param([string]$Message) Write-Host ('[{0}] [INFO] {1}' -f $CorrelationId, $Message) }
function Write-DebugLog { param([string]$Message) if ($DebugEnabled) { Write-Host ('[{0}] [DEBUG] {1}' -f $CorrelationId, $Message) } }
function Write-ErrorLog { param([string]$Message) Write-Host ('[{0}] [ERROR] {1}' -f $CorrelationId, $Message) }

function Send-JsonResponse {
    param([HttpStatusCode]$StatusCode, [object]$Body)
    Write-Host ('Queue worker completed with status {0}: {1}' -f $StatusCode,($Body | ConvertTo-Json -Depth 20 -Compress))
}

function Assert-BadRequest { param([string]$Message) throw [System.ArgumentException]::new($Message) }

function Test-HasProperty {
    param([object]$Object, [string]$PropertyName)
    return ($null -ne $Object -and $null -ne $Object.PSObject.Properties[$PropertyName])
}

function Get-RequestBodyObject {
    param($Request)
    if ($null -eq $Request.Body) { return $null }
    if ($Request.Body -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Request.Body)) { return $null }
        try { return $Request.Body | ConvertFrom-Json }
        catch { Assert-BadRequest 'Request body is not valid JSON.' }
    }
    return $Request.Body
}

function Get-InputValue {
    param($Request, $Body, [string[]]$Names)
    foreach ($name in $Names) {
        if ($Body) {
            $prop = $Body.PSObject.Properties | Where-Object { $_.Name.ToLowerInvariant() -eq $name.ToLowerInvariant() } | Select-Object -First 1
            if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) { return [string]$prop.Value }
        }
        if ($Request.Query) {
            foreach ($key in $Request.Query.Keys) {
                if ($key.ToString().ToLowerInvariant() -eq $name.ToLowerInvariant()) {
                    $value = $Request.Query[$key]
                    if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
                }
            }
        }
        if ($Request.Headers) {
            foreach ($key in $Request.Headers.Keys) {
                if ($key.ToString().ToLowerInvariant() -eq $name.ToLowerInvariant()) {
                    $value = $Request.Headers[$key]
                    if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
                }
            }
        }
    }
    return $null
}

function Assert-Settings {
    param([string[]]$Names)
    $missing = @()
    foreach ($name in $Names) {
        $value = [Environment]::GetEnvironmentVariable($name)
        if ([string]::IsNullOrWhiteSpace($value)) { $missing += $name }
    }
    if ($missing.Count -gt 0) { throw ('Missing required application settings: {0}' -f ($missing -join ', ')) }
}

function ConvertTo-CWConditionString {
    param([string]$Value)
    return $Value.Replace('\\', '\\\\').Replace('"', '\"')
}

function New-ConnectWiseHeaders {
    $auth = '{0}+{1}:{2}' -f $env:ConnectWisePsa_ApiCompanyId, $env:ConnectWisePsa_ApiPublicKey, $env:ConnectWisePsa_ApiPrivateKey
    $encoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($auth))
    return @{ Authorization = 'Basic {0}' -f $encoded; ClientId = $env:ConnectWisePsa_ApiClientId; Accept = 'application/json'; 'Content-Type' = 'application/json' }
}

function Invoke-ApiWithRetry {
    param([string]$Uri, [string]$Method = 'Get', [hashtable]$Headers, [object]$Body = $null, [int]$MaxRetries = 5)
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-DebugLog ('HTTP {0} {1}, attempt {2}' -f $Method, $Uri, $attempt)
            $params = @{ Uri = $Uri; Method = $Method; Headers = $Headers; ErrorAction = 'Stop'; TimeoutSec = 60 }
            if ($null -ne $Body) {
                $params.Body = $Body | ConvertTo-Json -Depth 80 -Compress
                $params.ContentType = 'application/json'
            }
            return Invoke-RestMethod @params
        }
        catch {
            $status = $null
            $retryAfter = $null
            try { $status = [int]$_.Exception.Response.StatusCode; $retryAfter = $_.Exception.Response.Headers['Retry-After'] } catch {}
            if ($attempt -ge $MaxRetries -or $status -notin @(408,429,500,502,503,504)) { throw }
            $sleep = if ($retryAfter) { [int]$retryAfter } else { [int][Math]::Pow(2,$attempt) + (Get-Random -Minimum 1 -Maximum 4) }
            Write-Info ('Retrying HTTP request in {0} seconds.' -f $sleep)
            Start-Sleep -Seconds $sleep
        }
    }
}

function Invoke-CWCollection {
    param([string]$Path, [string]$Conditions, [string]$Fields, [hashtable]$Headers, [int]$PageSize = 1000)
    $base = $env:ConnectWisePsa_ApiBaseUrl.TrimEnd('/')
    $all = New-Object System.Collections.ArrayList
    $page = 1
    while ($true) {
        $query = @('page={0}' -f $page, 'pageSize={0}' -f $PageSize)
        if ($Conditions) { $query += 'conditions={0}' -f [Uri]::EscapeDataString($Conditions) }
        if ($Fields) { $query += 'fields={0}' -f [Uri]::EscapeDataString($Fields) }
        $uri = '{0}/v4_6_release/apis/3.0/{1}?{2}' -f $base, $Path.TrimStart('/'), ($query -join '&')
        $result = @(Invoke-ApiWithRetry -Uri $uri -Headers $Headers)
        foreach ($item in $result) { [void]$all.Add($item) }
        if ($result.Count -lt $PageSize) { break }
        $page++
    }
    return @($all.ToArray())
}

function Invoke-CWItem {
    param([string]$Path, [hashtable]$Headers)
    $uri = '{0}/v4_6_release/apis/3.0/{1}' -f $env:ConnectWisePsa_ApiBaseUrl.TrimEnd('/'), $Path.TrimStart('/')
    return Invoke-ApiWithRetry -Uri $uri -Headers $Headers
}

function Get-CWCompany {
    param([string]$CustomerName, [hashtable]$Headers)
    $safe = ConvertTo-CWConditionString $CustomerName
    $companyMatches = @(Invoke-CWCollection -Path 'company/companies' -Conditions ('name="{0}"' -f $safe) -Fields 'id,name,identifier,status' -Headers $Headers)
    $exact = @($companyMatches | Where-Object { $_.name.Trim().ToLowerInvariant() -eq $CustomerName.Trim().ToLowerInvariant() })
    if ($exact.Count -eq 0) { Assert-BadRequest ('No exact ConnectWise company match found for "{0}".' -f $CustomerName) }
    if ($exact.Count -gt 1) { Assert-BadRequest ('Multiple exact ConnectWise company matches found for "{0}".' -f $CustomerName) }
    return $exact[0]
}

function Get-CWManagedServiceAgreements {
    param([int]$CompanyId, [datetime]$PeriodStart, [datetime]$PeriodEnd, [hashtable]$Headers)
    $conditions = 'company/id={0} AND startDate<=[{1}] AND (endDate=null OR endDate>=[{2}])' -f $CompanyId, $PeriodEnd.ToString('yyyy-MM-ddTHH:mm:ssZ'), $PeriodStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $agreements = @(Invoke-CWCollection -Path 'finance/agreements' -Conditions $conditions -Fields 'id,name,type,company,startDate,endDate,cancelledFlag,billTime,billExpenses,billProducts' -Headers $Headers)
    return @($agreements | Where-Object { $_.type.name -in $AllowedAgreementTypes -and -not $_.cancelledFlag })
}

function Get-CWOpenProjects {
    param([int]$CompanyId, [hashtable]$Headers)
    $conditions = 'company/id={0} AND closedFlag=false' -f $CompanyId
    return @(Invoke-CWCollection -Path 'project/projects' -Conditions $conditions -Fields 'id,name,status,type,company,description,manager,startDate,estimatedEndDate,actualEndDate,closedFlag' -Headers $Headers)
}

function Get-CWProjectCandidates {
    param([object[]]$Projects, [object]$Ticket, [int]$Maximum = 8)
    $sourceText = '{0} {1}' -f $Ticket.summary,$Ticket.initialDescription
    $terms = @($sourceText.ToLowerInvariant() -split '[^a-z0-9]+' | Where-Object { $_.Length -ge 4 } | Select-Object -Unique)
    $ranked = foreach ($project in $Projects) {
        $projectText = ('{0} {1} {2}' -f $project.name,$project.type.name,$project.description).ToLowerInvariant()
        $score = @($terms | Where-Object { $projectText.Contains($_) }).Count
        if ($score -gt 0) { [pscustomobject]@{ project=$project; score=$score } }
    }
    return @($ranked | Sort-Object score -Descending | Select-Object -First $Maximum | ForEach-Object { $_.project })
}

function Get-CWProjectPhases {
    param([int]$ProjectId, [hashtable]$Headers)
    try { return @(Invoke-CWCollection -Path ('project/projects/{0}/phases' -f $ProjectId) -Conditions $null -Fields 'id,project,name,description,status,billingMethod' -Headers $Headers) }
    catch { Write-ErrorLog ('Could not retrieve phases for project {0}: {1}' -f $ProjectId,$_.Exception.Message); return @() }
}

function Get-CWProjectTickets {
    param([int]$ProjectId, [hashtable]$Headers)
    try {
        return @(Invoke-CWCollection -Path 'project/tickets' -Conditions ('project/id={0}' -f $ProjectId) -Fields 'id,summary,project,phase,status,board,type,subType,item,company,actualHours,budgetHours,dateEntered,closedDate,closedFlag' -Headers $Headers)
    }
    catch { Write-ErrorLog ('Could not retrieve tickets for project {0}: {1}' -f $ProjectId,$_.Exception.Message); return @() }
}

function Get-CWServiceTickets {
    param([int]$CompanyId, [int[]]$AgreementIds, [datetime]$PeriodStart, [datetime]$PeriodEndExclusive, [hashtable]$Headers)
    if ($AgreementIds.Count -eq 0) { return @() }
    $tickets = New-Object System.Collections.ArrayList
    foreach ($agreementId in $AgreementIds) {
        # Candidate set includes tickets entered, closed, or last updated in the period.
        # Time-entry retrieval below adds older tickets worked during the audit month.
        $dateFilter = '((dateEntered>=[{0}] AND dateEntered<[{1}]) OR (closedDate>=[{0}] AND closedDate<[{1}]) OR (lastUpdated>=[{0}] AND lastUpdated<[{1}]))' -f $PeriodStart.ToString('yyyy-MM-ddTHH:mm:ssZ'), $PeriodEndExclusive.ToString('yyyy-MM-ddTHH:mm:ssZ')
        $conditions = 'company/id={0} AND agreement/id={1} AND {2}' -f $CompanyId,$agreementId,$dateFilter
        $items = @(Invoke-CWCollection -Path 'service/tickets' -Conditions $conditions -Fields 'id,summary,company,board,status,type,subType,item,priority,agreement,initialDescription,dateEntered,closedDate,closedFlag,lastUpdated,billingMethod,actualHours,budgetHours,owner,team' -Headers $Headers)
        foreach ($item in $items) { [void]$tickets.Add($item) }
    }
    return @($tickets.ToArray() | Sort-Object id -Unique)
}

function Get-CWTimeEntriesForPeriod {
    param([int]$CompanyId, [datetime]$Start, [datetime]$EndExclusive, [hashtable]$Headers)
    $conditions = 'company/id={0} AND timeStart>=[{1}] AND timeStart<[{2}]' -f $CompanyId,$Start.ToString('yyyy-MM-ddTHH:mm:ssZ'),$EndExclusive.ToString('yyyy-MM-ddTHH:mm:ssZ')
    try { return @(Invoke-CWCollection -Path 'time/entries' -Conditions $conditions -Fields 'id,company,chargeToId,chargeToType,timeStart,timeEnd,actualHours,billableOption,workType,workRole,notes,member,agreement' -Headers $Headers) }
    catch { Write-ErrorLog ('Time entry retrieval failed: {0}' -f $_.Exception.Message); return @() }
}

function Get-CWServiceTicketById {
    param([int]$TicketId,[hashtable]$Headers)
    return Invoke-CWItem -Path ('service/tickets/{0}' -f $TicketId) -Headers $Headers
}

function Get-CWTicketNotes {
    param([int]$TicketId,[hashtable]$Headers)
    try { return @(Invoke-CWCollection -Path ('service/tickets/{0}/notes' -f $TicketId) -Conditions $null -Fields 'id,text,detailDescriptionFlag,internalAnalysisFlag,resolutionFlag,dateCreated,createdBy' -Headers $Headers) }
    catch { Write-ErrorLog ('Notes retrieval failed for ticket {0}: {1}' -f $TicketId,$_.Exception.Message); return @() }
}

function Get-ManagedIdentityToken {
    param([string]$Resource)
    if (-not $env:IDENTITY_ENDPOINT -or -not $env:IDENTITY_HEADER) { return $null }
    $uri = '{0}?resource={1}&api-version=2019-08-01' -f $env:IDENTITY_ENDPOINT,[Uri]::EscapeDataString($Resource)
    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER }
    return $result.access_token
}

function Get-GraphAccessToken {
    $mi = Get-ManagedIdentityToken -Resource 'https://graph.microsoft.com/'
    if ($mi) { return $mi }
    $tenantId = if ($env:Graph_TenantId) { $env:Graph_TenantId } else { 'palaisparc.onmicrosoft.com' }
    $clientId = if ($env:Graph_ClientId) { $env:Graph_ClientId } else { 'c0d6f1f4-f988-47d3-874c-1ac5fc855463' }
    if ([string]::IsNullOrWhiteSpace($env:Graph_ClientSecret)) { throw 'Managed identity is unavailable and Graph_ClientSecret is not configured.' }
    $uri = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token' -f $tenantId
    $body = @{ client_id=$clientId; client_secret=$env:Graph_ClientSecret; scope='https://graph.microsoft.com/.default'; grant_type='client_credentials' }
    return (Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType 'application/x-www-form-urlencoded').access_token
}

function Invoke-GraphApi {
    param([string]$Uri, [string]$Method = 'Get', [string]$AccessToken, [object]$Body = $null)
    $headers = @{ Authorization = 'Bearer {0}' -f $AccessToken; Accept = 'application/json'; 'Content-Type' = 'application/json' }
    return Invoke-ApiWithRetry -Uri $Uri -Method $Method -Headers $headers -Body $Body
}

function Get-SharePointSite {
    param([string]$AccessToken)
    if ($env:SharePoint_SiteId) {
        Write-Info ('Using configured SharePoint_SiteId: {0}' -f $env:SharePoint_SiteId)
        return [pscustomobject]@{ id = $env:SharePoint_SiteId; webUrl = 'https://palaisparc.sharepoint.com/finance' }
    }
    $site = Invoke-GraphApi -Uri 'https://graph.microsoft.com/v1.0/sites/palaisparc.sharepoint.com:/finance' -AccessToken $AccessToken
    Write-Info ('Resolved SharePoint site ID: {0}' -f $site.id)
    return $site
}

function Get-SharePointList {
    param([string]$SiteId, [string]$ListName, [string]$ConfiguredListId, [string]$AccessToken)
    if ($ConfiguredListId) {
        Write-Info ('Using configured {0} list ID: {1}' -f $ListName,$ConfiguredListId)
        return [pscustomobject]@{ id = $ConfiguredListId; displayName = $ListName }
    }
    $filter = [Uri]::EscapeDataString("displayName eq '$($ListName.Replace("'", "''"))'")
    $result = Invoke-GraphApi -Uri ('https://graph.microsoft.com/v1.0/sites/{0}/lists?$filter={1}' -f $SiteId,$filter) -AccessToken $AccessToken
    $matches = @($result.value | Where-Object { $_.displayName -eq $ListName })
    if ($matches.Count -ne 1) { throw ('Expected exactly one SharePoint list named "{0}", found {1}.' -f $ListName,$matches.Count) }
    Write-Info ('Resolved {0} list ID: {1}' -f $ListName,$matches[0].id)
    return $matches[0]
}

function New-SharePointListItem {
    param([string]$SiteId, [string]$ListId, [hashtable]$Fields, [string]$AccessToken, [string]$AuditRunId, [string]$CustomerName, [string]$TicketId = '')
    Write-Info ('SharePoint create. AuditRunId={0}; CustomerName={1}; TicketId={2}' -f $AuditRunId,$CustomerName,$TicketId)
    $uri = 'https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items' -f $SiteId,$ListId
    # Queue redelivery uses lookup-and-update below, so an ambiguous POST is never replayed here.
    $headers = @{ Authorization = 'Bearer {0}' -f $AccessToken; Accept = 'application/json'; 'Content-Type' = 'application/json' }
    return Invoke-ApiWithRetry -Uri $uri -Method Post -Headers $headers -Body @{ fields = $Fields } -MaxRetries 1
}

function Update-SharePointListItem {
    param([string]$SiteId, [string]$ListId, [string]$ItemId, [hashtable]$Fields, [string]$AccessToken, [string]$AuditRunId, [string]$CustomerName, [string]$TicketId = '')
    Write-Info ('SharePoint update. AuditRunId={0}; CustomerName={1}; TicketId={2}' -f $AuditRunId,$CustomerName,$TicketId)
    $uri = 'https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items/{2}/fields' -f $SiteId,$ListId,$ItemId
    return Invoke-GraphApi -Uri $uri -Method Patch -AccessToken $AccessToken -Body $Fields
}

function Get-SharePointDetailItem {
    param([string]$SiteId, [string]$ListId, [string]$AuditRunId, [string]$TicketId, [string]$AccessToken)
    $filter = [Uri]::EscapeDataString("fields/AuditRunId eq '$($AuditRunId.Replace("'", "''"))' and fields/ConnectWiseTicketId eq '$($TicketId.Replace("'", "''"))'")
    $uri = 'https://graph.microsoft.com/v1.0/sites/{0}/lists/{1}/items?$expand=fields&$filter={2}' -f $SiteId,$ListId,$filter
    $result = Invoke-GraphApi -Uri $uri -AccessToken $AccessToken
    $items = @($result.value)
    if ($items.Count -gt 1) { throw ('Multiple detail records found for AuditRunId {0} and TicketId {1}.' -f $AuditRunId,$TicketId) }
    if ($items.Count -eq 1) { return $items[0] }
    return $null
}

function Save-SharePointAuditDetail {
    param([string]$SiteId, [string]$ListId, [hashtable]$Fields, [string]$AccessToken, [string]$AuditRunId, [string]$CustomerName, [string]$TicketId)
    $existing = Get-SharePointDetailItem -SiteId $SiteId -ListId $ListId -AuditRunId $AuditRunId -TicketId $TicketId -AccessToken $AccessToken
    if ($existing) {
        Update-SharePointListItem -SiteId $SiteId -ListId $ListId -ItemId ([string]$existing.id) -Fields $Fields -AccessToken $AccessToken -AuditRunId $AuditRunId -CustomerName $CustomerName -TicketId $TicketId
        return $existing
    }
    return New-SharePointListItem -SiteId $SiteId -ListId $ListId -Fields $Fields -AccessToken $AccessToken -AuditRunId $AuditRunId -CustomerName $CustomerName -TicketId $TicketId
}

function Invoke-AzureOpenAIClassification {
    param([object]$TicketContext)
    $systemPrompt = @'
You are the CloudFirst Finance Billing Classifier. Review one ConnectWise managed-services ticket against the customer's managed-services agreements and currently open projects.
Return JSON only. Do not use markdown. Do not invent IDs, project names, agreement scope, or facts.
Allowed classifications: Covered, ExistingProject, BillSeparately, ManualReview.
Decision rules:
1. ExistingProject when the documented work is part of a supplied open project's description, phase, or existing project-ticket pattern. Supply only a project ID present in the input.
2. Covered only when the evidence indicates ordinary support, maintenance, monitoring, repair, or administration of an existing managed environment and no supplied project is a stronger match.
3. BillSeparately when the evidence indicates new implementation, installation, migration, expansion, upgrade, custom build, procurement, training, or other separately chargeable work and no supplied project is a defensible match.
4. ManualReview when notes are incomplete, evidence conflicts, scope is unknown, more than one project is plausible, or confidence is below 0.80.
Use confidence from 0.00 through 1.00.
Output schema:
{"classification":"Covered|ExistingProject|BillSeparately|ManualReview","confidence":0.0,"recommendedProjectId":null,"recommendedProjectName":null,"billSeparately":false,"reasonCodes":[],"explanation":"","missingInformation":[]}
'@
    $userJson = $TicketContext | ConvertTo-Json -Depth 60 -Compress
    $messages = @(@{ role='system'; content=$systemPrompt }, @{ role='user'; content=$userJson })
    $requestBody = @{ messages=$messages; response_format=@{ type='json_object' } }
    # GPT-5 deployments use max_completion_tokens and omit temperature.
    $requestBody.max_completion_tokens = 1200

    $endpoint = $env:AzureOpenAI_Endpoint.TrimEnd('/')
    $deployment = $env:AzureOpenAI_Deployment
    $version = if ($env:AzureOpenAI_ApiVersion) { $env:AzureOpenAI_ApiVersion } else { '2024-10-21' }
    $uri = '{0}/openai/deployments/{1}/chat/completions?api-version={2}' -f $endpoint,[Uri]::EscapeDataString($deployment),$version
    $headers = @{ 'Content-Type'='application/json' }
    if ($env:AzureOpenAI_ApiKey) { $headers['api-key'] = $env:AzureOpenAI_ApiKey }
    else {
        $token = Get-ManagedIdentityToken -Resource 'https://cognitiveservices.azure.com/'
        if (-not $token) { throw 'Azure OpenAI authentication is not configured. Set AzureOpenAI_ApiKey or enable managed identity.' }
        $headers.Authorization = 'Bearer {0}' -f $token
    }
    $result = Invoke-ApiWithRetry -Uri $uri -Method Post -Headers $headers -Body $requestBody
    $text = [string]$result.choices[0].message.content
    try { return $text | ConvertFrom-Json }
    catch { throw ('Azure OpenAI did not return valid JSON. Raw content: {0}' -f $text) }
}

function New-SharePointAuditRunFields {
    param([object]$Audit,[string]$RequestedBy,[int]$AgreementsFound,[int]$ProjectsFound,[int]$TicketsFound)
    return @{
        Title                  = 'Billing Audit {0}' -f $Audit.auditRunId
        AuditRunId             = $Audit.auditRunId
        CustomerName           = $Audit.companyName
        ConnectWiseCompanyId   = [string]$Audit.companyId
        AuditMonth             = [int]$Audit.auditMonth
        AuditYear              = [int]$Audit.auditYear
        PeriodStart            = $Audit.periodStart
        PeriodEnd              = $Audit.periodEnd
        RequestedBy            = $RequestedBy
        StartedOn              = (Get-Date).ToUniversalTime().ToString('o')
        RunStatus              = 'Running'
        AgreementsFound        = $AgreementsFound
        ProjectsFound          = $ProjectsFound
        TicketsFound           = $TicketsFound
        CorrelationId          = $CorrelationId
    }
}

function New-SharePointAuditFields {
    param([object]$Audit,[object]$Ticket,[object]$Classification,[string]$RequestedBy,[object[]]$TicketTime)
    # IMPORTANT: Rename these keys if your SharePoint list uses different internal column names.
    $prefix = if ($env:SharePoint_ListItemTitlePrefix) { $env:SharePoint_ListItemTitlePrefix } else { 'Billing Audit' }
    return @{
        Title                    = '{0} {1} - Ticket {2}' -f $prefix,$Audit.auditRunId,$Ticket.id
        AuditRunId               = $Audit.auditRunId
        CustomerName             = $Audit.companyName
        ConnectWiseCompanyId     = [string]$Audit.companyId
        AuditMonth               = [int]$Audit.auditMonth
        AuditYear                = [int]$Audit.auditYear
        PeriodStart              = $Audit.periodStart
        PeriodEnd                = $Audit.periodEnd
        RequestedBy              = $RequestedBy
        ConnectWiseTicketId      = [string]$Ticket.id
        TicketSummary            = [string]$Ticket.summary
        BoardName                = [string]$Ticket.board.name
        TicketStatus             = [string]$Ticket.status.name
        CurrentAgreementId       = [string]$Ticket.agreement.id
        CurrentAgreementName     = [string]$Ticket.agreement.name
        CurrentProject           = [string]$Ticket.project.name
        HoursWorked              = [double](@($TicketTime | ForEach-Object { [double]$_.actualHours } | Measure-Object -Sum).Sum)
        BillableHours            = [double](@($TicketTime | Where-Object { $_.billableOption -match 'Billable' } | ForEach-Object { [double]$_.actualHours } | Measure-Object -Sum).Sum)
        Classification           = [string]$Classification.classification
        Confidence               = [double]$Classification.confidence
        RecommendedProjectId     = if ($null -ne $Classification.recommendedProjectId) { [string]$Classification.recommendedProjectId } else { $null }
        RecommendedProjectName   = [string]$Classification.recommendedProjectName
        RecommendBillSeparately  = [bool]$Classification.billSeparately
        ReasonCodes              = (@($Classification.reasonCodes) -join '; ')
        AIExplanation            = [string]$Classification.explanation
        MissingInformation       = (@($Classification.missingInformation) -join '; ')
        ReviewStatus             = if ($Classification.classification -eq 'ManualReview') { 'Manual Review Required' } else { 'AI Review Complete' }
        ActionStatus             = 'No Action Taken'
        FinanceDecision          = 'Pending'
        ExceptionFlag            = ($Classification.classification -eq 'ManualReview')
        GPTPromptVersion         = 'BillingAudit-v1'
        CorrelationId            = $CorrelationId
    }
}

try {
    Write-Info 'Starting managed-services billing audit.'
    $job = if ($QueueItem -is [string]) { $QueueItem | ConvertFrom-Json -ErrorAction Stop } else { $QueueItem }
    $customerName = [string]$job.customerName; $month = [int]$job.auditMonth; $year = [int]$job.auditYear
    $requestedBy = [string]$job.requestedBy; $auditRunId = [string]$job.auditRunId; $runSharePointItemId = [string]$job.runSharePointItemId
    if ([string]::IsNullOrWhiteSpace($customerName) -or $month -notin 1..12 -or $year -lt 2000 -or [string]::IsNullOrWhiteSpace($auditRunId) -or [string]::IsNullOrWhiteSpace($runSharePointItemId)) { throw 'Queue message is missing required billing-audit job data.' }

    Assert-Settings @(
        'ConnectWisePsa_ApiBaseUrl','ConnectWisePsa_ApiCompanyId','ConnectWisePsa_ApiPublicKey',
        'ConnectWisePsa_ApiPrivateKey','ConnectWisePsa_ApiClientId','AzureOpenAI_Endpoint',
        'AzureOpenAI_Deployment'
    )

    $periodStart = [datetime]::SpecifyKind([datetime]::new($year,$month,1),[DateTimeKind]::Utc)
    $periodEndExclusive = $periodStart.AddMonths(1)
    $periodEnd = $periodEndExclusive.AddSeconds(-1)
    $dryRun = ($env:BillingAudit_DryRun -eq '1')
    $maxTickets = if ($env:BillingAudit_MaxTickets) { [int]$env:BillingAudit_MaxTickets } else { 250 }
    $maxNotesChars = if ($env:BillingAudit_MaxNotesCharacters) { [int]$env:BillingAudit_MaxNotesCharacters } else { 24000 }

    $cwHeaders = New-ConnectWiseHeaders
    $company = Get-CWCompany -CustomerName $customerName -Headers $cwHeaders
    Write-Info ('Resolved company {0}, ID {1}.' -f $company.name,$company.id)

    $agreements = @(Get-CWManagedServiceAgreements -CompanyId $company.id -PeriodStart $periodStart -PeriodEnd $periodEnd -Headers $cwHeaders)
    if ($agreements.Count -eq 0) { Assert-BadRequest ('No active agreements of an approved managed-services type were found for {0}.' -f $company.name) }

    $projects = @(Get-CWOpenProjects -CompanyId $company.id -Headers $cwHeaders)
    $projectContexts = @($projects | ForEach-Object { [pscustomobject]@{ id=$_.id; name=$_.name; status=$_.status; type=$_.type; description=$_.description; manager=$_.manager } })

    $agreementIds = @($agreements | ForEach-Object { [int]$_.id })
    $serviceTickets = @(Get-CWServiceTickets -CompanyId $company.id -AgreementIds $agreementIds -PeriodStart $periodStart -PeriodEndExclusive $periodEndExclusive -Headers $cwHeaders)
    $timeEntries = @(Get-CWTimeEntriesForPeriod -CompanyId $company.id -Start $periodStart -EndExclusive $periodEndExclusive -Headers $cwHeaders)

    # Add older service tickets that received time during the month.
    $knownIds = @{}; foreach ($ticket in $serviceTickets) { $knownIds[[string]$ticket.id] = $true }
    $agreementSet = @{}; foreach ($id in $agreementIds) { $agreementSet[[string]$id] = $true }
    foreach ($entry in $timeEntries) {
        if ($entry.chargeToType -match 'Service' -and $entry.chargeToId) {
            $tid = [string]$entry.chargeToId
            if (-not $knownIds.ContainsKey($tid)) {
                try {
                    $ticket = Get-CWServiceTicketById -TicketId ([int]$tid) -Headers $cwHeaders
                    if ($ticket.company.id -eq $company.id -and $ticket.agreement.id -and $agreementSet.ContainsKey([string]$ticket.agreement.id)) {
                        $serviceTickets += $ticket; $knownIds[$tid] = $true
                    }
                } catch { Write-ErrorLog ('Could not add ticket {0} from time entry: {1}' -f $tid,$_.Exception.Message) }
            }
        }
    }
    $serviceTickets = @($serviceTickets | Sort-Object id -Unique)
    if ($serviceTickets.Count -gt $maxTickets) { throw ('Audit contains {0} tickets, exceeding BillingAudit_MaxTickets={1}.' -f $serviceTickets.Count,$maxTickets) }

    $timeEntriesByTicket = @{}
    foreach ($entry in $timeEntries) {
        if ($entry.chargeToType -match 'Service' -and $entry.chargeToId) {
            $ticketId = [string]$entry.chargeToId
            if (-not $timeEntriesByTicket.ContainsKey($ticketId)) { $timeEntriesByTicket[$ticketId] = [System.Collections.ArrayList]::new() }
            [void]$timeEntriesByTicket[$ticketId].Add($entry)
        }
    }

    $graphToken = if ($dryRun) { $null } else { Get-GraphAccessToken }
    $savedItems = @(); $errors = @(); $classifications = @()
    $audit = [pscustomobject]@{ auditRunId=$auditRunId; companyId=$company.id; companyName=$company.name; auditMonth=$month; auditYear=$year; periodStart=$periodStart.ToString('o'); periodEnd=$periodEnd.ToString('o') }
    $sharePointSite = $null; $runsList = $null; $detailsList = $null
    if (-not $dryRun) {
        $sharePointSite = Get-SharePointSite -AccessToken $graphToken
        $runsList = Get-SharePointList -SiteId $sharePointSite.id -ListName 'Billing Audit Runs' -ConfiguredListId $env:SharePoint_BillingAuditRunsListId -AccessToken $graphToken
        $detailsList = Get-SharePointList -SiteId $sharePointSite.id -ListName 'Billing Audit Details' -ConfiguredListId $env:SharePoint_BillingAuditDetailsListId -AccessToken $graphToken
        Update-SharePointListItem -SiteId $sharePointSite.id -ListId $runsList.id -ItemId $runSharePointItemId -AccessToken $graphToken -AuditRunId $auditRunId -CustomerName $company.name -Fields @{
            AgreementsFound=$agreements.Count; ProjectsFound=$projectContexts.Count; TicketsFound=$serviceTickets.Count; RunStatus='Running'
        }
    }

    foreach ($ticket in $serviceTickets) {
        try {
            Write-Info ('Reviewing ticket {0}: {1}' -f $ticket.id,$ticket.summary)
            $notes = @(Get-CWTicketNotes -TicketId $ticket.id -Headers $cwHeaders)
            $ticketTime = if ($timeEntriesByTicket.ContainsKey([string]$ticket.id)) { @($timeEntriesByTicket[[string]$ticket.id]) } else { @() }
            $notesText = (($notes | ForEach-Object { '[{0}] {1}' -f $_.dateCreated,$_.text }) -join "`n")
            if ($notesText.Length -gt $maxNotesChars) { $notesText = $notesText.Substring(0,$maxNotesChars) }

            $context = [pscustomobject]@{
                audit = $audit
                managedServiceAgreements = @($agreements | ForEach-Object { [pscustomobject]@{ id=$_.id; name=$_.name; type=$_.type.name; startDate=$_.startDate; endDate=$_.endDate; billTime=$_.billTime; billExpenses=$_.billExpenses; billProducts=$_.billProducts } })
                openProjects = @(Get-CWProjectCandidates -Projects $projectContexts -Ticket $ticket)
                ticket = [pscustomobject]@{
                    id=$ticket.id; summary=$ticket.summary; initialDescription=$ticket.initialDescription
                    board=$ticket.board; status=$ticket.status; type=$ticket.type; subType=$ticket.subType; item=$ticket.item
                    agreement=$ticket.agreement; dateEntered=$ticket.dateEntered; closedDate=$ticket.closedDate
                    actualHours=$ticket.actualHours; billingMethod=$ticket.billingMethod
                    notesText=$notesText; timeEntries=$ticketTime
                }
            }

            $classification = if ($dryRun) {
                [pscustomobject]@{ classification='ManualReview'; confidence=0; recommendedProjectId=$null; recommendedProjectName=$null; billSeparately=$false; reasonCodes=@('DRY_RUN'); explanation='Dry run. GPT was not called.'; missingInformation=@() }
            } else { Invoke-AzureOpenAIClassification -TicketContext $context }

            if ($classification.classification -notin @('Covered','ExistingProject','BillSeparately','ManualReview')) { throw 'Classifier returned an unsupported classification.' }
            if ($classification.classification -eq 'ExistingProject') {
                $validProject = @($projectContexts | Where-Object { [string]$_.id -eq [string]$classification.recommendedProjectId })
                if ($validProject.Count -ne 1) {
                    $classification.classification='ManualReview'; $classification.confidence=0
                    $classification.recommendedProjectId=$null; $classification.recommendedProjectName=$null
                    $classification.billSeparately=$false
                    $classification.explanation='The model recommended a project ID that was not in the supplied open-project list.'
                    $classification.reasonCodes=@('INVALID_PROJECT_REFERENCE')
                }
            }

            $classifications += [pscustomobject]@{ ticketId=$ticket.id; result=$classification }
            if (-not $dryRun) {
                $fields = New-SharePointAuditFields -Audit $audit -Ticket $ticket -Classification $classification -RequestedBy $requestedBy -TicketTime $ticketTime
                $saved = Save-SharePointAuditDetail -SiteId $sharePointSite.id -ListId $detailsList.id -Fields $fields -AccessToken $graphToken -AuditRunId $auditRunId -CustomerName $company.name -TicketId ([string]$ticket.id)
                $savedItems += [pscustomobject]@{ ticketId=$ticket.id; sharePointItemId=$saved.id }
            }
        }
        catch {
            Write-ErrorLog ('Ticket {0} failed: {1}' -f $ticket.id,$_.Exception.Message)
            $errors += [pscustomobject]@{ ticketId=$ticket.id; error=$_.Exception.Message }
        }
    }

    $counts = @{}
    foreach ($name in @('Covered','ExistingProject','BillSeparately','ManualReview')) {
        $counts[$name] = @($classifications | Where-Object { $_.result.classification -eq $name }).Count
    }
    $runStatus = if ($errors.Count -eq 0) { 'Completed' } else { 'Completed With Warnings' }
    if (-not $dryRun) {
        $errorSummary = (@($errors | ForEach-Object { 'Ticket {0}: {1}' -f $_.ticketId,$_.error }) -join ' | ')
        Update-SharePointListItem -SiteId $sharePointSite.id -ListId $runsList.id -ItemId $runSharePointItemId -AccessToken $graphToken -AuditRunId $auditRunId -CustomerName $company.name -Fields @{
            CompletedOn = (Get-Date).ToUniversalTime().ToString('o'); RunStatus = $runStatus
            TicketsClassified = $classifications.Count; CoveredCount = $counts['Covered']; ExistingProjectCount = $counts['ExistingProject']
            BillSeparatelyCount = $counts['BillSeparately']; ManualReviewCount = $counts['ManualReview']; FailedTicketCount = $errors.Count; ErrorSummary = $errorSummary
        }
    }

    Send-JsonResponse -StatusCode ([HttpStatusCode]::OK) -Body @{
        success=$true; auditRunId=$auditRunId; runSharePointItemId=$runSharePointItemId
        ticketsReviewed=$serviceTickets.Count; covered=$counts['Covered']; project=$counts['ExistingProject']; billSeparately=$counts['BillSeparately']; manualReview=$counts['ManualReview']; failed=$errors.Count
    }
}
catch {
    $message = $_.Exception.Message
    Write-ErrorLog $message
    if ($runSharePointItemId -and $sharePointSite -and $runsList -and $graphToken) {
        try {
            Update-SharePointListItem -SiteId $sharePointSite.id -ListId $runsList.id -ItemId $runSharePointItemId -AccessToken $graphToken -AuditRunId $auditRunId -CustomerName $company.name -Fields @{
                CompletedOn = (Get-Date).ToUniversalTime().ToString('o'); RunStatus = 'Failed'
                FailedTicketCount = if ($errors) { $errors.Count } else { 0 }
                ErrorSummary = $message
            }
        }
        catch { Write-ErrorLog ('Could not mark audit run {0} as failed: {1}' -f $auditRunId,$_.Exception.Message) }
    }
    $status = if ($_.Exception -is [System.ArgumentException]) { [HttpStatusCode]::BadRequest } else { [HttpStatusCode]::InternalServerError }
    Send-JsonResponse -StatusCode $status -Body @{ success=$false; correlationId=$CorrelationId; error=$message }
}
