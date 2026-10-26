# CloudFirst Opportunity Upsert Azure Function

## Purpose

This PowerShell 7.2 HTTP-triggered Azure Function creates or updates a ConnectWise PSA opportunity, creates or verifies its SharePoint opportunity folder, refreshes the folder metadata, and writes the SharePoint folder URL to ConnectWise Opportunity custom field ID `76` (`Sharepoint Link`).

## Files

Place these files together in the Azure Function folder:

- `run.ps1` - use `CloudFirst-Opportunity-Upsert-AzureFunction.ps1` as the contents of `run.ps1`
- `function.json` - HTTP trigger and response bindings

The Function App must also include `PnP.PowerShell` in the app-level `requirements.psd1`:

```powershell
@{
    'PnP.PowerShell' = '2.*'
}
```

## HTTP endpoint

The route configured by `function.json` is:

```text
POST /api/opportunities/upsert
```

The trigger uses `authLevel: function`, so callers must provide a valid Azure Function key.

## SharePoint routing

### New client or prospect

- Site: `https://palaisparc.sharepoint.com/sales`
- Library: `Company Opportunities`

### Existing client

- Site: `https://palaisparc.sharepoint.com/sales/AccountManagement`
- Library: `Client Opportunities`

## Folder naming

The function uses a flat opportunity-folder structure:

```text
<Opportunity Number> - <Opportunity Name>
```

Example:

```text
4832 - City of Rye Network Refresh
```

## SharePoint metadata

The function verifies or creates these library fields and updates them on the opportunity folder:

- `OpportunityName` - Opportunity Name
- `OpportunityNumber` - Opportunity Number
- `OpportunityType` - Opportunity Type
- `CompanyName` - Company Name
- `CompanyNumber` - ConnectWise Company Number
- `ConnectWiseLink` - ConnectWise Link
- `OpportunityBriefDescription` - Brief Description of the Opportunity

## ConnectWise custom field

The SharePoint opportunity-folder URL is read from and written to ConnectWise Opportunity custom field ID `76`. The default caption is `Sharepoint Link`; the ID is authoritative and caption matching is used as a fallback. Both values can be overridden with optional Function App settings:

```text
OpportunityWorkspace_CustomFieldId = 76
OpportunityWorkspace_CustomFieldCaption = Sharepoint Link
```

The custom field should be a hyperlink field. The script updates its `value` and does not create the ConnectWise custom field.

## Required Function App settings

```text
ConnectWisePsa_ApiBaseUrl
ConnectWisePsa_ApiCompanyId
ConnectWisePsa_ApiPublicKey
ConnectWisePsa_ApiPrivateKey
ConnectWisePsa_ApiClientId
Ms365_TenantId
Ms365_AuthAppId
Ms365_CertBase64
Ms365_CertPassword
```

`ConnectWisePsa_ApiBaseUrl` can be either the ConnectWise host or the complete API base. Examples:

```text
https://api-na.myconnectwise.net
```

```text
https://api-na.myconnectwise.net/v4_6_release/apis/3.0
```

## Optional Function App settings

```text
ConnectWisePsa_ApiVersion
OpportunityWorkspace_NewClientTypeRegex
OpportunityWorkspace_ExistingClientTypeRegex
OpportunityWorkspace_DefaultCategory
OpportunityWorkspace_DryRun
```

Defaults used by the function:

```text
ConnectWisePsa_ApiVersion = 2022.1
OpportunityWorkspace_NewClientTypeRegex = (?i)new client|new customer|prospect
OpportunityWorkspace_ExistingClientTypeRegex = (?i)existing client|existing customer|account management|expansion|renewal
OpportunityWorkspace_DefaultCategory = blank
OpportunityWorkspace_DryRun = false
OpportunityWorkspace_CustomFieldId = 76
OpportunityWorkspace_CustomFieldCaption = Sharepoint Link
```

The custom-field settings have working defaults and only need to be configured to use a different field:

```text
OpportunityWorkspace_CustomFieldId
OpportunityWorkspace_CustomFieldCaption
```

These legacy routing settings are not used; SharePoint routing is fixed in the function:

```text
OpportunityWorkspace_SalesSiteUrl
OpportunityWorkspace_SalesParentFolder
OpportunityWorkspace_AccountManagementSiteUrl
OpportunityWorkspace_AccountManagementParentFolder
```

## Microsoft 365 application permissions

The Entra application represented by `Ms365_AuthAppId` must have permission to both SharePoint sites and must be able to create folders, create library fields when missing, and update folder list-item metadata. The function authenticates using the PFX certificate stored in `Ms365_CertBase64` and `Ms365_CertPassword`.

## Create request

```json
{
  "mode": "Create",
  "workspaceCategory": "NewClient",
  "briefDescription": "Managed services opportunity for a new prospect.",
  "connectWiseLink": "https://example.invalid/connectwise/opportunity/1234",
  "opportunity": {
    "name": "Managed Services Opportunity",
    "company": {
      "id": 123
    },
    "contact": {
      "id": 456
    },
    "primarySalesRep": {
      "id": 10
    },
    "stage": {
      "id": 2
    },
    "status": {
      "id": 1
    },
    "type": {
      "id": 7
    }
  }
}
```

## Update request using an opportunity object

Object updates are merged into the current opportunity before sending the complete record to ConnectWise. Unspecified fields are preserved; nested objects are merged recursively and arrays are replaced as a whole. `opportunity.id`, if supplied, must match `opportunityId`.

```json
{
  "mode": "Update",
  "opportunityId": 4832,
  "workspaceCategory": "ExistingClient",
  "briefDescription": "Network modernization and firewall replacement.",
  "opportunity": {
    "id": 4832,
    "name": "City of Rye Network Refresh",
    "company": {
      "id": 321
    },
    "stage": {
      "id": 4
    },
    "status": {
      "id": 1
    },
    "type": {
      "id": 9
    }
  }
}
```

## Update request using JSON Patch

Alternatively, supply a JSON Patch array. Supported operations are `add`, `remove`, `replace`, `move`, `copy`, and `test`, subject to ConnectWise API support. A patch that changes the opportunity type must include an explicit `workspaceCategory`.

```json
{
  "mode": "Update",
  "opportunityId": 4832,
  "workspaceCategory": "ExistingClient",
  "opportunity": [
    {
      "op": "replace",
      "path": "name",
      "value": "City of Rye Network Refresh"
    }
  ]
}
```

## Workspace category behavior

The safest approach is to provide one of these values explicitly:

```text
NewClient
ExistingClient
```

If `workspaceCategory` is omitted, the function attempts to classify the ConnectWise Opportunity Type using the configured regular expressions. If it cannot classify the opportunity and no default category is configured, it returns an error instead of guessing. Create requests are classified and validated before the ConnectWise opportunity is created; provide an explicit category when the request only contains a type ID and classification is otherwise unavailable.

## Processing behavior

### Create

1. Validates the request and resolves the SharePoint category.
2. Creates the ConnectWise opportunity.
3. Retrieves the assigned opportunity ID.
4. Selects the Sales or Account Management SharePoint destination.
5. Creates or reuses the flat opportunity folder.
6. Updates the folder metadata.
7. Writes the folder URL to ConnectWise custom field ID 76.

The Create POST is not automatically retried to avoid duplicate opportunities when a response is lost. If a previous run created the opportunity folder but did not save its URL, a later run can reuse the matching unlinked folder.

### Update

1. Updates the ConnectWise opportunity.
2. Reloads the complete opportunity.
3. Reads custom field ID 76.
4. If a SharePoint URL exists, verifies that folder and refreshes its metadata.
5. If the field is blank, creates the opportunity folder and writes its URL to field ID 76.

## Successful response

```json
{
  "success": true,
  "operation": "Created",
  "dryRun": false,
  "opportunityId": 4832,
  "opportunityName": "City of Rye Network Refresh",
  "correlationId": "a1b2c3d4-e5f6-4789-a012-3456789abcde",
  "connectWiseCustomFieldId": "76",
  "workspaceCategory": "ExistingClient",
  "workspaceUrl": "https://palaisparc.sharepoint.com/...",
  "workspaceCreated": true,
  "workspacePath": "/sales/AccountManagement/Client Opportunities/4832 - City of Rye Network Refresh",
  "sharePointMetadata": {},
  "note": "Folder ensured and metadata updated."
}
```

Callers may send a valid UUID in the `x-correlation-id` header; otherwise the function generates one. The ID is included in the response and logs. Invalid request input returns HTTP 400, downstream ConnectWise or SharePoint failures return HTTP 502, and unexpected internal errors return HTTP 500. ConnectWise GET and PUT requests use at most three attempts with 30-second request timeouts and short backoff; POST and PATCH requests are not retried.

## Deployment notes

1. Create or open a PowerShell 7.2 Azure Function.
2. Place the Azure Function script in the function directory as `run.ps1`.
3. Place `function.json` in the same directory.
4. Add `PnP.PowerShell` to the Function App's root `requirements.psd1`.
5. Add the required Function App settings.
6. Restart the Function App after changing application settings or managed dependencies.
7. Test with a known ConnectWise opportunity before allowing production automation to call the endpoint.

## Security

Do not hard-code ConnectWise private keys, certificate passwords, or the PFX certificate into `run.ps1`. Keep those values in Azure Function App settings or an approved secret store. Rotate credentials that have previously been exposed in scripts, terminals, chat, or source control.
