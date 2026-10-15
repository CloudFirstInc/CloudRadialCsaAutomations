# ConnectWise Opportunity Package Azure Function

PowerShell Azure Function that:

1. Reads a ConnectWise Manage opportunity.
2. Retrieves optional related notes, products, team, forecast, activities, and document metadata.
3. Creates a SharePoint opportunity folder and an Attachments subfolder.
4. Downloads ConnectWise documents and uploads them to SharePoint.
5. Saves `opportunity-data.json` in the opportunity folder.
6. Returns the SharePoint folder links and consolidated opportunity data.

## Required application settings

- `ConnectWisePsa_ApiBaseUrl`
- `ConnectWisePsa_ApiCompanyId`
- `ConnectWisePsa_ApiPublicKey`
- `ConnectWisePsa_ApiPrivateKey`
- `ConnectWisePsa_ApiClientId`
- `M365_TenantId`
- `M365_AuthAppId`
- `M365_AuthSecretId` (the client secret **value** for the app registration, not its secret ID)
- Either `SharePoint_DriveId`, or both `SharePoint_SiteId` and `SP_DocLibraryName` (for example, `Documents`)

## Optional application settings

- `SecurityKey`
- `DebugLogging=1`
- `SharePoint_RootFolderPath`, for example `Sales/Opportunities`; when omitted, opportunity folders are created at the document library root
- `SharePoint_RootItemId`, defaults to `root`
- `ConnectWisePsa_OpportunityRecordType`, defaults to trying `Opportunity` and `SalesOpportunity`

## SharePoint app permission

Use Microsoft Graph application permissions and grant the app access to the target site/library. Prefer `Sites.Selected` with write access granted only to the intended SharePoint site. The function obtains a Graph token using the `M365_` client ID, client secret value, and tenant ID settings. If `SharePoint_DriveId` is not configured, it resolves the drive by matching `SP_DocLibraryName` against the drives listed for `SharePoint_SiteId`.

## Request

```json
{
  "OpportunityId": 12345,
  "SecurityKey": "optional",
  "IncludeAttachments": true,
  "OverwriteExistingFiles": false
}
```

The function also accepts `OpportunityId` as a query parameter.

Existing attachment files are skipped by default. If `OverwriteExistingFiles` is false, the upload also uses Graph's `fail` conflict behavior so a file created concurrently is not overwritten. Optional ConnectWise related-resource requests are treated as empty only when ConnectWise returns 404; other failures fail the function request.

## Important validation

ConnectWise document association and download behavior can vary by PSA version and configuration. The function tries both `Opportunity` and `SalesOpportunity` record types. Validate the document query and `/system/documents/{id}/download` route in your ConnectWise Developer Network/API Explorer before production deployment.

The included upload method is intended for attachments supported by the Microsoft Graph single-call upload API. If your ConnectWise attachments can exceed the Graph single-upload limit, replace `Upload-GraphBytes` with an upload-session implementation.
