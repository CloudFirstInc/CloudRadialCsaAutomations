# CloudFirst ConnectWise Activity Type Sync - Draft

## Purpose
Timer-triggered PowerShell Azure Function. SharePoint list **ConnectWise Activity Types** is authoritative. The function creates/updates ConnectWise Activity Types, synchronizes Point Value and Active state, writes ConnectWise IDs back to SharePoint, and can hide current ConnectWise types absent from SharePoint.

## Required Function App settings
- ConnectWisePsa_ApiBaseUrl
- ConnectWisePsa_ApiCompanyId
- ConnectWisePsa_ApiPublicKey
- ConnectWisePsa_ApiPrivateKey
- ConnectWisePsa_ApiClientId
- Ms365_TenantId
- Ms365_AuthAppId
- Ms365_CertBase64
- Ms365_CertPassword
- ActivityTypeSync_Schedule, for example `0 */15 * * * *`

## Defaults
- ActivityTypeSync_SharePointSiteUrl = `https://palaisparc.sharepoint.com/IT/connectwise`
- ActivityTypeSync_SharePointListName = `ConnectWise Activity Types`
- ActivityTypeSync_ConnectWiseResource = `sales/activityTypes`
- ActivityTypeSync_CWNameField = `name`
- ActivityTypeSync_CWPointsField = `pointsValue`
- ActivityTypeSync_CWInactiveField = `inactiveFlag`
- ActivityTypeSync_ApplyChanges = `false`
- ActivityTypeSync_HideUnmanaged = `false`
- ActivityTypeSync_ProtectedNames = `Historical Entry`

## Rollout sequence
1. Deploy with `ActivityTypeSync_ApplyChanges=false` and `ActivityTypeSync_HideUnmanaged=false`.
2. Run manually and inspect GET results, audit logs, endpoint, field names, and proposed create/update actions.
3. Confirm the Activity Type POST/PATCH schema in the ConnectWise Developer Network for the API version used by CloudFirst.
4. Enable `ActivityTypeSync_ApplyChanges=true` while keeping `ActivityTypeSync_HideUnmanaged=false`.
5. Verify newly created/updated types and SharePoint ConnectWise IDs.
6. Only after reviewing all unmatched current ConnectWise types, enable `ActivityTypeSync_HideUnmanaged=true`.

## Important behavior
- Deletion is not used. Unneeded types are hidden with the configured inactive field.
- An empty SharePoint list or empty ConnectWise response causes a hard stop.
- SharePoint matches by ConnectWise ID first, then exact name.
- In audit mode no SharePoint Sync Status or ConnectWise ID fields are changed.
