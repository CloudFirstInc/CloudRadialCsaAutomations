# CloudFirst ConnectWise Activity Type Sync - Draft

## Purpose
Timer-triggered PowerShell Azure Function. SharePoint list **ConnectWise Activity Types** is authoritative. The function creates/updates ConnectWise Activity Types, synchronizes Point Value and Active state, writes ConnectWise IDs back to SharePoint, and can hide current ConnectWise types absent from SharePoint.

## Required Function App settings
- ConnectWisePsa_ApiBaseUrl — HTTPS ConnectWise host, such as `https://api-na.myconnectwise.net`; the full `/v4_6_release/apis/3.0` suffix is also accepted. Do not include other path segments, query strings, or fragments.
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
- ActivityTypeSync_ConnectWiseResource = `sales/activities/types`
- ActivityTypeSync_CWNameField = `name`
- ActivityTypeSync_CWPointsField = `points`
- ActivityTypeSync_CWInactiveField = `inactiveFlag`
- ActivityTypeSync_HideUnmanaged = `true`
- ActivityTypeSync_ProtectedNames = `Historical Entry`

## Rollout sequence
1. Deploy. Each run creates and updates ConnectWise Activity Types, writes ConnectWise IDs and Sync Status back to SharePoint, and hides ConnectWise types not represented in SharePoint.
2. Verify newly created/updated types and SharePoint ConnectWise IDs.
3. Add any ConnectWise types that must stay active to `ActivityTypeSync_ProtectedNames` or `ActivityTypeSync_ProtectedIds`, or set `ActivityTypeSync_HideUnmanaged=false` to disable hiding.

## Important behavior
- Deletion is not used. Unneeded types are hidden with the configured inactive field.
- An empty SharePoint list or empty ConnectWise response causes a hard stop.
- SharePoint matches by ConnectWise ID first, then exact name.
- If any row fails to sync, the hide-unmanaged pass is skipped and the run fails.
