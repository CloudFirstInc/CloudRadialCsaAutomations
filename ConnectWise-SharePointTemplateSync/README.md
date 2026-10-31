# ConnectWise-SharePointTemplateSync

One-way SharePoint -> ConnectWise push for recurring service templates and schedules.
A Power Automate flow calls this function when an item in a `CW- <Dept> - Templates` or `CW - <Dept> - Schedules` list is created, modified or deleted.

## Status: write paths verified live (2026-10-30), with limits

A live test against ConnectWise (company "Cloudfirst Inc Test", all test records removed afterwards) found:

| Operation | Result |
|---|---|
| `POST service/templates` | Works, but **always creates a master** (`templateFlag=true`, no company), ignoring `company` and `templateFlag`. |
| `PATCH service/templates/{id}` on a company template | Not yet verified live (only masters were tested, and `company`/`templateFlag` are silently ignored). |
| `POST service/templates/{id}/resources` with `recurrence` | Works. `endDateTime` is computed as start + hours. |
| `PUT` / `PATCH` on a resource | Works (PATCH needs a JSON array, even for one operation). |
| End a recurrence / `DELETE` a resource | **Not possible**: no end field on `recurrence`, and resource DELETE returns 405. |
| Response to POST | 201 with the object in the body and an `http://` Location header; the function uses an `HttpClient` that does not follow redirects. |

Consequences:
- The API cannot create **company-level** templates. For a new template the function creates (or finds) the master, sets `PushStatus=AwaitingCompanyApply`, and stops. Apply the master to the company in ConnectWise, then re-save the SharePoint item: the function finds the company template by name and links it (outcome `Adopted`), after which updates are pushed.
- **Deleted** events change nothing in ConnectWise (outcome `NoChange` with a warning).
- Template names are unique on create via the API (`ObjectExists`), though existing data contains shared names.
- Dry run is ON by default. Keep it on until you have confirmed `PATCH` on a real company template with a harmless change.

Run the offline tests with `pwsh -File .\Test-Harness.ps1` (no network).

## Endpoint

`POST /api/recurring-templates/push?code=<function key>`

```json
{ "event": "created|modified|deleted", "siteUrl": "https://palaisparc.sharepoint.com/Service/servicedesk",
  "listName": "CW - Service Desk - Schedules", "itemId": 12, "dryRun": true }
```

- `siteUrl` and `listName` must match the built-in department allowlist. The item is re-read from SharePoint by `itemId`; field values in the body are ignored.
- `dryRun: true` can only force dry run; it cannot disable it.
- Set the flow's concurrency to 1.

## App settings

| Setting | Notes |
|---|---|
| `ConnectWisePsa_ApiBaseUrl`, `_ApiCompanyId`, `_ApiPublicKey`, `_ApiPrivateKey`, `_ApiClientId` | Existing settings. |
| `ConnectWisePsa_ApiVersion` | Default `2025.1` (must match the sync scripts so Fingerprints agree). |
| `Ms365_TenantId`, `Ms365_AuthAppId`, `Ms365_CertBase64`, `Ms365_CertPassword` | Existing certificate auth. App needs write on the department sites and the registry site. |
| `SpTemplateSync_DryRun` | Default `true`. Set `false` to write. |
| `SpTemplateSync_CreateMasterIfMissing` | Default `true`. Creates a master (`templateFlag=true`) when none matches the template name. |

## Behaviour

- Only non-blank SharePoint columns are pushed; a blank column leaves the CW value unchanged.
- Templates: names resolve to CW ids (board, status, type, subtype, item, team, priority, department). `CompanyId` is required. A matching master is looked up by name and created when missing. Existing templates are PATCHed with only the differing fields. Masters are never modified.
- Schedules: need `ParentTemplateId` (or a unique `ParentTemplateName`) of a synced company template, `DateStart`, and a valid `RecurType` setup (w needs DaysOfWeekRaw; m needs AbsDayNbr or SelectInterval+DaysOfWeekRaw; y needs AbsMonthNbr+AbsDayNbr). `Hours` sets the appointment length; `DateEnd` is not pushed.
- Duplicate protection: an existing same-name company template is not adopted unless a previous attempt left `PushStatus=Creating` or `AwaitingCompanyApply`; otherwise the item is Blocked.
- Echo guard: a `PushHash` of the input columns is stored. If unchanged and `PushStatus=Synced`, the call is a no-op. Write-backs use SystemUpdate.
- Write-back: `CWTemplateId`/`CWResourceId`, `Fingerprint`, `CWLastUpdated`, `CWUpdatedBy`, `RecurrenceId`, `MasterTemplateId` (if the column exists), `Push*` columns, plus a registry upsert. `SyncStatus` is never changed.
- Auto-created columns (text): `PushHash`, `PushStatus`, `PushMessage` (multi-line), `PushUtc`.
- Optional columns you may add: `AbsMonthNbr`, `SelectInterval` (schedules), `MasterTemplateId` (templates).

## Known caveats

- The existing CW -> SP sync can overwrite SharePoint edits; there is no conflict handling.
- The CW -> SP sync, audit and reconcile scripts were fixed to read `startDateTime`, `endDateTime`, `scheduleStatus.name`, `selectInterval` and `absMonthNbr`. Existing schedule items keep blank `DateStart`/`DateEnd`/`ScheduleStatus` until the sync (or reconcile) is re-run to repair them.
