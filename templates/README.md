# Onboarding role templates

A template is a JSON file under `templates/` named `role-<slug>.json`. The slug is the lowercase short name (`sales-rep`, `engineer`, `exec-assistant`, `contractor`, `default`). The `Templates.ps1` module discovers everything matching `role-*.json` automatically — to add a new role, drop a new file in this directory and it shows up in the onboarding picker on next launch.

The templates committed here are **examples**. The SKU part numbers, group names, DL names, and shared mailbox addresses all need to match what actually exists in your tenant. Run `Get-MgSubscribedSku`, `Get-MgGroup`, and `Get-DistributionGroup` to find the right values.

## Schema

```json
{
  "name":          "Display name shown in the picker",
  "description":   "One-line summary",
  "usageLocation": "ISO 3166-1 alpha-2 country code (US, GB, DE, ...)",

  "licenseSKUs":       ["SKU_PART_NUMBER_1", "SKU_PART_NUMBER_2"],
  "securityGroups":    ["SG-Name-1", "SG-Name-2"],
  "distributionLists": ["DL-Name-1"],
  "sharedMailboxes":   [
    { "identity": "mbox@contoso.com", "access": "Full" },
    { "identity": "mbox@contoso.com", "access": "SendAs" },
    { "identity": "mbox@contoso.com", "access": "FullSendAs" }
  ],
  "teams":             ["Team Display Name"],

  "oneDrive":             { "quotaGB": 1024, "provisionImmediately": true },
  "defaults":             { "Department": "...", "OfficeLocation": "..." },
  "contractorExpiryDays": null
}
```

| Field | Required | Notes |
|---|---|---|
| `name` | yes | Shown in the picker. |
| `description` | yes | Shown under the name. |
| `usageLocation` | yes | Required for license assignment. ISO 2-letter code. |
| `licenseSKUs` | no | Array of `SkuPartNumber` strings (e.g. `SPE_E3`, `POWER_BI_PRO`). Unknown SKUs are **skipped with a warning** and flagged in the bulk result CSV; the rest of the onboard continues. |
| `securityGroups` | no | Display names. Unknown groups are skipped with a warning. |
| `distributionLists` | no | Display names. Unknown DLs are skipped with a warning. |
| `sharedMailboxes` | no | Array of `{ identity, access }`. `access` is one of `Full`, `SendAs`, `FullSendAs`. |
| `teams` | no | Display names. Teams assignment lands in Phase 2; currently logged but not executed. |
| `oneDrive` | no | `quotaGB` (int) and `provisionImmediately` (bool). Provisioning hook lands in Phase 3. |
| `defaults` | no | Fields written to the user object if the operator left them blank (e.g. `Department`, `OfficeLocation`). Operator values always win. |
| `contractorExpiryDays` | no | If set to a positive integer, the user's `accountExpires` is set to `now + N days` after creation. Use for contractors / temps. |

`_comment` fields (or any key starting with `_comment`) are stripped on load — use them freely to annotate.

## Sample CSVs

`bulk-onboard-sample.csv` and `bulk-offboard-sample.csv` demonstrate the column shape expected by `Invoke-BulkOnboard` / `Invoke-BulkOffboard`. Copy these and replace with real data.

## Unknown-SKU policy (Phase 1 default)

When a template references a SKU that the tenant doesn't own, the onboarding **continues** with a warning printed and a `SkipReason` recorded in the bulk-onboard result CSV. Rationale: tenants often have a mix of bundled / overlapping licenses, and failing the whole onboard for a missing add-on (Power BI Pro, Phone System) would block the common case. Override by editing the template to drop the SKU.
