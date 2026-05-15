# Template schemas

JSON schemas for the three template types shipped under `templates/`. Drop new files matching these schemas into the respective dirs and they're auto-discovered.

## Role templates

`templates/role-<slug>.json` — used by `Onboard.ps1` and `BulkOnboard.ps1`.

```jsonc
{
  "name":                "sales-rep",
  "description":         "Standard sales rep — E3 + Sales channels + CRM shared mailbox.",
  "licenses":            ["SPE_E3"],
  "groups":              ["Sales-NorthAmerica", "All-Employees"],
  "distributionLists":   ["sales-announce@contoso.com"],
  "sharedMailboxes": [
    { "mailbox": "crm-shared@contoso.com", "permission": "FullAccess" }
  ],
  "calendars": [
    { "calendar": "team-calendar@contoso.com", "permission": "Editor" }
  ],
  "contractorExpiryDays": null
}
```

| Field | Required | Notes |
|---|---|---|
| `name` | yes | Lowercase, hyphen-separated. Matches the filename minus `role-` and `.json`. |
| `description` | yes | One-sentence summary. Shown in the picker. |
| `licenses` | no | Array of SKU part numbers (e.g. `SPE_E3`, `EMS_E3`). |
| `groups` | no | Array of security group display names. |
| `distributionLists` | no | Array of DL primary SMTP addresses. |
| `sharedMailboxes` | no | Array of `{ mailbox, permission }` (permission ∈ `FullAccess` / `SendAs` / `SendOnBehalf`). |
| `calendars` | no | Array of `{ calendar, permission }` (permission ∈ `Reviewer` / `Author` / `Editor` / `PublishingEditor`). |
| `contractorExpiryDays` | no | If non-null, records `employeeLeaveDateTime` at `(Get-Date).AddDays(N)`. Surfaces in stale-user reports. |

Unknown SKUs / groups / DLs / shared mailboxes log a per-item warning + skip; the onboard continues.

## Site templates

`templates/site-<slug>.json` — used by `SharePoint.ps1` site provisioning.

```jsonc
{
  "name":          "project",
  "description":   "Project collaboration site with external sharing disabled.",
  "siteUrl":       "/sites/proj-{slug}",
  "title":         "Project: {project}",
  "owner":         "{owner}",
  "template":      "STS#3",
  "sharing":       "ExistingExternalUserSharingOnly",
  "storageQuotaMb": 1024,
  "groupConnected": false
}
```

| Field | Required | Notes |
|---|---|---|
| `name` | yes | Picker key. |
| `description` | yes | |
| `siteUrl` | yes | URL template. Placeholders (`{slug}`, `{project}`, `{owner}`) get filled at create time. |
| `title` | yes | Display title template. Same placeholder rules. |
| `owner` | yes | UPN or `{owner}` placeholder. |
| `template` | yes | SPO web template id. `STS#3` = modern team site; `SITEPAGEPUBLISHING#0` = communication site. |
| `sharing` | no | `Disabled` / `ExistingExternalUserSharingOnly` / `ExternalUserSharingOnly` / `ExternalUserAndGuestSharing`. |
| `storageQuotaMb` | no | Quota in megabytes. |
| `groupConnected` | no | `true` to provision an M365 group + Teams team alongside. |

## Scheduled-check templates

Scheduled-check scripts live under `health-checks/*.ps1` — they're not template files but they share a contract:

| Convention | Notes |
|---|---|
| Filename | `health-<area>.ps1`. |
| Param block | `param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings')`. |
| Bootstrap | `$root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive -Modules ...`. |
| Result emit | `& "$PSScriptRoot/_writeresult.ps1" -CheckName 'area' -Status 'clean'\|'findings'\|'failure' -FindingCount N -Findings @{}`. |

Result file shape (written by `_writeresult.ps1` to `<stateDir>/health-results/health-<area>-<ts>.json`):

```jsonc
{
  "checkName":   "mfa-gaps",
  "ranUtc":      "2026-05-14T17:00:00Z",
  "status":      "findings",
  "findingCount": 23,
  "findings":   { ... arbitrary check-specific shape ... },
  "host":        "OPS-LAPTOP-01",
  "tenant":      "contoso.onmicrosoft.com"
}
```

The MSP dashboard (`MSPDashboard.ps1`) reads these to build the per-tenant card rollups.

## License prices

`templates/license-prices.json` — used by `LicenseOptimizer.ps1` for cost math.

```jsonc
{
  "SPE_E3":         { "monthly": 54.00, "annual": 648.00 },
  "SPE_E5":         { "monthly": 86.00, "annual": 1032.00 },
  "SPB":            { "monthly": 22.00, "annual": 264.00 },
  "ENTERPRISEPACK": { "monthly": 23.00, "annual": 276.00 },
  "POWER_BI_PRO":   { "monthly": 9.99,  "annual": 119.88 }
}
```

Keys are SKU part numbers (from `Get-MgSubscribedSku`). Edit to match your contract. The shipped values are list price; EA / partner discounts go in the override file via per-tenant config:

```jsonc
// tenant-overrides/contoso.json
{
  "LicensePriceOverrides": {
    "SPE_E3": { "monthly": 38.00, "annual": 456.00 }
  }
}
```

## AI prices

`templates/ai-prices.json` — used by `AICostTracker.ps1`.

```jsonc
{
  "Anthropic": {
    "claude-sonnet-4-20250514": { "input": 3.00, "output": 15.00 },
    "claude-opus-4-1-20250805":  { "input": 15.00, "output": 75.00 }
  },
  "OpenAI": {
    "gpt-4o": { "input": 2.50, "output": 10.00 }
  }
}
```

Prices are per million tokens. Update when negotiating volume discounts.

## Tabletop scenarios

See [`../guides/tabletop-exercises.md`](../guides/tabletop-exercises.md) — `templates/tabletop-scenarios/<name>.json` schema includes `name`, `description`, `severity`, `trigger`, `expectedActions`, `expectedFindings`, `recoveryNotes`.

## See also

- [`csv-formats.md`](csv-formats.md) — bulk operation CSV schemas.
- [`config-keys.md`](config-keys.md) — every config key.
- [`../getting-started/configuration.md`](../getting-started/configuration.md) — config overview.
