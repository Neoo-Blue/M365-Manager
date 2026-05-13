# Per-tenant configuration overrides (Phase 6)

Some configuration keys make more sense per-tenant than globally.
Phase 6 adds a thin resolution layer that lets one key take
different values in different tenants without forking the whole
config.

## Resolution order

`Get-EffectiveConfig -Key <name> [-GlobalConfig <ht>] [-Tenant <name>] [-CliValue <v>]`
walks four sources, **last wins**:

1. The caller's `-GlobalConfig` hashtable (the existing `ai_config.json`,
   `Notifications` block, etc.).
2. `<stateDir>/tenant-overrides/<tenant-name>.json` -- partial JSON
   file that overrides any subset of keys.
3. `$env:M365MGR_<KEY>` -- env var with key uppercased and `.`
   replaced by `_` (e.g. `M365MGR_AI_MONTHLYBUDGETUSD`).
4. `-CliValue` -- caller-supplied explicit flag value.

Empty strings and `$null` are treated as "not present" so an empty
CLI flag can't blow away a legitimate override.

## Overridable keys

`Get-TenantOverridableKeys` returns the canonical list:

| Key                          | Default origin               | Why per-tenant?                                            |
|------------------------------|------------------------------|------------------------------------------------------------|
| `StaleGuestDays`             | hard-coded 90                | Some customers run shorter (regulated) or longer SLAs.    |
| `OneDriveRetentionDays`      | Phase 3 default              | Per-customer compliance windows.                          |
| `OneDriveRetentionPolicy`    | Phase 3                      | Some customers archive, some delete.                      |
| `Notifications.Recipients`   | Notifications block          | Different IT contact per customer.                        |
| `Notifications.TeamsWebhook` | Notifications block          | Customer-specific Teams channel.                          |
| `Notifications.SmtpFrom`     | Notifications block          | Customer-branded sender address.                          |
| `AI.MonthlyBudgetUsd`        | AI config                    | Spend per customer, not per operator.                     |
| `AI.AlertAtPct`              | AI config (default 80)       | Per-customer risk tolerance.                              |
| `AI.AutoPlanThreshold`       | Phase 5 (default 3)          | Some customers want plans even for 2-step changes.        |
| `LicensePrices`              | Phase 4 templates/license-prices.json | Negotiated per-customer pricing.                  |
| `DefaultRoleTemplate`        | Phase 1 templates            | Per-customer onboarding flavors.                          |
| `BreakGlassReminderDays`     | Phase 4 (default 90)         | Per-customer rotation cadence.                            |
| `AuditRetentionDays`         | global                       | Per-customer audit-retention policy.                      |

Keys NOT in this list are global-only (provider type, API key,
SDK module versions, etc.). The list is intentionally small --
adding to it requires updating `$script:TenantOverridableKeys` in
`TenantOverrides.ps1` and making the consuming module read through
`Get-EffectiveConfig`.

## Override file schema

`<stateDir>/tenant-overrides/<tenant-name>.json` is a flat
hashtable. Only the keys you want to override need to be present;
omitted keys fall back to the global value.

```jsonc
{
  "StaleGuestDays":              30,
  "Notifications.Recipients":    ["bob@contoso.com","alice@contoso.com"],
  "Notifications.TeamsWebhook":  "https://outlook.office.com/webhook/...",
  "AI.MonthlyBudgetUsd":         25.00,
  "AI.AlertAtPct":               75,
  "LicensePrices": {
    "ENTERPRISEPACK": { "monthlyUsd": 35.00 }
  }
}
```

The tenant-name slug is the tenant's `name` lowercased, non-alnum
characters replaced with `-`. `Show-TenantOverrides -Name <n>` and
`Edit-TenantOverrides -Name <n>` are the friendly entry points.

## Adoption pattern for module authors

Old (global-only):
```powershell
$days = if ($Config.StaleGuestDays) { [int]$Config.StaleGuestDays } else { 90 }
```

New (tenant-overridable):
```powershell
$days = if (Get-Command Get-EffectiveConfig -ErrorAction SilentlyContinue) {
    $eff = Get-EffectiveConfig -Key 'StaleGuestDays' -GlobalConfig $Config
    if ($eff) { [int]$eff } else { 90 }
} else {
    if ($Config.StaleGuestDays) { [int]$Config.StaleGuestDays } else { 90 }
}
```

The `Get-Command` guard means modules still work when
TenantOverrides.ps1 isn't loaded (e.g. a smoke test that dot-sources
only the consumer module).

## What's wired today

Phase 6 commit E wires:
- `AICostTracker.ps1`: `AI.MonthlyBudgetUsd`, `AI.AlertAtPct`.
- `GuestUsers.ps1`: `StaleGuestDays` default in `Get-StaleGuests`.

The remaining keys are listed in `$script:TenantOverridableKeys`
but their consuming modules still read directly from `$Config`.
Converting them mechanically is a follow-up task -- the resolution
layer is the load-bearing piece and that's done.
