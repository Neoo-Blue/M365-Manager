# License optimization

Find license waste, project savings in USD, and remediate (in PREVIEW first, of course). Covers the License Optimizer report at menu slot 20.

## When to use this

- Quarterly true-up before the next M365 renewal.
- After a layoff or mass offboarding when stale licenses are likely.
- Audit prep — finance wants to know what you're paying for vs using.
- Onboarding a new tenant under MSP management.

## Prereqs

- `License Administrator` + `User Administrator` roles on the tenant.
- `templates/license-prices.json` has accurate per-SKU pricing for your contract. The shipped values are list price — operators with EA / partner discounts should override.

## What the optimizer reports

**Slot 20 → License & Cost → License optimizer**

Three categories of finding:

### 1. Anonymized usernames

Pattern: `^([A-Z0-9]{8}|[A-Z]{2,3}\d{3,5})@` — e.g. `ABC12345@contoso.com` or `XY7723@contoso.com`. These are usually stale: ex-employees whose accounts were renamed (a common compliance practice) but never delicensed.

Output:

```
  ANONYMIZED USERNAMES (likely stale)
  ----------------------------------
  UPN                            Last sign-in    Licenses     Monthly USD
  ABC12345@contoso.com           never          SPE_E3       $54
  XY7723@contoso.com             403d ago       SPE_E5       $86
  ...
```

### 2. License-family overlap

Same user, two SKUs that include each other's features. Examples:

- **E3 + Business Premium** — Business Premium includes everything E3 does.
- **E1 + E3** — E3 is a superset of E1.
- **Visio Plan 2 + E5** — E5 includes Visio.

The optimizer reads the SKU-to-service-plan map and surfaces every user with redundant SKU coverage.

### 3. Disabled users still consuming licenses

`accountEnabled=false` + licenses assigned. Common after offboarding when the operator blocked sign-in but didn't run the full 12-step flow.

## Cost math

The optimizer reads `templates/license-prices.json`:

```jsonc
{
  "SPE_E3":           { "monthly": 54.00, "annual": 648.00 },
  "SPE_E5":           { "monthly": 86.00, "annual": 1032.00 },
  "SPB":              { "monthly": 22.00, "annual": 264.00 },
  "ENTERPRISEPACK":   { "monthly": 23.00, "annual": 276.00 },
  ...
}
```

Update this file to reflect your contract. Per-user savings:

- Anonymized stale: full monthly cost (assumes you'd unassign).
- Family overlap: lower-cost SKU's monthly (assumes you'd keep the higher-tier).
- Disabled with licenses: full monthly cost.

Top of the report prints the rollup:

```
  PROJECTED MONTHLY SAVINGS
  -------------------------
  Anonymized stale       :   8 users  -> $ 568
  License-family overlap :  23 users  -> $1,242
  Disabled w/ licenses   :  17 users  -> $ 918
                                          ------
  Total potential        :              $2,728
```

## Remediation

The optimizer reports findings — it does NOT auto-remediate. You drive remediation from the same menu:

**Slot 20 → License & Cost → Remediate selected findings**

You pick which categories to act on (a multi-select picker). The tool walks each user:

1. PREVIEW (mode banner is yellow): writes per-user audit entries with `actionType=RemoveLicense`, `event=PREVIEW`.
2. Confirm "Apply N license removals across M users?"
3. LIVE: runs `Set-MgUserLicense` per user, audited, reversible.

Every removal gets a `reverse` recipe (`ReverseType=AssignLicense`) so `Undo-Incident` / `Invoke-Undo` can restore.

## Common gotchas

- **Disabled users with onboarding holds.** Some orgs delicense at 90 days, not at offboard. The optimizer doesn't know your policy — review the "disabled w/ licenses" list before bulk-remediating.
- **Pattern-matched stale UPNs that are real.** Some tenants use 6-char employee IDs as UPNs (`AB1234@`). The anonymized-username regex flags these. Skip them at remediation review.
- **SKU pricing accuracy.** If `license-prices.json` is out of date, the savings number is wrong. The tool surfaces the file path in the report header so the operator can verify.

## Bulk remediation from CSV

For very large license cleanups, export the findings to CSV and process with the bulk path:

```powershell
$savings = Get-LicenseOptimizationReport -Tenant Contoso
$savings.Findings | Export-Csv -Path .\license-cleanup.csv -NoTypeInformation
# Review the CSV
# Then:
Invoke-BulkLicenseRemoval -Path .\license-cleanup.csv -WhatIf
```

CSV columns: `UPN, SkuId, Reason, ApprovedBy`. Sample at `templates/license-remediation-sample.csv`.

## Scheduled tracking

`health-checks/health-license-usage.ps1` runs the optimizer non-interactively and writes the rollup to `<stateDir>/health-results/`. Wire it as a scheduled health check (option 21) for periodic visibility. See [`scheduled-checks.md`](scheduled-checks.md).

## See also

- [`../reference/csv-formats.md`](../reference/csv-formats.md) — license-cleanup CSV schema.
- [`../reference/template-schema.md`](../reference/template-schema.md) — `license-prices.json` schema.
- [`offboarding.md`](offboarding.md) — the offboard flow that should have delicensed in the first place.
- [`scheduled-checks.md`](scheduled-checks.md) — automating the optimizer.
