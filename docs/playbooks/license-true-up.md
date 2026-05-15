# License true-up

Quarterly or pre-renewal: find waste, project savings, remediate. Pair with finance for the negotiation.

## When to use this playbook

- 60-90 days before M365 contract renewal (EA / partner / direct).
- Quarterly cost review with finance.
- After a major change in headcount (layoff, M&A, mass onboarding).
- After an MSP onboarding when you're not sure what the new tenant has.

## Outcome

A defendable savings number, broken down by category, with a remediation plan per category. Most tenants find 5-15% recoverable.

## Step 1 — Snapshot current state (1 hour)

```powershell
# License inventory
Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId,
    @{Name='Consumed'; Expression={ $_.ConsumedUnits }},
    @{Name='Available'; Expression={ $_.PrepaidUnits.Enabled }},
    @{Name='UnitsAtRisk'; Expression={ $_.PrepaidUnits.Suspended }} |
    Export-Csv -Path .\inventory-pre.csv -NoTypeInformation

# User-level assignments
Get-MgUser -All -Property UserPrincipalName, AssignedLicenses, AccountEnabled |
    ForEach-Object {
        [PSCustomObject]@{
            UPN          = $_.UserPrincipalName
            Enabled      = $_.AccountEnabled
            SkuCount     = @($_.AssignedLicenses).Count
            SkuIds       = (@($_.AssignedLicenses).SkuId -join ';')
        }
    } | Export-Csv -Path .\assignments-pre.csv -NoTypeInformation
```

Save both CSVs as "as of <date>" for the audit trail.

## Step 2 — Run the optimizer (30 minutes)

```powershell
$report = Get-LicenseOptimizationReport
$report | Format-List
```

Expected output:

```
PROJECTED MONTHLY SAVINGS
-------------------------
Anonymized stale       :   8 users  -> $ 568
License-family overlap :  23 users  -> $1,242
Disabled w/ licenses   :  17 users  -> $ 918
                                          ------
Total potential        :              $2,728  (annual $32,736)
```

Export the findings:

```powershell
$report.Findings | Export-Csv -Path .\license-findings-q1.csv -NoTypeInformation
```

Annual potential = monthly * 12.

## Step 3 — Validate findings (1-2 days)

Each category needs verification before remediation. The optimizer is precise about matching rules; it's not perfect at reading intent.

### Anonymized stale users

Pattern matches `^([A-Z0-9]{8}|[A-Z]{2,3}\d{3,5})@`. False positives:

- Service accounts with employee-ID-like UPNs (`SVC1234@`).
- Real users with short UPN aliases.
- Test accounts (often pattern-named).

Walk the list. For each:

```powershell
$upn = "ABC12345@contoso.com"
Get-MgUser -UserId $upn -Property AssignedLicenses, AccountEnabled, SignInActivity |
    Format-List UserPrincipalName, AccountEnabled, @{Name='LastSignIn'; Expression={ $_.SignInActivity.LastSignInDateTime }}
```

If account is `Enabled` + has recent sign-in, exclude from remediation. Update the CSV with a `Skip` reason.

### License-family overlap

Verify each pairing. Most legitimate. Notable false positives:

- **E3 + Visio Plan 2** — Visio is bundled differently in some E3 SKUs. Confirm via `Get-MgSubscribedSku` that your E3 includes Visio before stripping the Visio SKU.
- **E5 + Defender** — Defender add-ons sometimes come bundled, sometimes separately purchased.
- **Mix of E1 + Project Plan 1** — Project Plan 1 doesn't include everything E1 does (per-app SKU vs suite). Keep both.

For each pair, decide which SKU to drop. Usually keep the higher tier; sometimes intent is to drop the higher tier in favor of the cheaper one (e.g. user moved to a less-intensive role).

### Disabled users still consuming licenses

Usually safe to remediate. But verify nothing's on litigation hold or in a 90-day retention bucket:

```powershell
# Litigation hold check (EXO):
Get-Mailbox -Identity "alice@contoso.com" | Select-Object DisplayName, LitigationHoldEnabled, RetentionHoldEnabled
```

If `LitigationHoldEnabled=True`, escalate to legal before delicensing. The mailbox needs an active license to honor the hold.

## Step 4 — Remediation plan (1 day)

Build a CSV of approved removals:

```csv
UPN,SkuId,Reason,ApprovedBy
ABC12345@contoso.com,05e9a617-0261-4cee-bb44-138d3ef5d965,Stale; never signed in,manager@contoso.com
alice@contoso.com,87f2-...,Family overlap E3+BizPrem; keeping BizPrem,manager@contoso.com
disabled-user@contoso.com,05e9a617-0261-4cee-bb44-138d3ef5d965,Disabled since 2025-09-15,it-mgr@contoso.com
```

Submit for approval (manager / finance / IT director per your process).

## Step 5 — Run in PREVIEW (15 minutes)

```powershell
Invoke-BulkLicenseRemoval -Path .\license-cleanup-q1-approved.csv -WhatIf
```

Review the result CSV. Every row should show `Status=Preview`.

## Step 6 — Run LIVE (30-60 minutes)

```powershell
Invoke-BulkLicenseRemoval -Path .\license-cleanup-q1-approved.csv
```

Result CSV next to input. Each row:
- `Status=Success` — license removed, audit entry written, reverse-recipe captured.
- `Status=PartialSuccess` — one of multiple SKU removals on the row failed; rest succeeded.
- `Status=Failed` — full failure; `Reason` shows why.

## Step 7 — Verify (1 day)

After 24 hours, re-snapshot:

```powershell
Get-MgSubscribedSku | Select-Object SkuPartNumber, ConsumedUnits |
    Export-Csv -Path .\inventory-post.csv -NoTypeInformation
```

Diff:

```powershell
$pre = Import-Csv .\inventory-pre.csv
$post = Import-Csv .\inventory-post.csv

Compare-Object $pre $post -Property SkuPartNumber, ConsumedUnits -PassThru | Format-Table
```

Confirm the savings number matches the plan. If not, walk the audit log:

```powershell
Read-AuditEntries | Where-Object {
    $_.actionType -eq 'RemoveLicense' -and $_.ts -ge (Get-Date).AddDays(-1)
} | Format-Table ts, target, result
```

## Step 8 — Renewal negotiation (variable)

Hand finance:

- Pre-remediation inventory (`inventory-pre.csv`).
- Optimizer report (`license-findings-q1.csv`).
- Approved removals (`license-cleanup-q1-approved.csv`).
- Post-remediation inventory (`inventory-post.csv`).
- Calculated annual savings = current monthly savings × 12.

Negotiating points:
- "We're paying for 2,150 E3 seats but actually using 1,997."
- "Family overlap was 23 users; corrected. Renew at 2,127 E3 + 23 Business Premium instead of 2,150 E3 + 23 Business Premium."
- "Visio Plan 2 add-ons: 47 users using, 12 stale. Renew 47."

## Step 9 — Schedule the next round (5 minutes)

Make this routine — schedule the optimizer to run monthly as a health check:

```powershell
Register-ScheduledHealthCheck -CheckName license-usage -Trigger "Daily 03:00" -OnFinding NotifyOperations
```

The check writes findings to `<stateDir>\health-results\health-license-usage-<ts>.json` + emails the operations team if findings exceed a configured threshold.

## Common pitfalls

- **Tenant billing cycle vs M365 license assignment** — removing a license today doesn't refund this month. Savings start next billing cycle.
- **Mailbox license requirements** — a Shared mailbox doesn't need a license. A mailbox on litigation hold needs the right license type. Verify before delicensing.
- **Reactivation cost** — if you delicense and the user comes back (rehire, return from leave), reactivating may take 24-48h depending on tenant policy. Document this for HR.
- **Per-user SKU bundles** — some tenants bought "E3 + Compliance + Identity Protection" as separate SKUs assigned together. Treat the bundle as one unit; remove the matching SKU set together.

## See also

- [`../guides/license-optimization.md`](../guides/license-optimization.md) — operator-facing guide.
- [`../reference/template-schema.md`](../reference/template-schema.md) — `license-prices.json` schema.
- [`../guides/scheduled-checks.md`](../guides/scheduled-checks.md) — automating the optimizer.
- [`audit-prep.md`](audit-prep.md) — embedding license efficiency into the audit evidence.
