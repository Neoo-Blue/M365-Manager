# Audit prep

Annual / quarterly compliance audit. The auditor wants evidence of: documented IR process, role-based access, license efficiency, MFA coverage, retention practices, audit-log integrity. This playbook walks through producing that evidence with the tool.

## When to use this playbook

- SOC 2 / ISO 27001 / HIPAA / FedRAMP audit cycle.
- Pre-due-diligence for an M&A.
- Annual access review.
- Internal compliance review.

## Lead time

Most audit prep takes 2-5 business days end-to-end. Start at least 2 weeks before the auditor's scheduled review.

## What the tool can demonstrate

| Evidence | Source |
|---|---|
| Documented IR process | [`compromised-account.md`](compromised-account.md) + [`../guides/incident-response.md`](../guides/incident-response.md) |
| Tabletop exercise / IR readiness | `Invoke-IncidentTabletop` reports |
| Role-based access | Role templates + onboarding audit log |
| License efficiency | License optimizer report |
| MFA coverage | MFA compliance views + CSV export |
| Audit-log integrity | Per-session JSONL files + tenant-stamped naming |
| Retention practices | 365-day default incident retention + UAL retention notes |
| Reversibility / undo | Undo handler dispatch table |
| PII redaction in AI flows | `/privacy` settings + redaction tests |
| Per-tenant isolation | Multi-tenant audit log + per-tenant incident dirs |

## Preparation checklist

### 1. Run a tabletop exercise (1 day)

Demonstrates documented IR readiness. Run all four shipping scenarios:

```powershell
foreach ($scenario in 'phishing-campaign','insider-mass-download','mfa-bypass','compromised-vendor') {
    Invoke-IncidentTabletop -ScenarioName $scenario
}
```

Output: four `tabletop-report.html` files. Bundle them:

```powershell
$bundle = Join-Path (Get-Location) "tabletop-reports-q1-2026.zip"
Get-Incidents -Days 1 | Where-Object { $_.reason -like 'Tabletop exercise*' } | ForEach-Object {
    Export-Incident -Id $_.id -Path (Join-Path (Get-Location) "tabletop-$($_.id).zip")
}
```

Hand the bundle to the auditor as "IR readiness evidence Q1 2026."

### 2. License + MFA + stale-account compliance reports (1 day)

```powershell
# License optimizer rollup
$savings = Get-LicenseOptimizationReport
$savings | Export-Csv -Path .\license-optimization-q1.csv -NoTypeInformation

# MFA gaps
Get-UsersWithNoMfa | Export-Csv -Path .\users-no-mfa-q1.csv -NoTypeInformation
Get-UsersWithOnlyPhoneMfa | Export-Csv -Path .\users-phone-only-mfa-q1.csv -NoTypeInformation
Get-Fido2Users | Export-Csv -Path .\users-fido2-q1.csv -NoTypeInformation

# Stale guests
Get-StaleGuests -DaysSinceSignIn 90 | Export-Csv -Path .\stale-guests-q1.csv -NoTypeInformation

# Break-glass posture
Test-BreakGlassPosture | ConvertTo-Json -Depth 6 | Set-Content .\breakglass-posture-q1.json
```

Each CSV is "as of this snapshot date" — annotate the run timestamp.

### 3. Audit log integrity demonstration (1 hour)

Auditors want to see that operations are logged + reversible + tenant-scoped.

```powershell
# Per-tenant log files exist:
Get-ChildItem "$env:LOCALAPPDATA\M365Manager\audit\" -Filter session-*.log | Format-Table Name, Length, LastWriteTime

# Sample a recent log file -- show structured fields:
Read-AuditEntries | Select-Object -First 10 | Format-Table ts, event, actionType, result, @{Name='Tenant'; Expression={ $_.tenant.name }}

# Demonstrate undo:
$lastReversible = (Get-UndoableEntries | Select-Object -First 1)
Write-Host "Reversible entry: $($lastReversible.entryId)"
Write-Host "Reverse recipe: $($lastReversible.reverse.type) -- $($lastReversible.reverse.description)"
```

Export a sanitized day's audit log:

```powershell
Read-AuditEntries | Where-Object { $_.ts -ge (Get-Date).AddDays(-1) } |
    Export-AuditEntriesHtml -Path .\audit-sample-day.html
```

The HTML report is self-contained — safe to hand the auditor as "here's what a day of operations looks like."

### 4. Access review (1-2 days)

Walk every employee + verify their group / license assignments match their role:

```powershell
# Generate the manifest of everyone's assignments:
Get-MgUser -All -Property UserPrincipalName, Department, JobTitle, AssignedLicenses |
    ForEach-Object {
        $upn = $_.UserPrincipalName
        $groups = (Get-MgUserMemberOf -UserId $upn).Value | ForEach-Object { $_.DisplayName }
        [PSCustomObject]@{
            UPN         = $upn
            Department  = $_.Department
            JobTitle    = $_.JobTitle
            SKUs        = ($_.AssignedLicenses.SkuId -join ';')
            GroupCount  = $groups.Count
            Groups      = ($groups -join ';')
        }
    } | Export-Csv -Path .\access-review-q1.csv -NoTypeInformation
```

Review with department heads — anyone with anomalous access flagged.

### 5. Configuration documentation (30 minutes)

Auditors often want to see how the tool is configured:

```powershell
# Sanitized config dump (DPAPI'd fields show as DPAPI: marker):
Get-Content $env:LOCALAPPDATA\M365Manager\ai_config.json |
    ConvertFrom-Json |
    ConvertTo-Json -Depth 8 |
    Set-Content .\tool-config-q1.json

# Tenant registry:
Get-Tenants | Select-Object Name, TenantId, Domain, AuthMode |
    Export-Csv -Path .\tenant-registry-q1.csv -NoTypeInformation
```

DPAPI-encrypted fields are visible as their `DPAPI:` prefix marker — the auditor can verify "yes, secrets are at-rest encrypted."

### 6. Incident retrospective (1-2 days)

If there were any real incidents this period:

```powershell
$periodStart = (Get-Date "2026-01-01")
$incidents = Get-Incidents -Status All | Where-Object { ([DateTime]$_.startedUtc) -ge $periodStart }
$incidents | ForEach-Object {
    Export-Incident -Id $_.id -Path ".\audit-bundle\incident-$($_.id).zip"
}
```

Bundle every Q1 incident. Each bundle includes the snapshot + audits + report + filtered audit log slice — auditor-grade evidence of the IR process.

## What to hand the auditor

The complete bundle:

```
audit-bundle-q1-2026/
├── tabletop-reports/
│   ├── tabletop-phishing-campaign.html
│   ├── tabletop-insider-mass-download.html
│   ├── tabletop-mfa-bypass.html
│   └── tabletop-compromised-vendor.html
├── compliance-reports/
│   ├── license-optimization-q1.csv
│   ├── users-no-mfa-q1.csv
│   ├── users-fido2-q1.csv
│   ├── stale-guests-q1.csv
│   └── breakglass-posture-q1.json
├── audit-log-samples/
│   └── audit-sample-day.html
├── access-review/
│   └── access-review-q1.csv
├── tool-config/
│   ├── tool-config-q1.json
│   └── tenant-registry-q1.csv
├── incidents-q1/
│   ├── incident-INC-2026-01-12-xxxx.zip
│   ├── incident-INC-2026-02-03-yyyy.zip
│   └── ...
└── docs/
    └── (copy of relevant /docs/ tree -- see-also below)
```

Plus copies of the docs that document YOUR process:

| Doc | Purpose for the audit |
|---|---|
| [`compromised-account.md`](compromised-account.md) | "This is our documented IR process." |
| [`incident-triggers.md`](incident-triggers.md) | "This is our auto-detection framework + trust model." |
| [`../guides/breakglass-accounts.md`](../guides/breakglass-accounts.md) | "This is our break-glass posture." |
| [`../concepts/security-model.md`](../concepts/security-model.md) | "This is what we encrypt + redact + audit." |
| [`../operations/permissions.md`](../operations/permissions.md) | "This is who has what." |

## Post-audit follow-up

After the audit, review every finding:

1. **Findings against documented process** — update docs to match what actually happened, OR fix the process to match documented intent.
2. **Findings against config** — adjust thresholds, recipients, retention. Land changes in `ai_config.json` + `tenant-overrides/` and document in git history.
3. **Findings against tool behavior** — file issues against the M365 Manager repo if the tool itself needs to change.

## See also

- [`compromised-account.md`](compromised-account.md) — the playbook auditors care most about.
- [`../guides/tabletop-exercises.md`](../guides/tabletop-exercises.md) — IR-readiness drills.
- [`../guides/audit-and-undo.md`](../guides/audit-and-undo.md) — audit log handling.
- [`../operations/permissions.md`](../operations/permissions.md) — role matrix.
