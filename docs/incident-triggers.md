# Incident detection triggers

Auto-detection framework that surfaces findings into the incident response workflow. Findings open a Low-severity (forensic-only) incident automatically and alert the security team. **No detector auto-runs the destructive playbook unless `AutoExecuteOnSeverity` is explicitly set** — the default is `None` (human always confirms).

## The seven detectors

| Detector | Severity | What it flags |
|---|---|---|
| `Detect-AnomalousLocationSignIn` | Low | Sign-in from a country never seen for this user in the last `AnomalousLocationLookbackDays` days. |
| `Detect-ImpossibleTravel` | High | Two sign-ins separated by more physical distance than `ImpossibleTravelMaxKmPerHour` permits. |
| `Detect-HighRiskSignIn` | High | Identity Protection flagged the sign-in as `risk=high`. |
| `Detect-MassFileDownload` | High | `>MassDownloadFileCount` FileDownloaded events within `MassDownloadWindowMinutes`. |
| `Detect-MassExternalShare` | High | `>MassShareCount` outbound shares to external recipients within `MassShareWindowMinutes`. |
| `Detect-SuspiciousInboxRule` | Critical | Inbox rule matching the classic AiTM-kit signature: forward-to-external + delete-from-sent + degenerate name. |
| `Detect-MFAFatigue` | High | `>MFAFatigueRejectCount` rejected MFA prompts within `MFAFatigueWindowMinutes`. |

## Default thresholds

All configurable in `ai_config.json` under the `IncidentResponse` block. Defaults shipped:

```jsonc
{
  "IncidentResponse": {
    "AutoExecuteOnSeverity":         "None",
    "AnomalousLocationLookbackDays": 90,
    "ImpossibleTravelMaxKmPerHour":  900,
    "MassDownloadFileCount":         50,
    "MassDownloadWindowMinutes":     5,
    "MassShareCount":                20,
    "MassShareWindowMinutes":        60,
    "MFAFatigueRejectCount":         10,
    "MFAFatigueWindowMinutes":       60,
    "DetectorIntervalMinutes":       15
  }
}
```

Tune per tenant via `tenant-overrides/<name>.json` (Phase 6's `Get-EffectiveConfig` layer resolves the right value).

## AutoExecuteOnSeverity — the consequential setting

This is the most consequential config key in the entire tool. It controls whether a detector finding will **auto-run the destructive playbook** without operator confirmation.

| Value | Behavior |
|---|---|
| `None` *(default)* | Findings open a Low incident + alert. Operator decides whether to escalate. **Conservative; matches typical M365 admin trust model.** |
| `Critical` | Findings with severity Critical (only `Detect-SuspiciousInboxRule` today) auto-run the playbook at Critical severity. Findings of lower severity still require operator confirmation. **Mature SOC posture; assumes the detector criteria are precise enough that false-positives are rare.** |
| `HighAndCritical` | All High and Critical findings auto-run. **Aggressive; suitable only for shops with a defined runbook for handling auto-containment false-positives.** |

`None` ships as the default. Override at the tenant level only after a documented review.

## How the detector framework works

`Invoke-IncidentDetectors -UPNs <list> | -All` iterates every detector against every user. Each detector is a function that takes a UPN, queries Graph / UAL / SharePoint as needed, and returns either `$null` (no finding) or a `Finding` hashtable:

```powershell
@{
  TriggerType        = 'SuspiciousInboxRule'
  UPN                = 'alice@contoso.com'
  Severity           = 'Critical'
  Evidence           = @{ ruleId = '...'; ruleName = '.'; forwardsExternal = $true; movesOutOfSight = $true }
  RecommendedAction  = "Classic AiTM rule signature. Run /incident alice@contoso.com Critical -QuarantineSentMail..."
  DetectedUtc        = '2026-05-14T17:42:11.000Z'
}
```

On each finding the framework:

1. **Auto-opens a Low-severity forensic incident** via `Invoke-CompromisedAccountResponse`. This captures the snapshot, three audits, and report immediately — so even if you ignore the alert, you have the forensic baseline on disk.
2. **Sends a notification** to the security team via `Send-Notification`.
3. **Checks `AutoExecuteOnSeverity`**. If the finding qualifies for auto-escalation, runs the destructive playbook at the finding's severity, non-interactive. Otherwise stops at step 1 and waits for the operator.

## Scheduled detection

Ship `health-checks/health-incident-triggers.ps1` runs the full detector sweep. Wire it into the scheduler (option 21 → Scheduled Health Checks → register):

- **Interval**: `IncidentResponse.DetectorIntervalMinutes` (default 15).
- **Scope**: scan all enabled users by default. On large tenants, pass `-UPNs alice@x,bob@x` to scope the sweep to a watchlist.

Results land via `_writeresult.ps1` in the same JSONL format every other health check uses, so AuditViewer / `Read-AuditEntries` / the MSP dashboard can all consume them.

## Adding a custom detector

1. Write a function `Detect-<YourName>` that:
   - Takes `[Parameter(Mandatory)][string]$UPN`.
   - Returns `$null` for no finding, or a finding hashtable via `New-IncidentFinding`.
   - Reads only — does not mutate.
2. Add it to the `$detectors` array in `Invoke-IncidentDetectors` (`IncidentTriggers.ps1`).
3. Add a Pester test in `tests/IncidentTriggers.Tests.ps1` exercising the predicate against canned data.

Example skeleton:

```powershell
function Detect-MyCustomTrigger {
    param([Parameter(Mandatory)][string]$UPN)
    $cfg = Get-IncidentTriggerConfig
    # ... your logic, querying Graph / UAL / etc.
    if ($somethingBad) {
        return New-IncidentFinding `
            -TriggerType 'MyCustomTrigger' `
            -UPN $UPN `
            -Severity 'High' `
            -Evidence @{ key1 = $value1; key2 = $value2 } `
            -RecommendedAction "What the operator should do."
    }
    return $null
}
```

## Sanity-check thresholds

Defaults are tuned for a mid-size tenant where:

- Users normally sign in from 1-3 countries; >2 hop in 1 day is suspicious.
- 50 downloads in 5 minutes is unusual for knowledge work; a legitimate bulk SharePoint sync would be slower.
- 20 external shares in 1 hour is way past the social-collaboration baseline.
- 10 rejected MFA prompts in an hour is unambiguous fatigue attack.

If your tenant has unusual workloads (M&A research, data-science teams pulling large datasets, vendor portals with high share traffic), raise the thresholds via tenant overrides:

```jsonc
// tenant-overrides/research-tenant.json
{
  "IncidentResponse": {
    "MassDownloadFileCount":     500,
    "MassDownloadWindowMinutes": 60,
    "MassShareCount":            100
  }
}
```

## False positives

When a finding is a false positive, **do not just dismiss the alert** — close the auto-opened Low incident with `-FalsePositive`:

```powershell
Close-Incident -Id INC-... -Resolution "Legitimate user travel to Frankfurt; manager confirmed." -FalsePositive
```

`-FalsePositive` is a separate flag from `closed` so reporting tooling can distinguish "we handled this and it was real" from "this should never have fired." Over time these false-positive labels are the data you'd use to tune your thresholds.

## See also

- [`incident-response.md`](incident-response.md) — the playbook the framework opens incidents against.
- [`incident-runbook-template.md`](incident-runbook-template.md) — printable IR runbook.
- [`scheduled-checks.md`](scheduled-checks.md) — wiring health checks into Task Scheduler.
- [`pre-merge-review.md`](pre-merge-review.md) §1 — the now-closed retrofit item for tenant-scoped incidents.
