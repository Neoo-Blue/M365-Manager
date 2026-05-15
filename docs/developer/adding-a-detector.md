# Adding an incident-response detector

Phase 7's detector framework looks for suspicious activity and auto-opens a forensic incident. This guide walks through adding a new detector.

## When to add a detector

Add when:

- You have a structured query that reliably flags suspicious activity (Graph API, UAL, sign-in log).
- The signal is high-precision (false-positive rate < 10%) — every detector firing opens an incident + alerts the team.
- You can articulate `Severity` (Low / Medium / High / Critical) per finding.

Don't add when:

- The signal needs human judgment to assess (then it's a manual procedure, not a detector).
- The query has unbounded cost (paging through millions of UAL rows per user per 15 minutes).
- The detector requires Graph SDK / EXO / SCC scopes the tool doesn't already need.

## Detector contract

Every detector is a function that:

- Takes `[Parameter(Mandatory)][string]$UPN`.
- Optionally takes pre-fetched data (sign-ins, UAL rows) to skip the Graph call when the driver already has them.
- Returns either `$null` (nothing suspicious) or a `Finding` hashtable.

The `Finding` hashtable shape:

```powershell
@{
    TriggerType        = 'YourDetectorName'           # PascalCase
    UPN                = $upn
    Severity           = 'Low' | 'Medium' | 'High' | 'Critical'
    Evidence           = @{ ... detector-specific structured data ... }
    RecommendedAction  = 'Operator-facing one-liner.'
    DetectedUtc        = (Get-Date).ToUniversalTime().ToString('o')
}
```

Use the helper `New-IncidentFinding` to build it:

```powershell
return New-IncidentFinding -TriggerType 'YourDetectorName' -UPN $UPN -Severity 'High' `
    -Evidence @{ kind = 'a thing'; count = 42 } `
    -RecommendedAction "Run /incident $UPN High."
```

## Step 1 — design the predicate

What signal triggers this? Pick something:

- **Specific.** "Forward-to-external + delete-from-sent + degenerate rule name" is specific. "Suspicious behavior" is not.
- **Cheap to compute.** A single Graph call per UPN per 15 minutes is fine. Loading every UAL row from the last 90 days per user is not.
- **Tunable.** The thresholds should be config-overridable so operators with unusual workloads (research, M&A, vendor-heavy) can adjust.

For our example: "user signed in but then immediately rejected a Conditional Access policy 5+ times in a window." This is a CA-bypass-attempt signal.

## Step 2 — pick the severity

| Severity | When to use |
|---|---|
| `Low` | Forensic capture only. Possible suspicious; operator should investigate manually. |
| `Medium` | Suggestive. Operator should consider containment if context corroborates. |
| `High` | Strong signal. Default action: run the playbook at High severity. |
| `Critical` | Smoking gun. AiTM kit signature, mass exfil in progress, etc. Default action: run the playbook at Critical with `-QuarantineSentMail`. |

For "CA bypass attempts": **High**. Multiple deliberate CA rejections is unusual and a strong signal.

## Step 3 — pick threshold config keys

Every detector should be tunable. Pick keys like:

- `IncidentResponse.CABypassAttemptCount` (default 5)
- `IncidentResponse.CABypassWindowMinutes` (default 15)

Add defaults to `Get-IncidentTriggerConfig` in `IncidentTriggers.ps1`:

```powershell
$defaults = @{
    # ... existing ...
    CABypassAttemptCount   = 5
    CABypassWindowMinutes  = 15
}
```

The resolver pulls from `ai_config.json`'s `IncidentResponse` block, then per-tenant override, then env, then CLI — see [`../concepts/tenant-overrides.md`](../concepts/tenant-overrides.md).

## Step 4 — write the function

Add to `IncidentTriggers.ps1`:

```powershell
function Detect-CABypassAttempts {
    <#
        N+ Conditional Access rejections in M minutes signals a
        deliberate bypass attempt -- the user (or an attacker
        with the user's password) is trying multiple ways to
        slip past the policy.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [array]$SignIns
    )
    $cfg = Get-IncidentTriggerConfig
    if (-not $SignIns) {
        if (-not (Get-Command Search-SignIns -ErrorAction SilentlyContinue)) { return $null }
        $windowMin = [int]$cfg.CABypassWindowMinutes
        try { $SignIns = @(Search-SignIns -User $UPN -From (Get-Date).AddMinutes(-1 * ($windowMin + 5)) -MaxResults 50) }
        catch { return $null }
    }
    $rejected = @($SignIns | Where-Object {
        $_.ConditionalAccessStatus -eq 'failure' -or $_.Status -match '530002'
    })
    if ($rejected.Count -lt [int]$cfg.CABypassAttemptCount) { return $null }

    return New-IncidentFinding -TriggerType 'CABypassAttempts' -UPN $UPN -Severity 'High' `
        -Evidence @{
            rejectionCount = $rejected.Count
            windowMinutes  = [int]$cfg.CABypassWindowMinutes
            threshold      = [int]$cfg.CABypassAttemptCount
            firstRejection = [string]$rejected[0].CreatedDateTime
            lastRejection  = [string]$rejected[-1].CreatedDateTime
        } `
        -RecommendedAction "Possible CA bypass attempt. Run /incident $UPN High and review CA policy."
}
```

Patterns to follow:

- **Take optional pre-fetched data** via the `-SignIns` parameter. The driver `Invoke-IncidentDetectors` fetches once and passes to every detector for the same user — avoids redundant Graph calls.
- **Tolerate missing data** — if the user's never signed in, `$SignIns` is empty. Return `$null`, don't error.
- **Wrap the Graph call in try/catch** — transient errors shouldn't crash the sweep.
- **Use config-resolved thresholds**, not hard-coded numbers.

## Step 5 — register in the driver

Add to `Invoke-IncidentDetectors`'s `$detectors` array:

```powershell
$detectors = @(
    @{ Name = 'AnomalousLocationSignIn'; Fn = ${function:Detect-AnomalousLocationSignIn} }
    @{ Name = 'ImpossibleTravel';        Fn = ${function:Detect-ImpossibleTravel}        }
    @{ Name = 'HighRiskSignIn';          Fn = ${function:Detect-HighRiskSignIn}          }
    @{ Name = 'MassFileDownload';        Fn = ${function:Detect-MassFileDownload}        }
    @{ Name = 'MassExternalShare';       Fn = ${function:Detect-MassExternalShare}       }
    @{ Name = 'SuspiciousInboxRule';     Fn = ${function:Detect-SuspiciousInboxRule}     }
    @{ Name = 'MFAFatigue';              Fn = ${function:Detect-MFAFatigue}              }
    @{ Name = 'CABypassAttempts';        Fn = ${function:Detect-CABypassAttempts}        }   # <-- new
)
```

## Step 6 — Pester tests

Add to `tests/IncidentTriggers.Tests.ps1`:

```powershell
Describe "Detect-CABypassAttempts" {
    It "fires when threshold is met" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T10:00:00Z'; ConditionalAccessStatus = 'failure'; Status = '530002 (CA blocked)' },
            @{ CreatedDateTime = '2026-05-14T10:02:00Z'; ConditionalAccessStatus = 'failure'; Status = '530002 (CA blocked)' },
            @{ CreatedDateTime = '2026-05-14T10:04:00Z'; ConditionalAccessStatus = 'failure'; Status = '530002 (CA blocked)' },
            @{ CreatedDateTime = '2026-05-14T10:06:00Z'; ConditionalAccessStatus = 'failure'; Status = '530002 (CA blocked)' },
            @{ CreatedDateTime = '2026-05-14T10:08:00Z'; ConditionalAccessStatus = 'failure'; Status = '530002 (CA blocked)' }
        )
        $f = Detect-CABypassAttempts -UPN 'u@x.com' -SignIns $signIns
        $f                       | Should -Not -BeNullOrEmpty
        $f.TriggerType          | Should -Be 'CABypassAttempts'
        $f.Severity             | Should -Be 'High'
        $f.Evidence.rejectionCount | Should -Be 5
    }
    It "does NOT fire below threshold" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T10:00:00Z'; ConditionalAccessStatus = 'failure'; Status = '530002' },
            @{ CreatedDateTime = '2026-05-14T10:02:00Z'; ConditionalAccessStatus = 'failure'; Status = '530002' }
        )
        (Detect-CABypassAttempts -UPN 'u@x.com' -SignIns $signIns) | Should -BeNullOrEmpty
    }
    It "honors config-tuned threshold" {
        function global:Get-EffectiveConfig {
            param([string]$Key)
            if ($Key -eq 'IncidentResponse.CABypassAttemptCount') { return 10 }
            return $null
        }
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T10:00:00Z'; ConditionalAccessStatus = 'failure' },
            # ... 5 failures
            @{ CreatedDateTime = '2026-05-14T10:08:00Z'; ConditionalAccessStatus = 'failure' }
        )
        # With threshold raised to 10, 5 rejections should NOT fire:
        (Detect-CABypassAttempts -UPN 'u@x.com' -SignIns $signIns) | Should -BeNullOrEmpty
        Remove-Item Function:Get-EffectiveConfig -ErrorAction SilentlyContinue
    }
}
```

## Step 7 — config + docs

Add the new config keys to `ai_config.example.json` under `IncidentResponse`:

```jsonc
"IncidentResponse": {
    // ... existing keys ...
    "CABypassAttemptCount":  5,
    "CABypassWindowMinutes": 15
}
```

Update [`../playbooks/incident-triggers.md`](../playbooks/incident-triggers.md) with the new detector row in the "seven detectors" table (now "eight").

Update [`../reference/config-keys.md`](../reference/config-keys.md) with the new keys.

## Severity vs auto-execute interaction

The detector framework's `AutoExecuteOnSeverity` config key controls whether findings auto-run the destructive playbook:

| `AutoExecuteOnSeverity` | Behavior for a `High` finding |
|---|---|
| `None` (default) | Open a Low forensic incident + alert. Operator decides whether to escalate. |
| `Critical` | Same as `None` for High findings. |
| `HighAndCritical` | Auto-run the destructive playbook at `High` severity (no confirmation). |

Authors should default-design for `None`. If a detector's signal is so strong that operators reasonably want auto-containment, it should justify `Critical` severity, not push the org toward `HighAndCritical`.

## Pre-merge fix history

The detector framework caught a hashtable-vs-PSCustomObject issue in v1 — `Get-IncidentSignInCountry` was added to handle both shapes after the first Pester pass found that `Location.countryOrRegion` on a hashtable doesn't work the same as on a PSCustomObject. If your detector touches structured Graph data, use the existing helpers (or add similar ones).

## See also

- [`../playbooks/incident-triggers.md`](../playbooks/incident-triggers.md) — operator-facing detector reference.
- [`../guides/incident-response.md`](../guides/incident-response.md) — the playbook the detector framework feeds into.
- [`adding-a-module.md`](adding-a-module.md) — for new modules entirely.
- [`testing.md`](testing.md) — Pester patterns for mocking Graph.
