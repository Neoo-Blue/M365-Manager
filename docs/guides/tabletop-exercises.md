# Tabletop exercises

Run an incident-response scenario against a sandbox user in PREVIEW mode. Grades the IR team's reaction against a documented expected-actions list. Designed for compliance audits demonstrating IR readiness + for IR team training.

## When to use this

- Quarterly IR drill — demonstrate documented process for your auditors.
- Onboarding a new IR team member — let them run the playbook end-to-end without risk.
- After a real incident — re-run the matching scenario to validate the response would have been thorough.

## Prereqs

- Phase 7 incident-response module loaded.
- A sandbox account in your tenant designated for IR exercises (e.g. `ir-sandbox@yourdomain.com`). The scenario will run the full playbook against this account in PREVIEW so it's never actually mutated.
- Configure the default sandbox via `IncidentResponse.TabletopUPN` in `ai_config.json` so future runs don't need the flag.

## Available scenarios

Four ship in `templates/tabletop-scenarios/`:

| Name | Severity | Triggers | Notes |
|---|---|---|---|
| `phishing-campaign` | High | Identity Protection high-risk + SuspiciousInboxRule | Standard credential phish without AiTM. |
| `insider-mass-download` | High | MassFileDownload | Departing employee exfil. **`expectedActions` correctly OMITS `ForcePasswordChange`** — legal hold means you don't want to lose the password. |
| `mfa-bypass` | Critical | AnomalousLocationSignIn + SuspiciousInboxRule | AiTM with session-cookie theft. |
| `compromised-vendor` | Medium | MassExternalShare for a Guest user | Compromised B2B account. **`expectedActions` correctly OMITS `RevokeAuthMethods` / `ForcePasswordChange`** — the vendor's auth lives in their home tenant. |

Each scenario is a JSON file declaring:

```jsonc
{
  "name":            "phishing-campaign",
  "description":     "Active credential-phishing campaign hits one or more accounts...",
  "severity":        "High",
  "trigger":         "Identity Protection flags high-risk sign-in plus a SuspiciousInboxRule detector match.",
  "expectedActions": [
    "Snapshot", "BlockSignIn", "RevokeSessions", "RevokeAuthMethods",
    "ForcePasswordChange", "DisableInboxRule", "ClearForwarding",
    "Audit24h", "AuditSentMail", "AuditShares", "Notify", "Report"
  ],
  "expectedFindings": [
    "Sign-in from unusual location",
    "Suspicious inbox rule (filter + delete + forward)",
    "Outbound mail to other internal users (phishing propagation)"
  ],
  "recoveryNotes":   "Issue Temporary Access Pass to the legitimate user..."
}
```

## Running a scenario

```powershell
Invoke-IncidentTabletop -ScenarioName phishing-campaign

# Or with an explicit sandbox UPN:
Invoke-IncidentTabletop -ScenarioName mfa-bypass -TabletopUPN ir-sandbox@contoso.com
```

What the tool does:

1. Loads the scenario JSON from `templates/tabletop-scenarios/`.
2. Calls `Invoke-CompromisedAccountResponse` against the sandbox UPN with `-WhatIf -NonInteractive`. PREVIEW mode, no actual mutations.
3. Reads back the audit log entries tied to the resulting incident id.
4. Compares the observed `Incident:<Step>` action types against the scenario's `expectedActions` list.
5. Grades: "N of M actions observed; missing actions: X, Y."
6. Writes `tabletop-report.html` to the incident snapshot directory.

Output:

```
[Incident tabletop -- mfa-bypass]
  Scenario  : AiTM with session-cookie theft. MFA looks satisfied but session is attacker's.
  Severity  : Critical
  Sandbox   : ir-sandbox@contoso.com

  [WhatIf] running playbook in PREVIEW...
  ... 12 steps complete in 8.3s

  TABLETOP GRADE: 11/12
  Wall-clock: 8.3s
  Steps observed: 12
  Missing actions (1): QuarantineSentMail

  Tabletop report: <stateDir>/contoso/incidents/INC-2026-05-14-xxxx/tabletop-report.html
```

The "missing action" here is expected — `QuarantineSentMail` is opt-in via `-QuarantineSentMail` even at Critical severity, so the tabletop run (without that flag) won't fire it.

## What the report contains

The tabletop-report.html:

- Scenario summary + expected actions.
- Grade (N/M).
- Wall-clock timing.
- Observed `Incident:<Step>` audit entries.
- Missing actions with a "why this might be missing" note.
- Compliance note explaining that PREVIEW was used + no tenant state changed.
- Link to the incident's full report.html (which has the snapshot + audits).

## Compliance handoff

For a quarterly drill demonstrating IR readiness:

1. Run all four shipping scenarios.
2. `Export-Incident -Id <each-incident-id> -Path /secure/quarterly-drill-q1.zip`.
3. Hand off the bundles to the compliance team along with the tabletop reports.

Each bundle includes the synthetic snapshot, the simulated audits, the report, and the filtered audit-log slice.

## Adding a custom scenario

Drop a new JSON into `templates/tabletop-scenarios/`. Schema:

| Field | Required | Notes |
|---|---|---|
| `name` | yes | Lowercase, hyphen-separated. Matches the filename minus `.json`. |
| `description` | yes | One-sentence summary. |
| `severity` | yes | `Low` / `Medium` / `High` / `Critical`. |
| `trigger` | yes | One-line note about what triggers this scenario. |
| `expectedActions` | yes | Array of `Incident:<Step>` action types (drop the `Incident:` prefix). |
| `expectedFindings` | no | Bullet points the playbook should surface. |
| `recoveryNotes` | no | Free-text guidance for the operator at the end. |

The available step names: `Snapshot`, `BlockSignIn`, `RevokeSessions`, `RevokeAuthMethods`, `ForcePasswordChange`, `DisableInboxRule`, `ClearForwarding`, `Audit24h`, `AuditSentMail`, `AuditShares`, `QuarantineSentMail`, `Notify`, `Report`.

## Sandbox account hygiene

The tabletop runs in PREVIEW. The sandbox UPN is NEVER actually mutated. But it IS read repeatedly — the snapshot pulls its current state, sign-ins, MFA methods, etc. Implications:

- The sandbox user should exist (real Entra user). The tool doesn't auto-create.
- Pick a low-privilege account. Even a read-only access leaks if the wrong logs go to a wrong place.
- Don't share the sandbox across tenants (the per-tenant audit log makes this unnecessary anyway).

## See also

- [`incident-response.md`](incident-response.md) — the playbook the tabletop exercises.
- [`../playbooks/compromised-account.md`](../playbooks/compromised-account.md) — paper companion to the playbook.
- [`../playbooks/incident-triggers.md`](../playbooks/incident-triggers.md) — auto-detection framework.
