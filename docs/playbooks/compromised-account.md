# Compromised-account paper runbook

A printable plain-English walkthrough of the M365 Manager incident-response playbook. Designed for compliance audits that want documented IR process and for IR team members who need a step-by-step away from the keyboard.

This is the "what does the tool do" companion to [`incident-response.md`](../guides/incident-response.md) (which is the operator reference).

---

## When to invoke

Trigger the playbook when any of these is true:

- A user reports they think their account is compromised.
- Identity Protection flags a high-risk sign-in.
- The auto-detection framework (`Invoke-IncidentDetectors`) surfaces a finding above the team's threshold.
- A peer reports unusual mail / chat / share activity from the user.
- The SOC sees outbound traffic matching a known phishing campaign and the user is among the recipients.

## Severity decision

| If the situation looks like… | Use severity |
|---|---|
| Just one suspicious sign-in, user hasn't confirmed compromise | **Low** (forensic only) |
| User confirms compromise OR Identity Protection says high-risk | **Medium** (contain) |
| User confirmed compromise OR clear evidence of misuse | **High** *(default)* |
| Active AiTM kit / sent mail to dozens of internal users / mass file download in progress | **Critical** (full + quarantine prompt) |

When in doubt, start at **Low** to capture the forensic snapshot, then escalate.

## What happens at each severity

### Low — forensic only

The tool captures:
- The user's current account state (enabled? last sign-in? from where?)
- Group memberships
- Assigned licenses
- Registered MFA methods
- Mailbox forwarding configuration
- Active inbox rules
- Recent (24h) sign-in history

It then runs three audits:
- The last 24 hours of UAL + sign-in activity
- The last 7 days of sent mail (including who was outside the user's domain)
- The last 7 days of outbound SharePoint shares (including external domains)

And writes an HTML report.

**No state changes happen.** This is safe to run while you decide whether to escalate.

### Medium — contain

Everything Low does, plus:

- **Block sign-in.** The user can no longer authenticate.
- **Revoke all active sessions.** Any logged-in device is signed out within ~15 minutes.
- **Revoke all MFA methods.** The Authenticator app, FIDO key, phone number — all gone. Recovery requires a Temporary Access Pass + re-enrollment.
- **Force a password change + reset the password to a random 24-char value.** The new password is delivered via clipboard (interactive) or written to a `chmod 600` file in the incident directory (non-interactive).

After Medium, the user cannot get back in without operator action.

### High *(default)* — standard response

Everything Medium does, plus:

- **Capture every inbox rule** the user had configured, then disable each one. Preserves the evidence (the AiTM rule signature is forensically important) and is reversible — `Undo-Incident` re-enables them one at a time.
- **Clear all mailbox forwarding.** Both internal and external. Snapshot preserves the prior values.
- **Notify the security team** via configured channels (email + Teams webhook + whatever else `Send-Notification` routes to).

### Critical — full

Everything High does, plus the tool **asks** whether to quarantine the user's sent mail from the last 7 days. The operator must explicitly type `PURGE` to confirm — even in unattended automation runs, this step short-circuits with a "manual step required" entry rather than auto-purging.

If the operator confirms, the tool:
- Creates a Compliance Search filtered to the user's last 7 days of sent mail
- Starts the search
- Issues a HardDelete purge action

This is **irreversible**. Messages purged this way are removed from the tenant permanently — restoring them is not possible.

## Recovery procedure

When the legitimate user is being returned to their account:

1. **Verify the user out-of-band** (phone, in-person). Do not use email — the attacker may still have access to forwarded mail.
2. **Issue a Temporary Access Pass** via the menu (option 16 → MFA → TAP) or `New-TemporaryAccessPass`. This lets the user re-enroll MFA without password access.
3. **Re-enable sign-in** by running `Undo-Incident -Id <id>`. The tool walks the reversible steps in order and asks per-step confirmation.
4. **Confirm identity in person** during MFA re-enrollment if possible. AiTM kits sometimes follow up with social engineering at the recovery moment.
5. **Close the incident** with a resolution note. The tool appends a `status=closed` record to the incident registry.

For false-positive incidents:

```
Close-Incident -Id INC-... -Resolution "Legitimate travel" -FalsePositive
```

The tool auto-walks the reversal steps when `-FalsePositive` is set.

## Artifacts produced

For every incident:

| File | What's in it |
|---|---|
| `snapshot.json` | Pre-incident user state. Forensic baseline. |
| `inbox-rules.json` | Mail rules at time of incident. |
| `audit-24h.json` | UAL + sign-in activity for the 24h before the incident opened. |
| `mail-sent-7d.json` | Sent items for the 7d before the incident, with `externalRecipients` rollup. |
| `shares-7d.json` | Outbound SharePoint shares for 7d, with `externalDomains` rollup. |
| `report.html` | Operator + auditor view. Step outcomes, findings, recommended next steps, links to artifacts. |
| `temp-password.txt` | Only created in non-interactive mode when the playbook forced a password change. `chmod 600`. |

For tabletop exercises:

| File | What's in it |
|---|---|
| `tabletop-report.html` | Grade against the scenario's `expectedActions` list, wall-clock timing, list of playbook steps observed. |

## Audit log

Every step is individually audited with the `incidentId` in the `target` field. The audit log lives at `%LOCALAPPDATA%\M365Manager\audit\session-<ts>-<pid>-<tenant>.log` (or `~/.m365manager/audit/` on POSIX). One JSON object per line; structured fields described in [`audit-format.md`](../reference/audit-format.md).

Filter the audit log to one incident:

```
Read-AuditEntries | Where-Object { $_.target.incidentId -eq 'INC-2026-05-14-a3f2' }
```

Or use the menu: option 14 → Audit & Reporting → filter by Target = `INC-...`.

## Retention

Incident artifacts are retained for **365 days** by default. Configure via `IncidentResponse.SnapshotRetentionDays` in `ai_config.json`. The tool does NOT auto-prune — integrate with your organization's compliance retention policy. For longer retention, copy the incident directories out to a write-once archive.

## Compliance handoff

```
Export-Incident -Id INC-2026-05-14-a3f2 -Path /secure/handoff/bundle.zip
```

Bundles the incident directory + a filtered slice of the audit log (this incident's lines only) + an `INDEX.json` with the registry record. Hand off via secure channel; bundles may contain sensitive forensic data.

## What the playbook does NOT do

- **It does not contact the user.** Notifying the legitimate user is the operator's job — and should happen out-of-band, not via the compromised account's email.
- **It does not audit downstream SaaS.** Slack, GitHub, Salesforce, the CRM — those are systems M365 Manager doesn't reach. Audit them separately.
- **It does not revoke OAuth grants.** Apps the user consented to retain their tokens until those are explicitly revoked via the Entra portal (Apps → User consented apps → Revoke).
- **It does not modify Conditional Access policies.** If a policy let the attacker through, fix the policy separately — the playbook recovers the affected account but doesn't address the upstream gap.

## Tabletop process

For compliance audits demonstrating IR readiness:

1. Pick a scenario from `templates/tabletop-scenarios/`.
2. Designate a sandbox account (`IncidentResponse.TabletopUPN`).
3. Run `Invoke-IncidentTabletop -ScenarioName <name>`.
4. Review the `tabletop-report.html` — it grades the run against the scenario's expected actions.
5. File the report with the compliance evidence.

The tabletop runs in PREVIEW so no tenant state changes. The sandbox account is never actually mutated.

## See also

- [`incident-response.md`](../guides/incident-response.md) — operator reference, full step-by-step technical detail.
- [`incident-triggers.md`](incident-triggers.md) — auto-detection framework.
- [`audit-format.md`](../reference/audit-format.md) — audit log schema.
