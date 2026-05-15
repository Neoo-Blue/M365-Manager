# Incident response (Phase 7)

The compromised-account playbook is the highest-blast-radius operation in M365 Manager. This doc covers the full reference: severity matrix, step-by-step behavior, recovery, retention, and compliance handoff.

## Quick start

```powershell
# Single user, default High severity, interactive confirmation
Invoke-CompromisedAccountResponse -UPN alice@contoso.com -Reason "Phishing campaign"

# From the AI assistant
/incident alice@contoso.com High

# Or the menu: option 22 -> "Run compromised-account response"
```

Returns an incident id like `INC-2026-05-14-a3f2`. All downstream view / replay / undo functions key off this id.

## Severity matrix

The playbook is 13 steps. Severity gates which steps run; the operator can ratchet up or down at runtime.

| Severity | Steps | Use case |
|---|---|---|
| **Low** | 1, 8, 9, 10, 13 | Forensic only. "I want to see what this account has been doing without taking action yet." No state changes. |
| **Medium** | Low + 2, 3, 4, 5 | Contain. Block + revoke + rotate. No mailbox cleanup, no notify. |
| **High** *(default)* | Medium + 6, 7, 12 | Standard response. Adds inbox-rule disable, forwarding clear, security team notification. |
| **Critical** | High + step 11 default-on prompt | Full response. Quarantine prompt asks the operator to type PURGE for compliance-purge of 7d sent mail. |

## The 13 steps

1. **Snapshot** *(always)* — capture user core, manager, groups, licenses, MFA methods, mailbox forwarding, inbox rules, recent sign-ins. Written to `snapshot.json` BEFORE any mutation so the forensic baseline survives even if a later step fails. Read-only.
2. **BlockSignIn** *(Medium+)* — `PATCH accountEnabled=false`. Reverse: `UnblockSignIn`.
3. **RevokeSessions** *(Medium+)* — `POST revokeSignInSessions`. Marked `NoUndoReason` — sessions are gone, the next sign-in re-authenticates anyway.
4. **RevokeAuthMethods** *(Medium+)* — calls `MFAManager.Remove-AllAuthMethods`. `NoUndoReason`; recovery via operator-driven TAP + re-enrollment.
5. **ForcePasswordChange** *(Medium+)* — `passwordProfile.forceChangePasswordNextSignIn=true` + a 24-char random password. Delivered via `Set-Clipboard` interactively or as a `chmod 600` `temp-password.txt` in NonInteractive mode.
6. **DisableInboxRules** *(High+)* — captures all rules into `inbox-rules.json` then PATCHes `isEnabled=false` on each previously-enabled rule (each as its own `Invoke-Action` with a `ReverseType=EnableInboxRule` recipe). Disabled-rule preservation > deletion — keeps the evidence.
7. **ClearForwarding** *(High+)* — `Set-Mailbox -ForwardingAddress $null -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false`. Reverse: `SetForwarding` from the snapshot's prior values.
8. **Audit24h** *(always)* — sign-ins (`Search-SignIns`) + UAL (`Search-UAL`) for the last 24 hours, written to `audit-24h.json`.
9. **AuditSentMail** *(always)* — last 7d sent items via Graph `/messages?$filter=sentDateTime ge ...`, written to `mail-sent-7d.json`. Includes an `externalRecipients` list (recipients outside the user's domain) — the phishing-propagation surface.
10. **AuditShares** *(always)* — last 7d outbound shares via `SharePoint.Get-UserOutboundShares`, written to `shares-7d.json`. Includes an `externalDomains` rollup — the data-exfil surface.
11. **QuarantineSentMail** *(Critical, opt-in via `-QuarantineSentMail`)* — `New-ComplianceSearch` → `Start-ComplianceSearch` → `New-ComplianceSearchAction -Purge -PurgeType HardDelete`. **Always requires the operator to type `PURGE`** — short-circuits with a "manual step required" entry in non-interactive mode rather than auto-purging. `NoUndoReason` — purges are irreversible.
12. **Notify** *(High+)* — `Send-Notification -Channels SecurityTeam -Severity Critical`. Body links to the incident report.
13. **Report** *(always)* — `report.html` summarizing every step's outcome + findings + recommended next steps.

## Recovery procedure

When the legitimate user has been contacted and the account is being returned to them:

1. **Verify the user** out-of-band (phone, in-person). Do NOT use email — the attacker may still have access through forwarded mail.
2. **Issue a Temporary Access Pass** via the menu (option 16 → MFA → Issue TAP) or `New-TemporaryAccessPass -User alice@contoso.com -LifetimeMinutes 60`. This lets the user re-enroll MFA without password access.
3. **Re-enable sign-in** by running `Undo-Incident -Id INC-...`. Walks the reversible steps and asks per-step confirmation.
4. **Confirm the user's identity** in person if possible during re-enrollment. AiTM kits often follow up with social-engineering attempts at the recovery moment.
5. **Close the incident**:
   ```powershell
   Close-Incident -Id INC-2026-05-14-a3f2 -Resolution "User confirmed identity, re-enrolled MFA in person on 2026-05-15. New password set, OAuth grants reviewed clean."
   ```

For false-positive incidents (detector tripped on legitimate activity):

```powershell
Close-Incident -Id INC-... -Resolution "Legitimate travel; confirmed with manager." -FalsePositive
# This auto-walks Undo-Incident for the reversible steps.
```

## Bulk response

For phishing campaigns that hit multiple accounts:

```powershell
Invoke-BulkIncidentResponse -Path incidents.csv
```

CSV columns: `UPN, Severity, Reason, QuarantineSentMail`. Sample at `templates/incidents-bulk-sample.csv`.

Validate-first pattern from Phase 1: per-row failures don't halt the batch. Result CSV written next to the input. Aggregate HTML at `<stateDir>/<tenant>/incidents/bulk-<ts>/index.html` linking each sub-incident's individual report.

## Tabletop exercises

Run a scenario in PREVIEW against a sandbox user:

```powershell
Invoke-IncidentTabletop -ScenarioName phishing-campaign -TabletopUPN ir-sandbox@yourdomain.com
```

Four scenarios ship in `templates/tabletop-scenarios/`:

- **phishing-campaign** — standard credential-phish (High)
- **insider-mass-download** — departing-employee exfil (High; correctly omits `ForcePasswordChange` for legal hold)
- **mfa-bypass** — AiTM with session-cookie theft (Critical)
- **compromised-vendor** — compromised Guest user (Medium; correctly omits auth-method revocation since the vendor's auth lives in their home tenant)

Each scenario declares `expectedActions` — the playbook is graded ("11/12 actions") against the actual run.

## On-disk layout

Tenant-scoped under `<stateDir>/<tenant-slug>/`:

```
<stateDir>/contoso/
├── incidents.jsonl                          # append-only registry
└── incidents/
    └── INC-2026-05-14-a3f2/
        ├── snapshot.json                    # pre-incident state
        ├── inbox-rules.json                 # rules at time of incident
        ├── audit-24h.json                   # sign-ins + UAL
        ├── mail-sent-7d.json                # sent mail with external rollup
        ├── shares-7d.json                   # outbound shares with domain rollup
        ├── report.html                      # operator + auditor view
        └── temp-password.txt                # only if non-interactive forced password change ran
```

`<stateDir>` resolves to `%LOCALAPPDATA%\M365Manager\state` on Windows or `~/.m365manager/state` (chmod 700) on POSIX.

## Retention

Snapshot retention defaults to **365 days** (`IncidentResponse.SnapshotRetentionDays`). The tool does not auto-prune — operators are expected to integrate with their organization's compliance retention policy. For longer retention, copy the directories out to a write-once archive.

## Compliance handoff

```powershell
Export-Incident -Id INC-2026-05-14-a3f2 -Path /secure/handoff/incident-bundle.zip
```

Bundles:
- The full incident directory (snapshot, audits, report, etc.)
- A filtered slice of the session audit log containing only this incident's lines
- An `INDEX.json` with the registry record + handoff metadata

The bundle may contain sensitive forensic data (UPNs, sign-in IPs, mail subjects). Hand off via secure channel only.

## See also

- [`docs/incident-runbook-template.md`](incident-runbook-template.md) — printable paper runbook for compliance audits demonstrating IR process.
- [`docs/incident-triggers.md`](incident-triggers.md) — auto-detection framework details.
- [`docs/sample-incident/`](sample-incident/) — anonymized example for training / demos.
- [`docs/audit-format.md`](audit-format.md) — the JSONL audit format the playbook emits to.
- [`docs/pre-merge-review.md`](pre-merge-review.md) — what was deferred in earlier phases (incidents-per-tenant retrofit landed here).
