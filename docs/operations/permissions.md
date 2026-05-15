# Permissions

Graph scopes + Entra roles + EXO / SPO / SCC role matrix, per feature area. Use this to grant a connecting account the minimum permission needed for the operations you intend to do.

## Trust model

The tool runs as the connecting account — every API call inherits that account's permission. There's no service-account-elevation trick. To minimize blast radius:

1. Pick the **most granular** role set that covers your needs (avoid Global Admin unless you have to).
2. Use **Privileged Identity Management (PIM)** if your tenant has it — request the role for the duration of the operation.
3. For unattended automation, use a **CertThumbprint tenant profile** (app-only auth) with the application permissions narrowed to what's needed.

## Read-only baseline

For anyone who just runs reports + audit views:

| Role | What it covers |
|---|---|
| `Global Reader` | Read everything (the entire tenant). |
| `Security Reader` | Read security-related blades (sign-in logs, risk events). |

Graph scopes for the app: `Directory.Read.All`, `User.Read.All`, `Group.Read.All`, `AuditLog.Read.All`.

This baseline runs:
- Slot 11 (Reporting)
- Slot 14 (Audit & Reporting, sign-in lookup, UAL search)
- Slot 18 (SharePoint, list-only)
- Slot 19 (Guest discovery, recertification list, no removal)
- Phase 7 incident response at `Low` severity (snapshot + 3 audits, no mutations)

## Write — narrow

If you only do one or two areas:

| Feature | Minimum role |
|---|---|
| User onboarding / offboarding | `User Administrator` + `License Administrator` |
| License management | `License Administrator` |
| Group / DL management | `Group Administrator` + `Exchange Recipient Administrator` |
| MFA management | `Authentication Administrator` |
| SharePoint sites + shares | `SharePoint Administrator` |
| Teams ownership / membership | `Teams Administrator` |
| Compliance search + purge | `eDiscovery Manager` + `Compliance Administrator` |
| Conditional Access (not directly mutated by this tool today) | `Conditional Access Administrator` |
| Scheduled task management (host-side) | Workstation local admin |

## Write — broad

To run the whole tool without role-juggling:

| Role | Scope |
|---|---|
| `Privileged Role Administrator` | Manage role assignments. |
| `Global Administrator` | Everything. Use sparingly + with PIM. |

## Phase-by-phase permissions

### Phase 0.5 (security hardening)

No tenant permissions beyond the baseline — these are local-machine features (DPAPI, audit log, redaction).

### Phase 1 (onboard / offboard / preview)

- `User Administrator` + `License Administrator` for user creation + license assignment.
- `Group Administrator` for group / DL membership.
- `Exchange Administrator` for mailbox-type changes (Shared / Resource).
- `SharePoint Administrator` for OneDrive provisioning checks during onboard.

Graph scopes: `User.ReadWrite.All`, `Directory.ReadWrite.All`, `Group.ReadWrite.All`, `MailboxSettings.ReadWrite`.

### Phase 2 (audit + MFA + lookups)

- `Authentication Administrator` (full MFA mgmt) or `Authentication Policy Administrator` (read-only) for MFA flows.
- `Audit Logs` role in Compliance Center for UAL search.
- `Security Reader` for sign-in log reads.

Graph scopes: `UserAuthenticationMethod.ReadWrite.All`, `AuditLog.Read.All`.

### Phase 3 (lifecycle completeness)

- `SharePoint Administrator` for outbound-share audit + revocation.
- `Teams Administrator` for membership / ownership / orphan reports.
- `Guest Inviter` + `User Administrator` for guest lifecycle.
- `Exchange Administrator` for OneDrive retention adjustments (`employeeLeaveDateTime`).

### Phase 4 (cost + health)

- `License Administrator` for the optimizer.
- `Reports Reader` for license usage reports.
- Workstation local admin (or the SYSTEM account, depending on the scheduled task setup) for Phase 4 scheduler.
- The Notifications module sends from the connecting account's mailbox via `/me/sendMail` by default — no extra Graph scope needed beyond `Mail.Send`.

### Phase 5 (AI v2)

No tenant permissions — the AI module is a chat layer over the same primitives the rest of the tool already needs.

### Phase 6 (multi-tenant)

For each tenant the operator registers:

| Tenant auth mode | Required setup |
|---|---|
| Interactive | Same as direct admin per the role matrix above. |
| CertThumbprint | Entra app registration in the target tenant with **application permissions** (not delegated) matching what you'll use. The Entra app's required permissions must be **admin-consented**. See [`../getting-started/tenant-setup.md`](../getting-started/tenant-setup.md). |
| ClientSecret | Same as CertThumbprint. |

For GDAP partner mode: GDAP-eligible delegated roles on the partner side (`User Administrator` etc. delegated from each customer). Partner Center role: `Admin Agent` to enumerate customers.

### Phase 7 (incident response)

Aggressive permission requirements — the playbook touches every domain.

| Step | Required permission |
|---|---|
| Snapshot | `Directory.Read.All` (Graph) + `Audit Logs` (Compliance) |
| BlockSignIn | `User Administrator` + `User.EnableDisableAccount.All` (Graph) |
| RevokeSessions | `User.RevokeSessions.All` (Graph) |
| RevokeAuthMethods | `Authentication Administrator` + `UserAuthenticationMethod.ReadWrite.All` (Graph) |
| ForcePasswordChange | `User Administrator` + `User.ReadWrite.All` (Graph) |
| DisableInboxRules | `Exchange Administrator` for the user + `MailboxSettings.ReadWrite` (Graph) |
| ClearForwarding | `Exchange Recipient Administrator` |
| Audit24h | `Audit Logs` (Compliance Center) + `AuditLog.Read.All` (Graph) |
| AuditSentMail | `Mail.Read` (Graph) + access to the user's mailbox (`ApplicationImpersonation` if running app-only) |
| AuditShares | `SharePoint Administrator` + `Audit Logs` (Compliance) |
| QuarantineSentMail | `eDiscovery Manager` + `Compliance Administrator` |
| Notify | None additional |
| Report | None additional |

The detector framework (Phase 7 Commit E) needs the same as the audit + snapshot rows above.

## Application permissions (app-only auth)

For CertThumbprint / ClientSecret tenant profiles, the Entra app registration needs **application** permissions (not delegated). Common set:

| Permission | Required for |
|---|---|
| `User.ReadWrite.All` | Onboarding, offboarding, profile management, incident response |
| `Directory.ReadWrite.All` | Group + role + tenant-level operations |
| `Group.ReadWrite.All` | Group + DL management |
| `MailboxSettings.ReadWrite` | Mailbox config (forwarding, OOO, inbox rules) |
| `Mail.Send` | Notifications |
| `Mail.Read` | Sent-mail audits (incident response step 9) |
| `AuditLog.Read.All` | Sign-in lookup + audit log reads |
| `UserAuthenticationMethod.ReadWrite.All` | MFA management |
| `User.RevokeSessions.All` | Session revocation (incident response step 3) |
| `User.EnableDisableAccount.All` | Block sign-in (incident response step 2) |
| `Sites.FullControl.All` | SharePoint operations |
| `TeamMember.ReadWrite.All` | Teams membership |

**Admin consent required** for every application permission. The Entra portal's "Grant admin consent" button on the app registration applies all at once.

## EXO / SCC role matrix

Some operations route through Exchange Online or Security & Compliance Center cmdlets (not Graph). These need RBAC roles in those services:

| Exchange role | Covers |
|---|---|
| `Mail Recipients` | Add / remove / set mailbox |
| `Distribution Groups` | DL membership |
| `Mailbox Search` | UAL search |
| `Mailbox Import Export` | Compliance purge (incident response step 11) |
| `Recipient Policies` | Inbox rules at scale |

| SCC role | Covers |
|---|---|
| `Compliance Administrator` | Compliance search + purge |
| `eDiscovery Manager` | eDiscovery operations |
| `Audit Logs` | UAL search via SCC |

## Verifying permissions

After granting roles, verify the connecting account can run a sample operation in PREVIEW:

```powershell
# Most permission failures show up loudly in PREVIEW even before any tenant call:
Set-PreviewMode -Enabled $true
Invoke-CompromisedAccountResponse -UPN test@yourdomain.com -Severity High
# Walk the [PREVIEW] output -- any "Insufficient privileges" failure
# indicates a missing role or scope for that specific step.
```

The [`validation-runbook.md`](validation-runbook.md) has a more thorough live-PREVIEW walkthrough covering every destructive flow.

## See also

- [`../getting-started/tenant-setup.md`](../getting-started/tenant-setup.md) — registering tenant profiles with the right app permissions.
- [`troubleshooting.md`](troubleshooting.md) — `Insufficient privileges` failure mode + diagnoses.
- [`validation-runbook.md`](validation-runbook.md) — PREVIEW-mode permission test against a real tenant.
