# Security model

What's protected, how, and what the operator trusts vs not. M365 Manager handles credentials, tenant data, and AI-tool routing — each lane has its own threat model.

## At a glance

| Threat | Mitigation |
|---|---|
| Stolen `ai_config.json` from disk | API key DPAPI-encrypted (CurrentUser, LocalMachine) — not portable. |
| Operator workstation compromise | Same DPAPI scope — the attacker still needs the user's logon credential to decrypt. |
| Bystander shoulder-surf during chat | Reverse-PII (tokenize-then-restore) keeps temp passwords / GUIDs off-screen mid-conversation. |
| External LLM logs prompts | PII tokenization before send; restoration on response. Operator's UPN / tenant ID never reaches the provider in cleartext. |
| Audit-log leak | Path defaults to per-user dir with restrictive ACL; secret-bearing params (`-Password`, `-Token`, etc.) always scrubbed; tokenization optional via `Privacy.RedactInAuditLog`. |
| Operator typo deletes wrong user | Confirmation prompts on every mutation; PREVIEW mode default-available; every action through `Invoke-Action` with a reverse recipe. |
| AI proposes a destructive cmdlet | AST allow-list + `requiresExplicitApproval` flag + per-step prompt — even with `[A]pprove-all`, the highest-blast-radius tools force a fresh yes. |
| Stolen DPAPI'd secret manifest | Useless on a different user / machine. Required to be re-registered after migration. |

## DPAPI at rest

Three secrets ride DPAPI in the standard flow:

1. **AI API key** — `ai_config.json`'s `ApiKey` field. Encrypted on first save via `Protect-ApiKey` (`AIAssistant.ps1:77`). The plaintext form gets re-encrypted in place if you paste it manually; you don't have to do anything special.
2. **Notification webhook URLs + SMTP password** — `Protect-Secret` / `Unprotect-Secret` in `Notifications.ps1`. URLs are sensitive (anyone with the Teams webhook URL can post to that channel) so they're encrypted at rest with a `DPAPI:` prefix marker.
3. **Tenant client-secret + cert thumbprint manifests** — `<stateDir>\secrets\tenant-<name>.dat` per Phase 6 tenant profile. Created by `Register-Tenant`; decrypted at switch time by `Get-TenantCredentialManifest`.

DPAPI bindings:

- **Scope**: `CurrentUser` (not `LocalMachine`). The encrypted blob can only be decrypted by the same Windows user account that wrote it.
- **Portability**: not portable across machines. Roaming-profile users carry the key with the user, but moving the file by USB stick to another machine produces an unreadable blob.
- **POSIX fallback**: writes a `Base64Plain:` form with a warning. On non-Windows hosts there's no equivalent OS-managed key store, so the tool falls back to base64 + chmod 600. Acceptable for development; not for production-secrets workloads.

## PII redaction layer

`Convert-ToSafePayload` (`AIAssistant.ps1:608`) tokenizes outbound text before it hits an external LLM:

- **Always tokenized** (regardless of provider): JWTs, `sk-...` / `sk-ant-...` API keys, 40-hex cert thumbprints. These are hard-coded — there's no config knob to disable.
- **Tokenized for external providers when `ExternalRedaction=Enabled`** (default): UPNs (`<UPN_1>`), GUIDs (`<GUID_3>`), tenant IDs (`<TENANT>`), display names captured from cmdlet idioms.
- **Tokenized in audit log when `RedactInAuditLog=Enabled`** (off by default — forensics-friendly default keeps raw values).

The map is per-session, hashtable-backed, with stable token assignment (the same UPN gets the same `<UPN_1>` across the whole session). `Restore-FromSafePayload` reverses the substitution before displaying the AI's response to the operator or dispatching its tool calls.

A subtle bug in this path (closure-scope failure) shipped briefly during the v1 build and was fixed in PR #9 — the `MatchEvaluator` scriptblocks now capture `Get-OrCreatePrivacyToken` via `${function:...}` so they survive being invoked from child scopes. See [`../operations/pre-merge-review.md`](../operations/pre-merge-review.md) finding #4.

## AST allow-list

`Test-AICommandAllowed` (`AIAssistant.ps1:819`) parses every AI-proposed command via the PowerShell language parser and checks the resolved command against `$script:AICmdAllowList`. Cmdlets not on the allow-list are rejected without prompting. The allow-list is a hard-coded set of patterns (`Get-Mg*`, `Set-Mailbox`, `New-DistributionGroup`, etc.) that span the modules the AI is expected to drive.

This is defense-in-depth for the legacy regex `RUN:` extractor path. The native tool-calling path (Phase 5) is gated by the catalog (`ai-tools/*.json`) which is an even tighter allow-list — only catalog'd tools can be invoked, and the dispatcher resolves the AI's tool name against the catalog before running anything.

## Audit log

Every mutation through `Invoke-Action` writes one JSONL line. Schema in [`../reference/audit-format.md`](../reference/audit-format.md). Highlights:

- **`entryId`** — UUID correlating PROPOSE / EXEC / OK / ERROR / UNDO entries for one logical operation.
- **`mode`** — `LIVE` or `PREVIEW`.
- **`actionType`** — stable filter key (`AssignLicense`, `BlockSignIn`, `Incident:RevokeSessions`, etc.).
- **`target`** — structured hashtable of operands.
- **`reverse`** — recipe for undo dispatch, or `null` when irreversible.
- **`noUndoReason`** — populated on irreversible operations explaining why.
- **`tenant`** — structured `{name, id, domain, mode}` block (Phase 6); legacy entries have a string tenant.
- **`session`** — OS PID of the M365 Manager process that wrote the line.

Log file lives at `%LOCALAPPDATA%\M365Manager\audit\session-<ts>-<pid>-<tenant>.log`. The directory inherits NTFS ACLs from `%LOCALAPPDATA%` on Windows (user-only by default) and is created with `chmod 700` on POSIX.

## Undo system

Every reversible mutation writes a `reverse` recipe specifying:

```jsonc
"reverse": {
  "type":        "RemoveFromGroup",
  "description": "Remove user 4f3a-... from security group 'SG-Sales'",
  "target":      { "userId": "4f3a-...", "groupId": "g-aaaa" }
}
```

The `Invoke-Undo -EntryId X` dispatcher looks up `$script:UndoHandlers[type]` and runs the handler against the target. The dispatch table is in `Undo.ps1`; today it covers 24 reverse types covering license + group + DL + mailbox permission + calendar + OOO + forwarding + sign-in block + OneDrive access + site owner + team membership / ownership. New reversible actions need a matching handler entry.

Irreversible actions (`Remove-MgUser`, compliance purge, `Revoke-MgUserSignInSession`, MFA method removal, etc.) populate `noUndoReason` explicitly so the operator knows what they're committing to.

## Confirmation prompts

Mutation cmdlets at the menu level all go through `Confirm-Action` (`UI.ps1:118`) before calling `Invoke-Action`. Confirmation defaults vary by feature; the audit-and-undo guide documents which paths are auto-confirmed vs explicit.

In NonInteractive mode (`-NonInteractive` flag for scheduled runs):

- `Confirm-Action` returns `$false` (declines).
- `Read-UserInput` returns empty.
- `Show-Menu` returns -1.

So a scheduled run never silently auto-approves anything. The Phase 7 incident playbook's quarantine step extends this principle: even in NonInteractive mode, step 11 (compliance purge) short-circuits with a `manual step required` audit entry rather than proceeding.

## AI explicit-approval gating

Phase 7 added a catalog flag `requiresExplicitApproval` (currently set on `Invoke-CompromisedAccountResponse` only). When the AI proposes a tool with this flag:

- The tool-use confirmation prompt prints a yellow `EXPLICIT APPROVAL REQUIRED` banner.
- The `[A]pprove all` shortcut is removed from the prompt — only `Y / N / Q`.
- Any prior approve-all decision in the session is ignored for this single tool.
- The plan approval flow (Phase 5) detects the flag in any plan step and forcibly downgrades `approveAll` to `stepByStep`.

So even an operator who's been hitting `A` to batch-approve safe tool calls gets a fresh stop sign when the highest-blast-radius operation comes up.

## What's still trusted

The threat model excludes:

- **An attacker with the operator's logon credential.** DPAPI decrypts under that account; anyone who can log in as the operator can read every secret the tool stores.
- **Compromised Graph SDK / EXO PowerShell modules.** The tool calls into those libraries; if they're tampered with, every mutation is suspect. Defense: sign your modules + use AllSigned execution policy on production workstations.
- **An attacker with elevated permission to the audit log directory.** A determined attacker can delete or modify session log files. Mitigation: ship audit logs to a SIEM in real time (out of scope for this tool today; see follow-up at [`../operations/pre-merge-review.md`](../operations/pre-merge-review.md)).
- **Provider-side prompt logging.** Even with redaction, conversation metadata (length, frequency, tool-call patterns) is visible to the AI provider. Default `Anthropic` is "zero retention" per the enterprise terms; `OpenAI` retains 30 days for abuse monitoring unless you have a ZDR contract. See [`../guides/ai-tools-overview.md`](../guides/ai-tools-overview.md) for the trust matrix per provider.

## See also

- [`../operations/permissions.md`](../operations/permissions.md) — Graph scope + role matrix.
- [`../reference/audit-format.md`](../reference/audit-format.md) — full audit JSONL field reference.
- [`../guides/audit-and-undo.md`](../guides/audit-and-undo.md) — operator-facing audit + undo walkthrough.
- [`../operations/pre-merge-review.md`](../operations/pre-merge-review.md) — known limitations + deferred items.
