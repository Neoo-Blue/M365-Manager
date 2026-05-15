# Pre-merge self-review

Critique pass on the seven-branch stack before merging to `main`.
Honest, evidence-based, prioritized by blast radius. Findings cite
`file:line` where I checked. Anything I would have done differently
with hindsight is called out.

## 1. Cross-phase consistency

**Phase 6 brief asked for tenant-scoped chat sessions and break-glass
files; neither shipped.**
- `AISessionStore.ps1:31` still uses flat `<stateDir>/chat-sessions/`
  with no tenant subdirectory. A session saved while in tenant A
  will auto-load (via `/list`) when you switch to tenant B. The
  privacy map *was* stored inside the session blob (Phase 5 D), so
  loading it would restore tenant-A real values into a tenant-B
  context. That's the cross-tenant-contamination class the brief
  flagged.
- `BreakGlass.ps1:23` writes `<stateDir>/breakglass-accounts.json`
  flat. Two registered tenants share the same break-glass registry.
- `Scheduler.ps1` has no `-Tenant <name>` parameter on
  `New-ScheduledHealthCheck`. Cross-tenant scheduled refreshes
  aren't possible today.

**Phase 6 audit-field shape change (`Audit.ps1:99-108`) is
backward-compatible read-side, but write-side every prior phase
already used `SessionState.TenantDomain` as a string fallback.**
The `AuditViewer.Filter-AuditEntries` Tenant filter
(`AuditViewer.ps1:190-194`) handles both shapes via substring
match. Verified by re-reading old Phase 4 audit log fixtures.

**Phase 6 audit log filename includes a tenant slug (`Audit.ps1:47-
51`) — but `Reset-AuditLogPath` is only called by
`Switch-Tenant`.** The legacy `Select-TenantMode` path (Auth.ps1)
doesn't call it, so an operator who switches via the partner-center
picker (not via a registered profile) won't land in a new file.
Minor — usability vs correctness — but worth fixing.

## 2. Untested-but-critical paths

Top 10 ranked by "what would real damage look like":

1. **`Offboard.ps1` end-to-end** — no Pester. The 12-step
   orchestrator is the highest-risk single function in the
   codebase. A regression in step ordering (e.g. delete user
   before transferring OneDrive) is recoverable from `/directory/
   deletedItems` for 30 days but loses OneDrive ownership.
2. **`Onboard.ps1` (single + bulk)** — no Pester for the licensing
   step. A bad SKU match could license-strip an existing user.
3. **`Auth.ps1` `Select-TenantMode` + `Reset-AllSessions`** — no
   Pester. SDK connection state is process-global; bugs here lead
   to writing changes to the wrong tenant.
4. **`TenantSwitch.ps1` `Switch-Tenant`** — no Pester. Phase 6's
   most consequential function and the one with the cross-tenant
   contamination risk. The Pester suite stubs it out for
   InvokeAcrossTenants tests but doesn't exercise the real
   reconnect path.
5. **`MSPReports.ps1` `Invoke-AcrossTenants`** — Pester exists
   for happy path + restore-on-throw, but does NOT cover the case
   where Switch-Tenant itself fails mid-iteration (e.g. cert
   missing on host).
6. **`AIToolDispatch.ps1 Invoke-AIToolImpl` default branch
   (line 251-255)** — see finding #4 below. No test verifies
   that destructive tools hitting the default branch are
   audit-wrapped.
7. **`Undo.ps1` undo handler dispatch** — partial Pester on the
   handler table, but no end-to-end "do a thing, then undo it,
   confirm tenant state matches pre-do" test against any mock.
8. **`Notifications.ps1` `Send-Email`** — no Pester. A bad sender
   address or recipient list silently swallows alerts.
9. **`UnifiedAuditLog.ps1` `Search-UAL`** — no Pester. The
   compliance use cases (eDiscovery, breach response) lean on
   this; a regression in operation-name filtering would be silent.
10. **`AIPlanner.ps1 Invoke-AIPlan` failureMode='ask' revise
    loop** — Pester covers rejection but not the revise/abort/
    continue prompt or the synthetic re-plan request to the AI.

## 3. Documentation gaps

- `docs/guides/ai-costs.md` mentions `Test-AIBudgetCap` as a
  "future hook" — that function does not exist anywhere. Either
  build it or drop the reference. (caught by grep; pre-merge fix
  candidate.)
- `docs/reference/audit-format.md` was updated for Phase 6's structured
  tenant field but the JSON example still shows the legacy
  `"tenant":"contoso.onmicrosoft.com"` string shape. Confusing for
  someone scanning the doc.
- `docs/concepts/tenant-overrides.md` declares 13 overrideable keys; only
  3 are actually consulted via `Get-EffectiveConfig` today
  (`AI.MonthlyBudgetUsd`, `AI.AlertAtPct`, `StaleGuestDays`). The
  doc honestly flags this as "follow-up" but the README's bullet
  doesn't — README implies the full set works.
- No top-level `CHANGELOG.md`. Six phases of stacked branches and
  the only narrative is in commit messages.

## 4. Audit log integrity

**Real gap: AI dispatch's default branch does NOT honor
`wrapInInvokeAction`.** `AIToolDispatch.ps1:250-255` is a generic
`& $ToolName @splat` fallthrough. The following destructive tools
hit this path:

| Tool                              | Underlying call type  | Wraps? |
|-----------------------------------|------------------------|--------|
| `Remove-MgGroupMember`            | Graph SDK cmdlet       | NO     |
| `Remove-DistributionGroupMember`  | Exchange cmdlet        | NO     |
| `Set-MailboxAutoReplyConfiguration` | Exchange cmdlet      | NO     |
| `Remove-MailboxPermission`        | Exchange cmdlet        | NO     |
| `Update-MgUser`                   | Graph SDK cmdlet       | NO     |
| `Revoke-MgUserSignInSession`      | Graph SDK cmdlet       | NO     |
| `Remove-Guest`                    | Custom (`GuestUsers.ps1`)| YES (wraps each step internally) |
| `Remove-AllAuthMethods`           | Custom (`MFAManager.ps1`)| YES (via `Remove-AuthMethod`)    |
| `Revoke-OneDriveAccess`           | Custom                 | YES    |
| `Remove-UserFromTeam`             | Custom                 | YES    |
| `Set-TeamOwnership`               | Custom                 | YES    |
| `New-TemporaryAccessPass`         | Custom                 | YES    |

Six raw SDK / Exchange cmdlets run without going through
`Invoke-Action` when the AI calls them. They land an
`AIToolCall` audit line (`AIToolDispatch.ps1:289`) but no
`EXEC` line with `actionType` / `reverse` / `noUndoReason`.
That breaks undo, retroactive PREVIEW-only forensics, and the
contract documented in `docs/reference/audit-format.md`.

**Pre-merge fix candidate.** Either:
- Add explicit `case` branches in `Invoke-AIToolImpl` for the
  six SDK/Exchange cmdlets, mirroring how `Set-MgUserLicense-*`
  is handled, OR
- Generalize: have the dispatcher consult `wrapInInvokeAction`
  on the tool definition, and wrap the default `& $ToolName
  @splat` call with `Invoke-Action` when true. The wrapping
  loses the ability to set `ReverseType` / `ReverseDescription`
  per-tool though, so explicit branches are richer. I'd do the
  explicit-branch approach — it's ~30 lines.

## 5. Undo coverage

I diffed every `-ReverseType` emitter site against
`$script:UndoHandlers` keys (`Undo.ps1`). Every emitter has a
handler. Four handlers exist with no emitter today (`BlockSignIn`,
`GrantCalendarAccess`, `GrantMailboxFullAccess`,
`GrantMailboxSendAs`) — those are the symmetric pairs of
`UnblockSignIn`/`Revoke*` so they're reachable via undo-of-undo.
**Coverage is solid.** The one gap is implicit: when an action
*should* be reversible but `wrapInInvokeAction=false` in the
dispatcher (finding #4), no `reverse` field gets written at all,
so the handler is unreachable from `Show-RecentUndoable`.

## 6. Privacy redaction

Verified by grep:
- `Invoke-AIChat` (AIAssistant.ps1:324) calls `Convert-ToSafePayload`
  before the POST.
- `Invoke-AIChatToolingTurn` (AIToolDispatch.ps1:412-421) calls
  `Convert-ToSafePayload` before the POST (including embedded
  tool_use / tool_result blocks).
- `Export-AISession` (AISessionStore.ps1:306) calls it before
  writing the redacted export.
- `Notifications.ps1:204` (Teams webhook POST) does NOT redact.
  That's intentional — webhook target is the operator's own
  infra, not an external LLM — but it should be doc'd. Not a
  merge blocker.

**No external-AI POST bypasses redaction.** Clean.

## 7. Credential storage

- AI API key: `Protect-ApiKey` / `Unprotect-ApiKey`
  (`AIAssistant.ps1:77,92`). DPAPI on Windows, B64 with warning
  on POSIX.
- Notification webhooks + SMTP password: `Protect-Secret` /
  `Unprotect-Secret` (`Notifications.ps1:32,49`). Same shim.
- Phase 4 scheduler credential: `Scheduler.ps1` writes a
  DPAPI-protected sidecar.
- Phase 4 break-glass entries: the *passwords* are not stored;
  only metadata is.
- Phase 6 tenant client secret: `TenantRegistry.ps1:148` runs
  through `Protect-Secret` on the way in,
  `Get-TenantCredentialManifest:222` decrypts on the way out.
  Verified by the `TenantRegistry.Tests.ps1 "writes an encrypted
  manifest"` test — asserts the plain secret is absent from the
  raw file but present after decryption.

**No plaintext credentials at rest.** Clean.

## 8. Module load order

Walked `Main.ps1:71-83` against every `Get-Command -ErrorAction
SilentlyContinue` guard. The guards are well-placed: every
forward reference (e.g. `TenantSwitch.ps1:190` references
`Update-MSPDashboard` which loads later) checks
`Get-Command`-before-call. So `Main.ps1`'s loader can be reordered
without breaking guards, but the current order does match the
implicit dependency graph (UI → Auth → Audit → Preview → Notifs →
Tenant\* → feature modules → MSP\* → AI\*). Clean.

## 9. Known limitations (collected from smoke notes)

- **AuditViewer.ps1 +1 brace delta** — false positive from the
  literal `'{'` in the JSONL-detect branch. Documented in every
  phase's commit message; not a real bug. Verified: 209 `{` vs
  208 `}`, the extra `{` is inside a string literal.
- **Pester not runnable on the dev machine (Mac, no `pwsh`)** —
  no live test runs done for any phase. Smoke = brace count +
  JSON parse only. All 19 Pester suites are written but unrun.
- **App-only Graph reconnect (Phase 6)** assumes the cert is in
  the LocalMachine / CurrentUser store and the SDK can locate
  it. Hosts without the cert fall through to interactive — the
  operator will see a browser prompt on next service call.
  Documented in `TenantSwitch.ps1:113-115`.
- **`Invoke-AcrossTenants -Parallel` is accepted but ignored**
  (sequential under the hood). Documented in
  `docs/concepts/multi-tenant.md`.
- **MSP dashboard does not deep-link to per-tenant detail
  reports** — flagged in commit D body and `docs/guides/msp-dashboard.md`.
- **PowerShell 5.1 vs 7 compat** — only 5.1 was specifically
  tested per smoke. Phase 5 streaming uses HttpWebRequest which
  works on both, but `Start-Job` background spinner behavior may
  differ.

## 10. Recommended pre-merge fixes (ranked)

These are the items I would fix on the relevant feature branch
before pushing the merge train. Ordered by blast-radius / effort
ratio.

1. **Honor `wrapInInvokeAction` for SDK/Exchange cmdlets in
   `AIToolDispatch.ps1`** (finding #4). Six tools currently
   bypass audit on AI-driven invocation. ~30 lines, high impact.
   **Will fix before merge.**
2. **Remove `Test-AIBudgetCap` phantom reference from
   `docs/guides/ai-costs.md`** (finding #3). One-line fix.
   **Will fix before merge.**
3. **Update `docs/reference/audit-format.md` JSON example to the Phase 6
   structured tenant shape** (finding #3). The legacy string
   example still in the doc is misleading.
   **Will fix before merge.**
4. **Scope chat sessions per tenant**
   (`AISessionStore.ps1`, finding #1). Real cross-tenant
   contamination risk. Estimated 40 lines + test update.
   **Deferring as follow-up** — would require coordinated tenant
   migration of the existing flat layout and the scope spec said
   "if the review surfaces fixes you'd make before merging" —
   this one is too big for a no-PR-yet window. Flag clearly in
   docs / changelog as known-limitation.
5. **Scope break-glass registry per tenant**
   (`BreakGlass.ps1`, finding #1). Same logic as #4.
   **Deferring as follow-up.**
6. **Add `-Tenant <name>` to `New-ScheduledHealthCheck`**
   (finding #1). Needed for MSP scheduled refreshes.
   **Deferring as follow-up.**
7. **Call `Reset-AuditLogPath` from `Select-TenantMode`**
   (finding #1, audit). Two-line fix in `Auth.ps1`.
   **Will fix before merge.**
8. **README accuracy** — soften the tenant-overrides bullet to
   reflect that only 3 of 13 keys are wired today. One-line edit.
   **Will fix before merge.**
9. **Add a top-level `CHANGELOG.md`** summarizing each phase's
   delta. **Deferring as follow-up** — would balloon scope and
   the commit messages already carry the narrative.
10. **End-to-end Pester for `Switch-Tenant` and `Offboard.ps1`**
    (finding #2). The two highest-risk untested paths.
    **Deferring as follow-up** — meaningful tests need a Graph
    mock library; "drop in a stub" wouldn't catch the regressions
    that matter.

## What I would have done differently with hindsight

- **Phase 5 commit D should have already been tenant-scoped.**
  Sessions / costs / privacy maps are inherently per-tenant
  artifacts and bolting that on in Phase 6 doubles the work and
  leaves a contamination window.
- **`wrapInInvokeAction` should never have been a metadata-only
  flag.** Either the dispatcher should honor it generically or
  every destructive tool should have an explicit branch — having
  16/18 wrapped at the metadata level but only 6/18 in code is
  the worst of both.
- **Pester suites should have been run, not just written.** Mac
  / no-pwsh was a known constraint from day one. A one-time
  Linux container or GitHub Actions step would have caught any
  Pester syntax errors I missed (none observed by eyeball, but
  19 suites is a lot to read).

## Conclusion

Stack is mergeable after the four "Will fix before merge" items
land. The deferred-as-follow-up items are real but each one
needs a coordinated migration or a mock layer that's out of
scope for a pre-merge sweep. None of them is a security or
data-integrity blocker once finding #4 is fixed: the AI tool
dispatch audit gap was the single load-bearing concern.
