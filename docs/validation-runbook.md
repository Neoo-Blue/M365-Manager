# Tenant validation runbook

Live PREVIEW-mode walkthrough for the M365 Manager tool against
a real tenant. Designed to catch regressions in the high-risk
flows before they touch real users. Every destructive step is
done in PREVIEW first; LIVE is only used for the read-side
diagnostics the runbook explicitly calls out.

Read end-to-end first, then execute. If anything diverges from
the expected output, **stop** and capture the audit log line
under `%LOCALAPPDATA%\M365Manager\audit\session-*.log` before
continuing.

## 0. Prerequisites

- **PowerShell 7+** (the tool runs on 5.1 too, but the smoke
  notes were captured on 7.6 and a couple of bug fixes are
  specific to 7's stricter overload resolution; running 7
  matches what `tests/` exercises).
- **Microsoft Graph SDK + Exchange Online + SharePoint Online
  modules**. The launcher auto-installs missing ones, but if
  your machine has Constrained Language Mode or
  AllSigned execution policy the auto-install will fail; pre-
  install with:

  ```powershell
  Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Users.Actions, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
  ```

- **`gh` CLI** authenticated against the GitHub repo (only
  needed if you'll be filing follow-up issues for findings).
- **Connecting account** must have one of:
  - Global Administrator (full coverage of every flow below).
  - User Administrator + License Administrator + Exchange
    Administrator + SharePoint Administrator + Teams
    Administrator + Compliance Administrator (granular; matches
    the principle-of-least-privilege story in the README).
  - **Read-only validators** can use Global Reader + Security
    Reader for the diagnostic sections; the offboard / MFA /
    license sections need write scopes even in PREVIEW because
    Graph requires the scope to enumerate target objects.

- About 60 minutes of uninterrupted time. Most of the wall-clock
  is waiting for Graph reads (large tenants) or for Pester to
  pull modules on first run.

## 1. Bootstrap

```cmd
:: From a regular cmd / PowerShell prompt:
cd path\to\m365-manager
Launch.bat
```

or directly:

```powershell
$env:MSAL_BROKER_ENABLED = "0"
pwsh -ExecutionPolicy Bypass -NoProfile -File .\Main.ps1
```

Expected first screen:
1. Banner.
2. Module check (auto-install if anything missing).
3. **Operating mode picker** — choose `PREVIEW` for this entire
   runbook except where a step explicitly calls out LIVE.
4. **Tenant selection**:
   - First run: pick "My own organization" or "A customer
     tenant (GDAP partner access)" and connect. If you have
     Phase 6 tenant profiles already, you'll see the registered-
     tenants list instead — pick the one you want to validate.
   - Tip: register the current tenant as a profile during the
     first-run prompt; the runbook below references
     `Switch-Tenant` which needs a profile.
5. Tenant + Graph/EXO/SCC connection status bar appears, with
   the `[ PREVIEW MODE ]` banner above the main menu in yellow.

If the colored banner is **red**, you're in LIVE — go back via
"Switch Tenant" → choose PREVIEW.

### Quick sanity checks (LIVE — read-only)

From the menu, switch the mode banner to LIVE for the diagnostic
section only (these are read-only):

```powershell
# In a separate pwsh window, after Main.ps1 has finished its
# connect dance, dot-source the tool into that window:
. "<repo>\UI.ps1"
. "<repo>\Auth.ps1"
. "<repo>\Audit.ps1"
. "<repo>\Preview.ps1"
. "<repo>\Notifications.ps1"
. "<repo>\Scheduler.ps1"
. "<repo>\BreakGlass.ps1"

# Three sanity tests:
Test-NotificationChannels   # If you've configured email/webhooks,
                            # sends a low-severity test to each.
                            # PASS: returns @{ Email=$true; Teams=$true }
                            # FAIL: returns false for the channel that
                            #       didn't deliver. Check Notifications.ps1
                            #       and re-run /privacy in the AI assistant.

Test-ScheduledHealthCheck   # Stub call -- verifies the registered
                            # ScheduledTask credential decrypts.
                            # PASS: prints "Scheduled health check OK"
                            # FAIL: prints the DPAPI decryption error;
                            #       likely cred was encrypted under a
                            #       different user. Re-run
                            #       Register-ScheduledHealthCheck.

Test-BreakGlassPosture      # Posture predicates on the registered
                            # break-glass accounts.
                            # PASS: green dot per account
                            # FAIL: yellow/red dot with reason
                            #       (e.g. "password age 187 days >
                            #       threshold 90").
```

If any of the three returns a non-OK shape, capture the output
and address before continuing — the destructive flows below
depend on the notifications and audit paths working.

Switch the mode banner back to PREVIEW for the rest of the runbook.

## 2. Smoke tests (PREVIEW mode)

For each smoke test below: pick the menu option, run it on a
**test account** in your tenant. The audit log writes
`mode=PREVIEW` and `event=PREVIEW` for every line — verify in the
audit viewer (option 14 → "Open audit log viewer") at the end of
each test that the expected entries appear.

### 2a. Single-user offboard end-to-end

Highest-risk single workflow in the tool — the 12-step canonical
offboarding orchestration.

Pre-conditions: pick a test user (a real but disposable account)
who has:
- At least one assigned license.
- Membership in 1+ security groups and 1+ distribution lists.
- A mailbox (so the Shared-mailbox conversion has something to
  do).
- A OneDrive (so the handoff has something to transfer).

Menu: **2. Offboard User** → enter UPN → answer the prompts
(give a "Reason", set "ForwardTo" to a test admin, choose
"ConvertToShared = yes", "HandoffOneDriveTo = a test admin",
"RemoveFromAllGroups = yes").

Expected output (PREVIEW):
1. "Step 1/12: Revoke sign-in sessions" — `[would run]
   Revoke-MgUserSignInSession -UserId <id>`
2. "Step 2/12: Block sign-in" — `[would run] Set-MgUser
   -AccountEnabled $false`
3. "Step 3/12: Set OOO auto-reply" — `[would run] Set-
   MailboxAutoReplyConfiguration ...`
4. "Step 4/12: Set forwarding" — `[would run] Set-Mailbox
   -ForwardingSmtpAddress smtp:admin@..."`
5. "Step 5/12: Convert mailbox to Shared" — `[would run] Set-
   Mailbox -Type Shared`
6. "Step 6/12: Remove all licenses" — `[would run] Set-
   MgUserLicense -RemoveLicenses @(...)` (one per assigned SKU)
7. "Step 7/12: OneDrive handoff" — `[would run] Add-SPOSiteOwner
   ...` (multiple, one per site)
8. "Step 8/12: Remove from security groups" — `[would run]
   Remove-MgGroupMember` per group
9. "Step 9/12: Remove from distribution lists" — `[would run]
   Remove-DistributionGroupMember` per DL
10. "Step 10/12: Teams handoff" — only emits steps if user
    owns / belongs to teams; otherwise prints "no Teams
    membership found".
11. "Step 11/12: Revoke MFA methods" — `[would run] DELETE
    /authentication/<segment>/<id>` per method
12. "Step 12/12: Audit summary" — prints the entryId of every
    preview line written.

Audit log: every step should produce one JSONL line with
`event="PREVIEW"`, `actionType` matching the step (e.g.
`"BlockSignIn"`, `"SetMailboxType"`, `"RemoveLicense"`), and a
`reverse` recipe so each one would be undoable.

**Broken looks like:**
- A step is silently skipped (no audit line) — Invoke-Action
  wrap missing on that step.
- A real cmdlet error fires (e.g. "Cannot find recipient") —
  the offboard halted rather than continuing past one
  unreachable resource. This is currently expected behavior for
  some steps; check Offboard.ps1 for the `try/catch` shape.
- An audit line has `result="success"` instead of `result="preview"`
  — the cmdlet actually executed against the tenant. Stop
  immediately and verify mode banner.

### 2b. Bulk offboard from CSV

Use `templates/bulk-offboard-sample.csv` (replace the example
UPNs with two test accounts from your tenant) or create your
own:

```csv
UserPrincipalName,ForwardTo,ConvertToShared,HandoffOneDriveTo,RemoveFromAllGroups,Reason
test01@yourdomain.com,admin@yourdomain.com,yes,admin@yourdomain.com,yes,Smoke test
test02@yourdomain.com,admin@yourdomain.com,yes,admin@yourdomain.com,yes,Smoke test
```

Menu: **13. Bulk Offboard from CSV** → point at the CSV.

Expected output:
- Validation passes ("2 row(s) validated, 0 errors").
- "Continue with offboard?" prompt — confirm.
- 12-step orchestration runs for each row in sequence.
- Result CSV written next to input (`bulk-offboard-<ts>.csv`)
  with `Status` per row in `{Success, PartialSuccess, Failed,
  Preview}`. In PREVIEW mode every row should be `Preview`.

**Broken looks like:** validation errors that don't match the
CSV shape, or a row halts the whole batch (per-row errors must
not halt — that's the explicit contract).

### 2c. MFA reset

Pick a test user with at least one MFA method registered.

Menu: **16. MFA & Authentication** → "Revoke a specific method"
→ pick the user → pick the method.

Expected output:
- `[would run] DELETE /users/<id>/authentication/<segment>/<methodId>`
- Audit line: `actionType="RevokeAuthMethod"`, `noUndoReason`
  populated ("Auth method revocation cannot be undone via API;
  the user must re-register the method.").

Try the "Revoke ALL methods" path too against the same user —
should preview one DELETE per method.

Try TAP issuance — menu "Issue Temporary Access Pass":
- `[would run] POST /users/<id>/authentication/temporaryAccessPassMethods`
- Audit line carries the requested lifetime and `isUsableOnce`.

**Broken looks like:** the menu shows zero methods for a user
who you know has methods (Graph permission missing — check the
connection prompt asked for `UserAuthenticationMethod.ReadWrite.All`).

### 2d. License remediation

Menu: **20. License & Cost...** → "License optimizer".

Expected output:
- Loads `templates/license-prices.json`.
- Surfaces three categories:
  1. Anonymized usernames (SKUs assigned to users whose UPN
     matches `^([A-Z0-9]{8}|[A-Z]{2,3}\d{3,5})@`) — likely
     stale / orphaned.
  2. License-family overlap (e.g. user has E3 + Business
     Premium).
  3. Disabled users still consuming licenses.
- Each entry has an estimated monthly savings in USD.

If you proceed with a remediation (e.g. "Remove E1 from all
overlap users"), expect PREVIEW audit lines with
`actionType="RemoveLicense"` and full `reverse` recipes.

**Broken looks like:** zero savings reported on a tenant you
know has at least one disabled licensed user (filter logic
miss), or the cost math doesn't match the price-table values
(check `templates/license-prices.json` against your tenant's
actual purchased SKUs).

## 3. AI assistant smoke (PREVIEW)

Menu: **99** (hidden) → AI assistant.

Run through the chat-mode smoke:

```
You: /tools
   ... verify 30+ tools listed grouped by category, with
   [DESTRUCTIVE] markers in red.

You: /about
   ... verify provider, model, tool support, plan mode,
   cost session totals, audit/session dirs.

You: List the top 5 stale guest users.
   ... expect a single Get-StaleGuests tool_use, [Y]es to run,
   table of 5 guests rendered.

You: /plan
You: Offboard alice@yourdomain.com (transfer OneDrive to
     bob@yourdomain.com, convert mailbox to shared, remove
     licenses, remove from all groups).
   ... expect a submit_plan tool_use, the plan render with
   6-12 steps + [DESTRUCTIVE] markers + dependency arrows,
   [A]/[S]/[E]/[R] approval prompt. Choose [R] to reject for
   this smoke.

You: /save smoke-test
You: /quit
   ... expect "[auto-saved as <id>]".

# Re-launch the tool, go back to AI assistant:
You: /list
   ... expect the saved 'smoke-test' session in the list with
   [enc] marker on Windows.
You: /load smoke-test
   ... expect history reload.
You: /export <id> ./smoke-export.json
   ... expect a JSON file written; open it and verify all real
   UPNs are tokenized as <UPN_N>.
```

**Broken looks like:**
- `/tools` lists 0 tools — `ai-tools/` directory missing from
  the deployed copy.
- A destructive tool_use is **not** marked red — `wrapInInvokeAction`
  or `destructive` field misclassified in the catalog JSON.
- `/save` then `/quit` produces no "auto-saved" message — DPAPI
  encryption failed; check `<stateDir>/chat-sessions/` for
  `.session` files.
- `/export` produces a file containing real UPNs — the redaction
  pass via `Convert-ToSafePayload` failed silently; this is the
  bug fixed in PR #9 — confirm `main` includes that commit.

## 4. Critical-but-untested paths to watch

From `docs/pre-merge-review.md` §2, ranked by blast radius.
Trace through each during your live PREVIEW run:

1. **`Offboard.ps1` end-to-end** — already covered by §2a above.
   Look for: step ordering matches doc; PREVIEW doesn't halt on
   a missing resource (e.g. user has no OneDrive); audit lines
   are written in execution order.
2. **`Onboard.ps1` (single + bulk)** — covered by templates/bulk-
   onboard-sample.csv. Smoke this too with PREVIEW + a single
   row. Look for: SKU resolution matches a real subscribed SKU;
   bad UPN format halts that row only.
3. **`Auth.ps1 Select-TenantMode` + `Reset-AllSessions`** —
   exercised by switching tenant mid-session. Look for: audit
   log filename changes after switch (PR #7 fix), no stale Graph
   token from prior tenant; `Get-MgContext` reflects the new
   tenant id.
4. **`TenantSwitch.ps1 Switch-Tenant`** — exercised by
   registering 2 tenants and `Switch-Tenant -Name <other>`.
   Look for: app-only reconnect lands without an interactive
   browser prompt (cert thumbprint case); banner color rotates;
   audit log gets the new tenant slug in its filename.
5. **`MSPReports.ps1 Invoke-AcrossTenants`** — run
   `Get-CrossTenantMFAGaps` if you have 2+ registered tenants.
   Look for: per-tenant section in the output; one tenant's
   error doesn't abort the next; restored tenant context
   matches the originating context.
6. **`AIToolDispatch.ps1 Invoke-AIToolImpl` default branch** —
   PR #6 fix: destructive SDK cmdlets now wrap. Validation:
   ask the AI to do something that resolves to the default
   branch (e.g. `Update-MgUser`) and confirm the audit log has
   an `EXEC` line with `actionType="AI:Update-MgUser"`,
   `result="preview"`, `noUndoReason` populated.
7. **`Undo.ps1` undo handler dispatch** — covered by §2a's
   preview entries: pick one with `reverse.type` non-null and
   run `Invoke-Undo -EntryId <id>` in LIVE mode against a test
   resource. (LIVE because Undo is by definition a real
   reversal; if you want PREVIEW, manually inspect the handler
   scriptblock instead.)
8. **`Notifications.ps1 Send-Email`** — covered by §1 sanity
   check (Test-NotificationChannels).
9. **`UnifiedAuditLog.ps1 Search-UAL`** — run
   `Search-UAL -From (Get-Date).AddDays(-1) -To (Get-Date)
   -Operations @('MailItemsAccessed')`. Look for: returns rows
   if your tenant has UAL enabled; clean error if it doesn't
   ("UAL not enabled for this tenant").
10. **`AIPlanner.ps1 Invoke-AIPlan failureMode='ask' revise
    loop`** — craft a plan with one step that will fail in
    PREVIEW (e.g. point at a non-existent UPN), set
    `failureMode='ask'` in the plan JSON via `/plan` →
    `[E]dit`. Choose `[R]evise` at the failure prompt. Verify
    the AI receives the partial trace and submits a revised
    plan.

## 5. Exit criteria checklist

You're done when ALL of these are true:

- [ ] All Pester suites pass on the validator machine:
  `pwsh -c "Invoke-Pester ./tests/"` → 181 passed / 0 failed.
- [ ] §1 sanity checks all green (Test-NotificationChannels,
  Test-ScheduledHealthCheck, Test-BreakGlassPosture).
- [ ] §2a single-user offboard preview produces 12 audit lines
  in correct order, each with `mode=PREVIEW`, `result=preview`,
  and (for reversible ops) a non-null `reverse` recipe.
- [ ] §2b bulk offboard produces a result CSV with `Status=Preview`
  on every row.
- [ ] §2c MFA flows produce `actionType=RevokeAuthMethod` audit
  lines with `noUndoReason` populated.
- [ ] §2d License optimizer surfaces at least one savings
  opportunity OR explicitly says "no savings opportunities
  found" (i.e. doesn't silently return zero rows on a tenant
  that has overlap).
- [ ] §3 AI smoke flows: tools listed, plan mode triggers,
  session save / load / export round-trips, export contains
  zero real UPNs.
- [ ] §4 ten paths: at least the top 5 (Offboard, Onboard,
  Auth tenant switch, Switch-Tenant via profile, MSP reports)
  exercised against your tenant with no unexpected errors.
- [ ] No audit lines with `result=success` while operating in
  PREVIEW — that's a sign the wrap was missed and a real call
  hit the tenant.
- [ ] Audit log filename includes the tenant-name slug after
  any `Switch-Tenant` or `Select-TenantMode`.

If anything is unchecked, file the finding in
`docs/pre-merge-review.md` or open an issue.

## 6. Cleanup

- Drop the `/save smoke-test` session: `/delete smoke-test`.
- Remove the test users you created (manually).
- Remove `<stateDir>/audit/session-*.log` files older than the
  smoke run if you don't want them in your real audit history.
- If you registered test tenant profiles for the runbook, run
  `Remove-Tenant -Name <name>` and confirm the corresponding
  `<stateDir>/secrets/tenant-<name>.dat` was deleted.

## See also

- `docs/audit-format.md` — JSONL field reference.
- `docs/offboard-flow.md` — the canonical 12-step flow + diagram.
- `docs/multi-tenant.md` — tenant profile / switch / audit semantics.
- `docs/ai-tools.md` — catalog schema + dispatch contract.
- `docs/ai-planning.md` — submit_plan + approval UX.
- `docs/pre-merge-review.md` — the unfixed deferred items
  (chat-sessions / break-glass / scheduler not yet tenant-scoped,
  10 of 13 override keys not yet wired through Get-EffectiveConfig).
