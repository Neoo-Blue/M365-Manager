# Troubleshooting

Operator-grade error-message → likely cause → fix index. Indexed by the error text or symptom the operator actually sees. Where possible, the entries cite the module + function so deeper investigation is possible.

## How to use this doc

Search the page (Ctrl+F) for the error text you got. If you find an exact match, follow the fix. If you don't, search for keywords — most entries are written so common substrings hit.

If nothing matches, capture the full error + the audit log line for the failing operation + open an issue. The audit log line contains the `entryId` + `actionType` + `target` + raw error from Graph / EXO / SPO, which is usually enough for diagnosis.

---

## Launch + connect

### `Cannot find module Microsoft.Graph.Authentication`

The required Graph SDK isn't installed.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
```

After install, restart pwsh and try again. If it still fails, you may have an `AllSigned` execution policy that's blocking unsigned modules — see [`deployment.md`](deployment.md).

### `Module conflict: Microsoft.Graph.Identity vs Exchange Online`

The Microsoft Graph SDK and Exchange Online Management module both ship newer-than-each-other versions of MSAL DLLs. The tool sets `$env:MSAL_BROKER_ENABLED = "0"` to mediate, but only if you launch via `Launch.bat`.

Always launch via `Launch.bat`. If you're driving the tool from another script:

```powershell
$env:MSAL_BROKER_ENABLED = "0"
pwsh -ExecutionPolicy Bypass -NoProfile -File .\Main.ps1
```

### `AADSTS50034: The user account does not exist in tenant`

You signed in with an account that doesn't have access to the tenant you picked. Quit (`Q`), relaunch, pick the right tenant at the picker. For GDAP partner access, make sure your account has been delegated to the customer.

### `AADSTS500113: No reply address is registered for the application`

The connecting client app's redirect URI is misconfigured. This usually means you're using a custom Entra app registration (not the default Microsoft Graph PowerShell). Add `http://localhost` to the app's redirect URIs.

### `Connect-MgGraph: User canceled authentication`

You closed the browser before the OAuth flow completed. Relaunch and complete the consent dialog. First-time sign-in requires consenting to every scope the tool requests.

### Browser opens to OAuth but tool says "no consent given"

You clicked through the consent prompt without granting. Some scopes are required for the tool to function at all (e.g. `User.Read.All`, `Directory.Read.All`). Re-launch and accept the full set, or use a more permissive account.

---

## Graph / Exchange errors during operations

### `Insufficient privileges to complete the operation`

The connecting account lacks the required Graph scope or Entra role for the operation. See [`permissions.md`](permissions.md) for the role matrix per feature. Common cases:

- License operations need `License Administrator` + `Organization.Read.All` Graph scope.
- MFA reset needs `Authentication Administrator` + `UserAuthenticationMethod.ReadWrite.All`.
- Group / DL membership changes need `Group Administrator` or `Exchange Recipient Administrator`.

### `Resource not found: 'alice@contoso.com'`

The user doesn't exist in the current tenant. Common causes:
- Typo in the UPN.
- You're connected to the wrong tenant (check the banner color + tenant name).
- The user is in the recycle bin (`/directory/deletedItems`) — restore first if needed.

### `Authentication_MissingOrMalformed`

The current Graph token expired and re-auth failed. Run `Reset-AllSessions` then re-connect.

### `Set-MgUserLicense: License assignment cannot be applied to a user with no usage location`

The user's `UsageLocation` (ISO country code) isn't set. Fix:

```powershell
Update-MgUser -UserId alice@contoso.com -UsageLocation "US"
```

### `Set-Mailbox: The mailbox is not found`

The user has a Graph account but no Exchange mailbox provisioned. Common for:
- Brand-new users (mailbox provisioning takes 5-60 minutes).
- Users who never had a license assigned.
- Guest users (they don't get mailboxes in your tenant).

Wait for provisioning or skip the mailbox-related steps.

### `Cannot find recipient 'sales-announce@contoso.com'`

The distribution list doesn't exist (or its primary SMTP changed). Check with:

```powershell
Get-Recipient -Identity sales-announce@contoso.com
```

Update the role template to reference the correct address.

---

## SharePoint

### `Connect-SPOService: The remote server returned an error: (404)`

The SPO admin URL is wrong. Fix by setting:

```powershell
Set-SPOAdminUrl -TenantId <tenant-guid> -Url "https://contoso-admin.sharepoint.com"
```

Check the cached value at `<stateDir>\spo-tenants.json`.

### `Get-UserOutboundShares returns empty when you know there are shares`

UAL isn't enabled for the tenant, OR the user is outside the audit retention window (90d default; 1y with E5). Verify UAL is on in Compliance Center → Audit. The function reads `SharingSet` operations from UAL — if UAL has nothing, the function has nothing to return.

### `Add-PnP module not found`

Some flows need the PnP SharePoint module. Install:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

---

## Teams

### `Cannot add member to team: insufficient privileges`

The connecting account isn't a Teams Administrator and doesn't own the target team. Fix: either grant `Teams Administrator` to the account, or have an existing team owner add the member, or run as Global Admin.

### `Multiple teams match 'Sales'`

You used `-TeamName` and the name matched more than one team (contains-match). Either refine the name or pass `-TeamId`.

### `Cannot remove the only owner`

Microsoft 365 prevents removing the only owner of a team. Promote a successor first, then remove. The offboarding flow's `Invoke-TeamsOffboardTransfer` handles this automatically.

---

## MFA

### `Insufficient privileges to read this user's authentication methods`

Need `UserAuthenticationMethod.Read.All` Graph scope. Re-consent or use `Authentication Administrator` role.

### `Cannot revoke the Password method`

Passwords aren't auth methods — they're set via `passwordProfile`. Use the password reset flow instead.

### `User has no methods to revoke`

Either:
- The user has only a password registered (passwords don't appear via the auth-methods API).
- The user is a Guest (their auth lives in the home tenant).

Both are expected — verify the user is the one you intended.

### `TAP issuance failed: User is excluded from policy`

Your tenant's TAP policy excludes this user. Common case: Global Admins are excluded by default. Adjust the policy in Entra → Authentication methods → TAP, OR use a TAP-eligible account for the issue.

---

## Audit + undo

### `Could not parse line N: Unexpected character`

A non-JSON line slipped into the audit log. Usually from a crash that left partial output. The viewer skips the line and continues. Capture the file + line number for an issue.

### `Invoke-Undo: handler not registered for reverseType 'X'`

The audit entry references a reverse type that's no longer in `$script:UndoHandlers`. Check [`../reference/undo-handlers.md`](../reference/undo-handlers.md) — possibly the entry was written by an older version. Manual reversal is your only option.

### `Already reversed in undo-state.json`

The original entry was undone in a prior session. You can verify in `<stateDir>\undo-state.json`. If you legitimately need to redo the reversal, delete the entry from the sidecar — but understand you're losing the dedup guarantee.

### Audit log file is missing

Check `%LOCALAPPDATA%\M365Manager\audit\` for `session-*.log`. If the dir is empty, either:
- This is a fresh install before any tool run.
- An old version that didn't generate the per-session file naming.

The viewer's `-Path` flag accepts an explicit file if needed.

---

## AI assistant

### `Mark says: "I don't have access to do that."` when you know it should

The tool the AI proposed isn't in the catalog. Run `/tools` to verify — every callable tool must be in `ai-tools/*.json`. If the tool is missing, add it (see [`../developer/adding-an-ai-tool.md`](../developer/adding-an-ai-tool.md)).

### `Convert-ToSafePayload: 'Get-OrCreatePrivacyToken' is not recognized`

Pre-fix bug from the v1 ship — fixed in PR #9 / commit 4dd12b7. Make sure you're on `main` past that commit; pull if not.

### Chat is suspiciously slow

Several causes:
- The provider is rate-limiting. Try a different model or wait.
- The conversation history is very long; `/clear` to reset.
- A tool with a slow Graph call (Search-SignIns on a wide window, Get-MgUser without `-Property` filter). Narrow the request.

### `/save` then `/quit` produces no "auto-saved" message

DPAPI encryption failed. Look for the WARN line in the audit log. Common cause: the operator account doesn't have a valid DPAPI scope (very rare; usually only on Active Directory accounts with broken roaming profile).

### `/export` produces a file containing real UPNs

Redaction failed. Pre-fix bug (closure scoping); fixed in PR #9. Update to `main` past that commit.

### Cost shows zero for every call

Pricing JSON missing. Check `templates/ai-prices.json` — if the provider/model combo isn't listed, cost defaults to zero. Add the entry.

---

## Incident response (Phase 7)

### `Step 4 (RevokeAuthMethods) failed: Remove-AllAuthMethods not loaded`

`MFAManager.ps1` didn't load. Verify `Main.ps1`'s module list includes it. If you're running from a slimmed-down install, re-clone to get every module.

### `Snapshot mfaMethods is empty but user has methods`

`Get-UserAuthMethods` requires `UserAuthenticationMethod.Read.All` Graph scope. The snapshot does its best — the capturedErrors array surfaces what failed. Re-consent the app or use a more permissive account.

### `Audit24h returns no UAL rows`

UAL not enabled for the tenant or operation is outside the 30-min ingestion window. Wait + re-run, or verify UAL in Compliance Center.

### Tabletop run fails with "TabletopUPN not configured"

Set `IncidentResponse.TabletopUPN` in `ai_config.json`, OR pass `-TabletopUPN <upn>` to `Invoke-IncidentTabletop`.

### `Detect-ImpossibleTravel` flags a known business trip

Add geo-anchors for the new country if missing (`Get-CountryDistanceKm` table) OR raise `IncidentResponse.ImpossibleTravelMaxKmPerHour` in the per-tenant override.

---

## Multi-tenant

### `Switch-Tenant: app-only auth failed`

The cert thumbprint or client secret stored in the tenant manifest is wrong or expired. Re-run `Register-Tenant` with current credentials.

### `Audit log file didn't change after Switch-Tenant`

`Reset-AuditLogPath` should fire automatically — check that you're running `feature/multi-tenant` or later (the fix landed in PR #7 / commit f9d0c5e).

### Tenant override doesn't take effect

Only 3 of 13 catalog'd override keys are wired through `Get-EffectiveConfig` today. See [`../reference/config-keys.md`](../reference/config-keys.md) for the wired-vs-not status. Mechanical conversion is tracked as a follow-up.

---

## Scheduled checks

### Task Scheduler entry doesn't run

Verify the credential decrypts:

```powershell
Test-ScheduledHealthCheck
# Should print: "Scheduled health check OK"
# Fails if the DPAPI cred was encrypted under a different user.
```

Re-register if needed: `Register-SchedulerCredential`.

### Task runs but writes no result

Check `<stateDir>\health-results\` for output. If empty, check Task Scheduler's history for the task's exit code; common causes are connect failure (account in the cred doesn't have access) and module-load failure (missing modules on the host).

---

## File system + state

### `<stateDir>` location is wrong

The tool resolves `<stateDir>` to:
- `%LOCALAPPDATA%\M365Manager\state` on Windows.
- `~/.m365manager/state` on POSIX.

If your environment lacks `%LOCALAPPDATA%`, set it before launching. If you want a non-default path, set `$env:M365MGR_STATEDIR` (planned; not currently implemented).

### Can't read DPAPI'd file after machine migration

DPAPI is bound to the user account that wrote the file. Migrate by re-running the relevant setup function (`Register-Tenant`, `Start-NotificationsSetup`, AI `/config`) on the new machine. The original file is unrecoverable.

---

## See also

- [`permissions.md`](permissions.md) — required Graph scopes + roles per feature.
- [`validation-runbook.md`](validation-runbook.md) — live PREVIEW smoke tests against a tenant.
- [`pre-merge-review.md`](pre-merge-review.md) — known issues + deferred items.
- Open an issue: https://github.com/Neoo-Blue/M365-Manager/issues
