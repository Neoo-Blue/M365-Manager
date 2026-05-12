# M365 Administration Tool

A modular PowerShell TUI application for common Microsoft 365 admin tasks. Blue background, multi-color interface, confirmation prompts on every change, and browser-based OAuth authentication.

## Prerequisites

Install the required PowerShell modules (the tool will auto-install if missing):

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Groups -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

> **Note:** On launch, the tool automatically checks every module: installs if missing, imports into the session, and verifies key commands are available. If any dependency fails, it shows the specific error and offers to continue or exit.

## Launch

**Always use `Launch.bat`** — it sets a required environment variable (`MSAL_BROKER_ENABLED=0`) that prevents DLL conflicts between the Microsoft Graph and Exchange Online modules.

Double-click `Launch.bat` in the M365Admin folder.

If running manually, set the environment variable first:

```powershell
$env:MSAL_BROKER_ENABLED = "0"
powershell -ExecutionPolicy Bypass -NoProfile -File Main.ps1
```

## Configuration & secrets

Anything that may contain a secret (API keys, tokens) lives in `ai_config.json`, which is **gitignored** and DPAPI-encrypted on first save (per-user, per-machine — not portable).

- `ai_config.example.json` is the template, safe to commit.
- `ai_config.json` is created automatically on first run of the AI assistant (hidden option `99`). To pre-seed it manually, copy the example and run the assistant once — your plaintext key will be encrypted in place on first load.
- Build artifacts (`M365Admin_Merged.ps1`, `M365Admin.exe`) and audit logs (`audit/`) are also gitignored.

## AI privacy / PII handling

When the assistant talks to a non-local LLM (Anthropic, OpenAI, Azure OpenAI, or a remote Ollama / custom endpoint), PII in your prompts and tool output is **tokenized** before send: UPNs, emails, GUIDs, tenant IDs, JWTs, API keys, cert thumbprints, and display names captured from cmdlet arguments are each replaced with a stable opaque placeholder (`<UPN_1>`, `<GUID_3>`, `<TENANT>`, ...). The reverse map is session-scoped and dropped on `/clear` or assistant exit. The AI's response is restored before display and before any command runs, so you see real values and commands target the real objects.

**Secrets (JWT / `sk-…` / `sk-ant-…` / 40-hex thumbprints) are ALWAYS tokenized regardless of provider.** That rule is hardcoded.

Configure from the assistant chat with `/privacy`:

| Setting | Default | Notes |
|---|---|---|
| `ExternalRedaction` | `Enabled` | Full PII tokenization for non-local providers. Disable only if you have an alternate redaction layer upstream. |
| `RedactInAuditLog` | `Disabled` | When Disabled, the audit log under `%LOCALAPPDATA%\M365Manager\audit\` (or `~/.m365manager/audit/` on POSIX) records real values for forensics; secret-bearing params (`-Password`, `-Token`, etc.) are always scrubbed. Enable to also tokenize PII in the audit log itself. |
| `ExternalPayloadCapBytes` | `8192` | After redaction, outbound message content is truncated at this many bytes with an explicit `[TRUNCATED N BYTES]` marker. `0` disables the cap. Applies only to external providers. |
| `TrustedProviders` | `[]` | Lowercase provider names treated like localhost: PII is sent raw, but secrets are still scrubbed. Example: `["azure-openai"]` if your Azure OpenAI deployment lives inside the same tenant you're administering. **Only add a provider here after reviewing its data-handling terms.** |

**Provider retention defaults** (as of writing — verify against the current provider terms):

- **Anthropic** (enterprise API key) — zero retention.
- **OpenAI** — 30-day abuse-monitoring retention unless your organization has opted out via a Zero Data Retention agreement.
- **Azure OpenAI** — governed by your Azure subscription's data terms; if your deployment is in your own tenant, data does not leave that boundary.
- **Ollama / LM Studio on localhost** — never leaves the machine. The assistant does not redact for local providers by default.

The audit directory is created with mode `0700` on non-Windows and inherits user-only NTFS ACLs from `%LOCALAPPDATA%` on Windows.

## Onboarding role templates

Drop a `role-<slug>.json` into `templates/` and it shows up in the onboarding picker automatically. Five examples ship in the repo (`sales-rep`, `engineer`, `exec-assistant`, `contractor`, `default`); see `templates/README.md` for the schema and `Get-MgSubscribedSku` / `Get-MgGroup` for finding the right values for your tenant.

When the single-user onboarding asks "Apply a role-based onboarding template?", choosing one skips the file/manual/replicate picker — the operator is prompted only for the user-specific fields (name, UPN, optional manager and job title) and the template fills the rest. Unknown SKUs / groups / DLs / shared mailboxes are skipped with a per-item warning; the onboard continues. Set `contractorExpiryDays` to record `employeeLeaveDateTime` automatically (auto-disable on that date still needs Entra lifecycle workflows).

## Bulk operations

Two CSV-driven workflows, available from the main menu and as scriptable entry points:

```powershell
Invoke-BulkOnboard  -Path users.csv   [-Template <name>] [-WhatIf]
Invoke-BulkOffboard -Path leavers.csv [-WhatIf]
```

Both follow the same pattern: validate the whole CSV first (required fields, UPN format, duplicate UPNs, unknown templates), show errors per row, then ask for confirmation before any tenant call. Per-row failures don't halt the batch — every row gets a line in a result CSV written next to the input (`bulk-onboard-<ts>.csv` / `bulk-offboard-<ts>.csv`) with `Status` ∈ `Success | PartialSuccess | Failed | Preview` and a `Reason` column.

Samples to copy from: `templates/bulk-onboard-sample.csv`, `templates/bulk-offboard-sample.csv`. Onboard CSV column reference:

| Column | Required | Notes |
|---|---|---|
| `FirstName` / `LastName` | yes | |
| `DisplayName` | no | Defaults to `FirstName LastName`. |
| `UserPrincipalName` (or `UPN`) | yes | Used as both sign-in name and primary email. |
| `Manager` | no | UPN of an existing user; skipped with a warning if not found. |
| `Department`, `JobTitle`, `OfficeLocation`, `UsageLocation` | mixed | `UsageLocation` is required (template fills it if blank). |
| `Template` | no | Role-template key (`sales-rep`, `engineer`, ...). Row value wins over `-Template` flag. |
| `Password` | no | If blank, a 16-char password is generated. Either way, ForceChangePasswordNextSignIn is set. |
| `IssueTAP` | no | `true` / `yes` / `1` issues a 60-minute single-use Temporary Access Pass via `/authentication/temporaryAccessPassMethods`. |

Offboard CSV columns: `UserPrincipalName, ForwardTo, ConvertToShared, HandoffOneDriveTo, RemoveFromAllGroups, Reason`. `HandoffOneDriveTo` is validated and surfaced in the result CSV but the actual SPO transfer is a Phase 3 deliverable — for now the column being set logs a TODO line and the rest of the row continues.

## Preview / dry-run mode

On launch the tool asks `LIVE` vs `PREVIEW`. Every menu refresh paints a colored badge above the main menu so the mode is visible at a glance — red `[LIVE]` or yellow `[PREVIEW]`. The choice is session-scoped; switch via `Switch Tenant` (which re-shows the picker) or by quitting and re-launching.

In `PREVIEW` mode, every state-mutating call is routed through `Invoke-Action` (see `Preview.ps1`). The cmdlet does **not** run; instead a `[PREVIEW] Would: <description>` line is written to the console and to the per-session audit log. Stub return values let downstream code (e.g. anything that reads `$newUser.Id`) keep working. The pattern at every call site is:

```powershell
$ok = Invoke-Action -Description "Block sign-in for $upn" -Action {
    Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop; $true
}
if ($ok -and -not (Get-PreviewMode)) { Write-Success "Sign-in blocked." }
```

Bulk operations accept `-WhatIf` independently — it flips `$script:PreviewMode` to `$true` for the duration of that single call (try/finally restores it on return) so a one-off dry-run does not disturb the rest of the session.

Audit log location: `%LOCALAPPDATA%\M365Manager\audit\session-<ts>-<pid>.log` (Windows) or `~/.m365manager/audit/session-<ts>-<pid>.log` (POSIX, mode `0700`). Every line is tagged `MODE=PREVIEW` or `MODE=LIVE` — grep one or the other to filter. A sample preview log shape is at `docs/preview-sample.log`.

## Audit & reporting

A new "Audit & Reporting..." main-menu entry groups everything an operator might need after the fact. Five sub-entries:

1. **Audit log viewer** — `Show-AuditLogViewer`. Loads every `session-*.log` (and optionally `mark-*.log`) under `%LOCALAPPDATA%\M365Manager\audit\`, normalizes the JSON-line and legacy human-readable shapes to one record format, and offers paged display with filters (UPN, date range, action type, event, mode, result, target substring). Export the current filter to CSV or single-file HTML.
2. **Undo recent operation...** — see "Undo" below.
3. **Sign-in lookup (Graph search)** — filter wizard over `/v1.0/auditLogs/signIns` (see "Sign-in lookup").
4. **Sign-in lookup (recent activity for one user)** — UPN one-shot, last 30 days.
5. **Unified audit log** — wraps `Search-UnifiedAuditLog`; health-checks ingestion is enabled first.

The session log itself is now JSON-per-line. See `docs/audit-format.md` for the field reference. The viewer transparently handles both new JSONL and pre-Phase-2 human-readable lines.

## Undo recent operations

Reversible state changes (license assignment, group / DL membership, mailbox FullAccess + SendAs grants, calendar permissions, OOO, forwarding, sign-in block) write a `reverse` recipe into their audit entry: a stable type tag plus the operand hashtable needed to dispatch the inverse cmdlet. The "Undo recent operation..." menu walks the audit log, lists everything that's still reversible (success + reverse + not-already-reversed), and lets the operator pick a row to invert.

`Invoke-Undo` runs the reverse through `Invoke-Action` so the reversal itself is audited and respects PREVIEW mode (you can stage an undo with `-WhatIf` before committing). On success, the original entry is marked reversed in `audit/undo-state.json` so it doesn't reappear.

Destructive operations (user / mailbox / DL / group deletion, sign-in session revocation, shared-mailbox conversion, MFA method revocation, account creation) carry an explicit `noUndoReason` and are excluded from the undo list. If a reverse cmdlet fails because the target object no longer exists, the operator sees the underlying Graph / EXO error (the default chosen here is "fail clearly with the target id", not auto-mark-superseded).

## Sign-in lookup

`SignInLookup.ps1` wraps `GET /v1.0/auditLogs/signIns` (Microsoft Graph). Required scopes: `AuditLog.Read.All` + `Directory.Read.All`; missing scopes are surfaced as a clear warning before any query runs. Pagination through `@odata.nextLink` is automatic; Graph caps a page at 1000.

Two surfaces:
- Filter wizard (UPN, app, IP, country, risk level, CA status, success/failure-only, date range as `7d` / `24h` / `YYYY-MM-DD / YYYY-MM-DD`).
- Recent-activity-for-one-user (last 30 days, single prompt).

Output: paged table + on-demand CSV + single-file HTML with risk-tier row shading.

## Unified audit log

`UnifiedAuditLog.ps1` wraps `Search-UnifiedAuditLog` (EXO / Purview). Health-checks `Get-AdminAuditLogConfig` first — if `UnifiedAuditLogIngestionEnabled` is false, the viewer prints the exact enable command and aborts. Retention varies by license; the module surfaces a best-effort estimate (E5 → 365d, E3 → 180d, otherwise the generic "90 / 180 / 365") so the operator knows the floor.

The wizard prompts for an operation **group** (Mail, File access, Identity, Groups, Compliance, Sharing) and lets you narrow to specific operations within it, then runs the query and offers CSV + HTML export.

## MFA management

New "MFA & Authentication..." main-menu entry plus `MFAManager.ps1`. Operations per-user:

- **View methods** — table by type, color-coded by strength.
- **Revoke specific method** — pick from list.
- **Revoke all methods** — locks the user out until re-registration.
- **Reset MFA** — revoke all + revoke sign-in sessions + issue a fresh single-use TAP (clipboard-delivered with the same scrub-on-Enter pattern as onboard).
- **Issue Temporary Access Pass** — configurable lifetime (1 / 8 / 24 hours), one-time vs reusable.

Four compliance views (each prompts for a scan cap; full-tenant scans are slow):
- Users with only SMS / voice MFA.
- Users with no MFA registered.
- Users with an active TAP.
- Users with FIDO2 keys.

Bulk: `Invoke-BulkMfa -Path .csv [-WhatIf]` over a CSV with columns `UPN, Action` (`Revoke` / `Reset` / `IssueTAP`), optional `TAPLifetime` and `TAPUsableOnce`.

Integration with existing flows:
- Onboard / BulkOnboard's TAP issuance now delegates to `New-TemporaryAccessPass` — single source of truth, single audit trail.
- Offboard / BulkOffboard adds a new Step 0 that calls `Remove-AllAuthMethods` before session revocation, so a hostile actor with the user's device cannot re-authenticate during the window between session revoke and account block.

## Tests

```powershell
Invoke-Pester ./tests/
```

Phase 2 adds three Pester suites alongside the privacy tests:
- `tests/AuditViewer.Tests.ps1` — JSONL + legacy + AI line parsing, filter correctness, CSV/HTML export shape.
- `tests/Undo.Tests.ps1` — dispatch-table coverage, target-hashtable round-trip, state-transitions (mocked sidecar so the test doesn't touch real machine state).
- `tests/MFAManager.Tests.ps1` — method-type classification, compliance-view predicates against canned method sets.

No network or M365 connection required for any test.

## File Structure

```
M365Admin/
├── Launch.bat             # Double-click to start the tool
├── Main.ps1               # Entry point, main menu, session status bar
├── UI.ps1                 # Shared TUI helpers (colors, menus, confirmations, box drawing)
├── Auth.ps1               # Browser-based OAuth authentication & session management
├── Onboard.ps1            # New user onboarding (file, manual, or replicate)
├── Offboard.ps1           # User offboarding workflow (7 steps)
├── License.ps1            # Add / remove license management
├── Archive.ps1            # Mailbox archiving & retention policies
├── SecurityGroup.ps1      # Security group membership
├── DistributionList.ps1   # Distribution list membership & permissions
├── SharedMailbox.ps1      # Shared mailbox access & permissions
├── CalendarAccess.ps1     # Calendar permission management
├── UserProfile.ps1        # View & edit full user profile
├── Reports.ps1            # Interactive reporting with sort & CSV export
├── eDiscovery.ps1         # eDiscovery simple & advanced mode
├── GroupManager.ps1       # Unified group membership view, edit, replicate
├── samples/
│   ├── onboard_sample.txt # Sample onboard input file
│   └── offboard_sample.txt# Sample offboard input file
└── README.md
```

## Features

| # | Feature | Services Used |
|---|---------|---------------|
| 1 | **Onboard** | Graph, EXO |
| 2 | **Offboard** | Graph, EXO |
| 3 | **License Management** | Graph |
| 4 | **Mailbox Archiving** | Graph, EXO |
| 5 | **Security Groups** | Graph |
| 6 | **Distribution Lists** | Graph, EXO |
| 7 | **Shared Mailboxes** | Graph, EXO |
| 8 | **Calendar Access** | Graph, EXO |
| 9 | **User Profile** | Graph |
| 10 | **Group Membership Manager** | Graph, EXO |
| 11 | **Reporting** | Graph, EXO |
| 12 | **eDiscovery** | SCC, Graph |

---

### 1. Onboard New User

Three ways to provide new user data:

- **Parse from text file** — reads a `Key: Value` file, shows parsed results for review/edit before proceeding.
- **Manual entry** — prompts for each field one by one with a review/edit table.
- **Replicate from existing user** — searches for an existing user by name or email, then copies their full profile (job title, department, company, office, address, usage location, phone numbers, manager), all assigned licenses, and all security group memberships. Name and email are left blank for you to fill in. After review, the tool walks through each replicated license and group with per-item confirmation, and offers to add extras.

Onboarding steps:
1. Create account with all profile fields
2. Assign licenses (replicated or picked from tenant)
3. Assign security groups (replicated or searched)
4. Set manager
5. Display temporary password

Profile fields supported: FirstName, LastName, DisplayName, UPN, JobTitle, Department, CompanyName, OfficeLocation, StreetAddress, City, State, PostalCode, Country, UsageLocation, BusinessPhone, MobilePhone, Manager.

### 2. Offboard User

Supports text file parsing or manual entry for offboarding data.

Offboarding steps:
1. **Revoke all sessions** — invalidates all refresh tokens immediately
2. **Block sign-in** — disables the account
3. **Set Out-of-Office** — configures internal and external auto-reply
4. **Email forwarding** — validates target in tenant, choice of forward-only or forward-and-keep-copy
5. **Convert to shared mailbox**
6. **Grant mailbox access** — loop to add one or more users with Full Access, plus optional Send As and/or Send on Behalf per user
7. **Remove licenses** — lists all current licenses and removes them

### 3. Add / Remove License

- Search user by name or email
- Shows all current licenses with **friendly names** (e.g. "Microsoft 365 E3" instead of "SPE_E3")
- Tags each license as **[Direct]**, **[Group-assigned]**, or **[Direct + Group]**
- Add: displays full tenant license inventory with friendly names and total/used/free counts, multi-select
- Remove: blocks removal of group-assigned licenses with an explanation of which group assigned it, only removes directly-assigned licenses

### 4. Mailbox Archiving

- Shows current archive status and retention policy
- Enable archive and start Managed Folder Assistant immediately
- Choose or change retention policy from available tenant policies

### 5. Security Group Management

- **Create** — set name, description, mail nickname, choose mail-enabled or standard. Option to add members immediately after creation.
- **Add / remove members** — find a group by name, view current members, then add (with user search loop) or remove (multi-select with commas).
- **View / edit properties** — shows name, description, mail nickname, mail-enabled status, member count. Edit name, description, or mail nickname.
- **Delete** — shows group details and member count, requires confirmation plus typing the group name to prevent accidents.

### 6. Distribution List Management

- **Create** — set name, alias, primary email, owner, sender restrictions (open vs internal-only). Option to add members immediately after creation.
- **Add / remove members** — find a user, add to a DL with optional Send As / Send on Behalf, or list all DLs a user belongs to for multi-select removal.
- **View / edit properties** — shows name, email, alias, owner, sender auth, GAL visibility, member count, Send on Behalf list. Edit name, description, owner, toggle sender auth, toggle hidden from address book.
- **Delete** — shows DL details and member count, requires confirmation plus typing the DL email address.

### 7. Shared Mailbox Management

- **Create** — set name, email, alias. Option to grant access to users immediately after creation.
- **Add / remove user access** — shows current Full Access and Send As permissions. Add users in a loop (Full Access + optional Send As / Send on Behalf per user), or multi-select remove.
- **View / edit properties** — shows name, email, alias, type, GAL visibility, forwarding, archive status, aliases, auto-reply state, Send on Behalf list. Edit name, add/remove email aliases, set/remove forwarding, toggle GAL visibility, configure auto-reply.
- **Delete** — shows mailbox details and current permissions, requires confirmation plus typing the email address. Warns that all mailbox contents will be permanently lost.

### 8. Calendar Access Management

- Identifies the calendar owner, shows current permissions
- Add: search for user to grant access, choose level (Reviewer, Editor, Author, Contributor), auto-updates if permission already exists
- Remove: lists all custom permissions, multi-select removal

### 9. User Profile Management

- Displays the full Azure AD user profile in a formatted view
- Editable fields (numbered): Display Name, First/Last Name, Job Title, Department, Company Name, Office Location, full address, Usage Location, Business/Mobile/Fax phone, Mail Nickname
- Read-only fields (unnumbered): UPN, Mail, Proxy Addresses, Other Emails, Account Enabled, User Type, Object ID, Dir Sync status
- Set or remove manager with search/confirm
- Refresh to see changes in real time
- Loops until you pick Done, so multiple edits in one session

### 10. Group Membership Manager

Unified view of all group types (Security Groups, Distribution Lists, M365 Groups, Mail-Enabled SGs) for any user. Fetches from both Microsoft Graph and Exchange Online to get the complete picture.

- **View all memberships** — shows every group a user belongs to, organized by type with counts
- **Bulk remove** — filter by group type (SG/DL/M365), multi-select for removal across all types in one operation
- **Add to groups** — search by name, choose Security Group, Distribution List, or search all types. Multi-select add.
- **Replicate memberships** — copy group memberships from one user to another with three modes:
  - **Selective** — pick which groups to copy, shows which ones the target already has
  - **Full Copy** — add all source groups to target, keeping any groups the target already has
  - **Full Replace** — remove target's current groups that aren't in the source, add any that are missing. Shows a full diff (removals in red, additions in green, unchanged count) before confirming.

### 11. Reporting

All reports share these common features:

- **Scope picker** — run against the entire org, a single user, or members of a specific group
- **Active/inactive filter** — before each report, choose to include all users, active (enabled) users only, or disabled users only
- **Interactive sorting** — after data is collected, pick any column to sort by (ascending or descending) before display
- **TUI display** — formatted table in the console, capped at 50 rows for readability
- **CSV export** — export full data to `C:\Temp` or a custom folder, timestamped filename (e.g. `UserLicenses_20260406_141500.csv`)

Available reports:

| Report | Description |
|--------|-------------|
| **License Summary** | Every SKU in the tenant with friendly name, total/assigned/available seats, usage percentage |
| **Per-User Licenses** | Each user with every license they hold, filterable by active/inactive status |
| **Mailbox Sizes** | Item count, total size, last activity date per mailbox |
| **Archive Mailboxes** | Archive enabled/disabled, retention policy, archive size and item count |
| **User Account Status** | Enabled/disabled, department, job title, license count, creation date, last sign-in |
| **Group Membership** | Three modes: all groups for a user, all members of a group (with active filter), or all security groups with member counts |
| **Shared Mailboxes** | Every shared mailbox with size, item count, access count, forwarding, GAL visibility |
| **Inactive Users** | Configurable threshold (default 90 days), shows users with no sign-in, with license count to spot wasted seats. Can filter to enabled-only (consuming licenses) or disabled-only (cleanup candidates) |

### 12. eDiscovery

Two modes: **Simple** for quick ad-hoc searches, **Advanced** for full case-based investigations.

**Simple Mode — Quick Content Search**
- **New quick search** — guided wizard that builds a KQL query from prompts: keywords, date range, sender, recipient, message type (email/documents/IM/meetings). Choose to search all mailboxes or specific ones. Automatically creates and starts the search.
- **View existing searches** — lists all content searches with status, item count, size, query, and creation date.
- **View search results** — detailed view of a completed search including per-mailbox breakdown. Actions: create a preview, create an export (PST/individual messages/both), re-run the search, or check export status.
- **Delete a search** — removes the search and all associated actions, with double confirmation.

**Advanced Mode — Case Management**

*Case Management:*
- Create new eDiscovery Standard or Premium (Advanced) cases
- List all cases with status, type, and creation date
- View case details including all searches and holds within it
- Add members to a case
- Close, reopen, or delete cases (only closed cases can be deleted, with double confirmation)

*Search Management (within a case):*
- **Guided query builder** — same wizard as Simple mode: keywords, dates, sender, recipient, subject, attachment filters, message type
- **Raw KQL mode** — type any KQL query directly, with a built-in reference card showing common operators (`from:`, `to:`, `subject:`, `sent>=`, `kind:`, `hasattachment:`, `filetype:`, `participants:`, `cc:`, `bcc:`, `size>`)
- Search scope: all mailboxes, specific mailboxes, all SharePoint sites, or both
- Run/re-run searches, view results, delete searches

*Hold Management:*
- **Create legal hold** — place one or more mailboxes on hold within a case, with optional KQL query to hold only matching content
- **List holds** — shows all holds in a case with locations, enabled status, and query filters
- **Modify hold** — add/remove mailboxes, enable/disable, update the KQL query filter
- **Remove hold** — releases held content with confirmation

*Export Management:*
- View all export and preview actions across searches with status, creation time, and result counts

---

## Authentication

All authentication uses **browser-based interactive OAuth**. On startup, the tool asks whether you're managing your own organization or a customer tenant via GDAP.

### Direct Mode (Own Tenant)
Standard admin login. `Connect-MgGraph` and `Connect-ExchangeOnline` connect to your own tenant.

### GDAP Partner Mode (Customer Tenants)
For Microsoft Partner Center / CSP organizations with Granular Delegated Admin Privileges:

1. Tool first connects to your **partner tenant** to list all customer relationships
2. Fetches customer tenants via Microsoft Graph contracts API
3. Presents a searchable list of all managed customers (name + domain)
4. After you pick a customer, disconnects from partner and reconnects to the **customer tenant** using `-TenantId` (Graph) and `-DelegatedOrganization` (EXO)
5. All subsequent operations run against the selected customer tenant
6. **Switch Tenant** option on the main menu lets you disconnect and pick a different customer without restarting

Required GDAP roles on the customer tenant: User Administrator, Groups Administrator, License Administrator, Exchange Administrator, Directory Readers. For eDiscovery: eDiscovery Manager or eDiscovery Administrator.

If the contract listing fails (permissions, no relationships), the tool falls back to manual tenant ID / domain entry.

Only two modules are required: **Microsoft Graph PowerShell SDK** and **ExchangeOnlineManagement** (which also provides `Connect-IPPSSession` for eDiscovery/Security & Compliance Center).

## Design Principles

- **Browser-based OAuth** — modern interactive login via Microsoft Graph. Supports MFA, conditional access, SSO, and GDAP partner access.
- **GDAP partner support** — manage customer tenants with a built-in tenant picker. Switch between customers from the main menu.
- **Friendly license names** — displays "Microsoft 365 E3" instead of "SPE_E3", with group-assignment detection.
- **Confirmation on every change** — all destructive or modifying steps show a yellow confirmation prompt before executing.
- **Modular** — each function lives in its own `.ps1` file.
- **Encoding-safe** — all box-drawing characters are built at runtime via `[char]` hex codes, so files survive any encoding/download.
- **Text file parsing** — Onboard and Offboard support `Key: Value` files with flexible label matching.
- **Multi-select** — license assignment, group removal, and similar operations support comma-separated selections (e.g. `1,3,5`).
- **Clean disconnect** — quitting the app disconnects all active PowerShell sessions.

## Adding a New Feature

1. Create `NewFeature.ps1` with a `Start-NewFeature` function.
2. Dot-source it in `Main.ps1`:
   ```powershell
   . "$ScriptRoot\NewFeature.ps1"
   ```
3. Add the task name to the `ValidateSet` and `$map` in `Connect-ForTask` inside `Auth.ps1`.
4. Add the menu entry in `Main.ps1`'s options array and switch block.

## Sample Input Files

See `samples/` for the expected `Key: Value` format. The parser is flexible with labels — e.g., `First Name`, `FirstName`, or `first name` all map correctly.
