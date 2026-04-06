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

Double-click `Launch.bat` in the M365Admin folder.

Or run manually from PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File Main.ps1
```

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
