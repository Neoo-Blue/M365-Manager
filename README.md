# M365 Administration Tool

A modular PowerShell TUI application for common Microsoft 365 admin tasks. Blue background, multi-color interface, confirmation prompts on every change, and browser-based OAuth authentication.

## Prerequisites

Install the required PowerShell modules (the tool will prompt if missing):

```powershell
Install-Module AzureAD -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module MSOnline -Scope CurrentUser -Force
```

## Launch

```powershell
powershell -ExecutionPolicy Bypass -File Main.ps1
```

## File Structure

```
M365Admin/
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
├── samples/
│   ├── onboard_sample.txt # Sample onboard input file
│   └── offboard_sample.txt# Sample offboard input file
└── README.md
```

## Features

| # | Feature | Services Used |
|---|---------|---------------|
| 1 | **Onboard** | AzureAD, MSOnline, EXO |
| 2 | **Offboard** | AzureAD, MSOnline, EXO |
| 3 | **License Management** | AzureAD, MSOnline |
| 4 | **Mailbox Archiving** | AzureAD, EXO |
| 5 | **Security Groups** | AzureAD |
| 6 | **Distribution Lists** | AzureAD, EXO |
| 7 | **Shared Mailboxes** | AzureAD, EXO |
| 8 | **Calendar Access** | AzureAD, EXO |
| 9 | **User Profile** | AzureAD |

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
4. Enforce MFA
5. Set manager
6. Display temporary password

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
- Shows all current licenses on the user
- Add: displays full tenant license inventory with total/used/free counts, multi-select with commas
- Remove: lists user's licenses, multi-select removal

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

---

## Authentication

All authentication uses **browser-based interactive OAuth** — no stored passwords, no credential dialogs. When a task needs a service that isn't connected yet, a browser window opens for Microsoft sign-in. The session status bar on the main menu shows which services (AzureAD, EXO, MSOnline) are currently connected.

Each service connects at most once per session. Quitting the app disconnects all active sessions.

## Design Principles

- **Browser-based OAuth** — modern interactive login that supports MFA, conditional access, and SSO across services.
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
