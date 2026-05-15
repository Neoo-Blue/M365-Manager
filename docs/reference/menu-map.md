# Menu map

Complete navigation tree for the M365 Manager TUI. Slot numbers come straight from `Main.ps1:205-228`.

## Top-level main menu

```
1   Onboard New User                  -> Start-Onboard
2   Offboard User                     -> Start-Offboard
3   Add / Remove License              -> Start-LicenseManagement
4   Mailbox Archiving                 -> Start-ArchiveManagement
5   Security Group Management         -> Start-SecurityGroupManagement
6   Distribution List Management      -> Start-DistributionListManagement
7   Shared Mailbox Management         -> Start-SharedMailboxManagement
8   Calendar Access Management        -> Start-CalendarAccessManagement
9   User Profile Management           -> Start-UserProfileManagement
10  Group Membership Manager          -> Start-GroupManagerMenu
11  Reporting                         -> Start-ReportingMenu
12  eDiscovery                        -> Start-eDiscoveryMenu
13  Bulk Onboard from CSV...          -> Start-BulkOnboard
14  Bulk Offboard from CSV...         -> Start-BulkOffboard
15  Audit & Reporting...              -> Start-AuditReportingMenu
16  MFA & Authentication...           -> Start-MFAMenu
17  Teams Management...               -> Start-TeamsMenu
18  SharePoint...                     -> Start-SharePointMenu
19  Guest Users...                    -> Start-GuestUsersMenu
20  License & Cost...                 -> Start-LicenseOptimizerMenu
21  Scheduled Health Checks...        -> Start-SchedulerMenu
22  Tenants...                        -> Start-TenantMenu                (Phase 6)
23  Incident Response...              -> Start-IncidentResponseMenu      (Phase 7)
99  (hidden) AI Assistant             -> Start-AIAssistant
-   Quit and Disconnect               -> Disconnect-AllSessions + exit
```

Slot 99 is hidden (not shown in the menu) — typed as `99` even though the visible list goes 1-23.

## Slot 11 — Reporting

```
1   Tenant overview
2   User report (active / disabled / age)
3   License usage by SKU
4   Group membership summary
5   Mailbox size distribution
6   Recently created users / groups / sites
```

## Slot 14 — Audit & Reporting

```
1   Open audit log viewer
2   Filter audit log (current session)
3   Export to CSV
4   Export to HTML
5   List recent undoable operations
6   Undo a specific operation (by entry id)
7   View undo state
8   Sign-in lookup            (Search-SignIns)
9   Unified audit log search  (Search-UAL)
```

## Slot 15 — MFA & Authentication

```
1   List a user's MFA methods
2   Revoke a specific MFA method
3   Revoke ALL MFA methods (incident-response use)
4   Issue Temporary Access Pass
5   Compliance views
      a. Users with NO MFA registered
      b. Users with ONLY phone-based MFA
      c. Users with an active TAP
      d. Users registered for FIDO2
      e. Users with self-service password reset configured
6   CSV export (all users, all methods)
7   Bulk MFA reset (CSV-driven)
```

## Slot 17 — Teams Management

```
1   List a user's teams
2   Add a user to a team
3   Remove a user from a team
4   Promote / demote owner
5   Orphan-team report (zero owners)
6   Single-owner classification report
7   Bulk Teams membership from CSV
```

## Slot 18 — SharePoint

```
1   List sites
2   Create a site (from template)
3   Add / remove site owner
4   List a user's outbound shares
5   Revoke a single share
6   Revoke ALL of a user's outbound shares
7   Stale-site report
```

## Slot 19 — Guest Users

```
1   List all guests
2   Stale-guest report
3   Guests grouped by inviter / domain
4   Recertification campaign (start / progress / apply)
5   Remove a guest (full teardown)
6   Bulk guest removal from CSV
```

## Slot 20 — License & Cost

```
1   License optimizer (find waste + project savings)
2   Remediate selected findings (PREVIEW then LIVE)
3   Bulk license remediation from CSV
4   Cost report (rollup by SKU + by department)
5   Notifications setup        -> Start-NotificationsSetup
```

## Slot 21 — Scheduled Health Checks

```
1   List scheduled checks
2   Register a new scheduled check
3   Unregister a scheduled check
4   Test-NotificationChannels (verify configured channels work)
5   Test-ScheduledHealthCheck   (verify credential decrypts)
6   Test-BreakGlassPosture      (audit break-glass accounts)
```

## Slot 22 — Tenants

```
1   List registered tenants
2   Register a new tenant
3   Switch to a tenant
4   Update tenant metadata
5   Remove a tenant
6   MSP portfolio dashboard      -> Update-MSPDashboard
7   Cross-tenant report (license / MFA / stale / orphan / breakglass)
```

## Slot 23 — Incident Response

```
1   Run compromised-account response (single UPN)
2   Run bulk incident response (CSV)
3   Run tabletop exercise
4   List open incidents
5   List incidents (all)
6   View incident report
7   Close incident
8   Mark incident as false positive (with undo walk)
9   Undo an incident's reversible steps
10  Export incident for compliance handoff
```

## Slot 99 — AI Assistant (hidden)

Drops into the AI chat REPL. Slash commands: see [`chat-commands.md`](chat-commands.md).

## See also

- [`chat-commands.md`](chat-commands.md) — slash commands inside the AI assistant.
- [`csv-formats.md`](csv-formats.md) — every bulk CSV schema.
- [`../guides/`](../guides/) — task-oriented walkthroughs of each menu area.
