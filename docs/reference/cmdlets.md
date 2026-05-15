# Public function reference

Every public function in the tool, grouped by module. Brief one-liner per function; for detail run `Get-Help <function>` or read the module source.

## Foundation

### `UI.ps1`

| Function | Purpose |
|---|---|
| `Initialize-UI` | Set up colors + read terminal capabilities. |
| `Write-Banner` | Print the M365 Admin banner. |
| `Write-SectionHeader` | Print a section title bar. |
| `Write-StatusLine` | One-line status output (key/value/color). |
| `Write-InfoMsg` / `Write-Warn` / `Write-ErrorMsg` / `Write-Success` | Standard log helpers. |
| `Read-UserInput` | Prompt the operator. Returns empty in NonInteractive. |
| `Confirm-Action` | Y/N prompt. Returns `$false` in NonInteractive. |
| `Show-Menu` | Numbered menu picker. Returns -1 on back/cancel. |
| `Pause-ForUser` | "Press any key." No-op in NonInteractive. |
| `Set-NonInteractiveMode` / `Get-NonInteractiveMode` | Flip the non-interactive flag. |

### `Auth.ps1`

| Function | Purpose |
|---|---|
| `Connect-ForTask <area>` | Lookup table — connects every service the area needs. |
| `Reset-AllSessions` | Disconnect Graph + EXO + SCC + SPO. |
| `Disconnect-AllSessions` | Same as Reset, plus the partner-center session. |
| `Get-StateDirectory` | Resolve `<stateDir>` (per-OS). |
| `Select-TenantMode` | Direct / partner / registered profile picker. |

### `Audit.ps1`

| Function | Purpose |
|---|---|
| `Write-AuditEntry` | Write one JSONL line to the session log. |
| `Get-AuditLogPath` | Resolve the session log path (with tenant slug). |
| `Reset-AuditLogPath` | Force regeneration after tenant switch. |
| `New-AuditEntryId` | UUID generator. |

### `Preview.ps1`

| Function | Purpose |
|---|---|
| `Set-PreviewMode <bool>` / `Get-PreviewMode` | Flip / read PREVIEW mode flag. |
| `Invoke-Action` | The mutation wrapper. PROPOSE / EXEC / OK / ERROR audit, reverse-recipe stamping. |

### `Notifications.ps1`

| Function | Purpose |
|---|---|
| `Send-Email` | Send via Graph `/me/sendMail` or configured SMTP. |
| `Send-TeamsCard` | Post adaptive card to a webhook. |
| `Send-Notification` | High-level dispatcher (channels + severity). |
| `Test-NotificationChannels` | Verify configured channels by sending a test. |
| `Start-NotificationsSetup` | Wizard for configuring recipients + webhooks. |

### `TenantRegistry.ps1` / `TenantSwitch.ps1` / `TenantOverrides.ps1` (Phase 6)

| Function | Purpose |
|---|---|
| `Register-Tenant` | Add a new tenant profile. |
| `Update-Tenant` | Patch an existing profile. |
| `Remove-Tenant` | Delete a profile + its secret manifest. |
| `Get-Tenants` | Enumerate all registered. |
| `Get-Tenant -Name <n>` | Single profile by name. |
| `Get-CurrentTenant` / `Set-CurrentTenant` | Active tenant slot. |
| `Switch-Tenant -Name <n>` | Reset sessions + reconnect to the named tenant. |
| `Get-EffectiveConfig -Key <k>` | Resolve through Global → Tenant → Env → CLI. |
| `Get-TenantOverrides` / `Set-TenantOverrides` | Read / write `<stateDir>\tenant-overrides\<tenant>.json`. |

## Lifecycle

### `Onboard.ps1` / `BulkOnboard.ps1`

| Function | Purpose |
|---|---|
| `Start-Onboard` | Interactive single-user flow. |
| `Invoke-BulkOnboard -Path <csv>` | CSV-driven bulk. |
| `New-OnboardedUser` | Underlying primitive — create + license + groups. |
| `Get-OnboardingTemplate -Name <n>` | Load a role-template JSON. |

### `Offboard.ps1` / `BulkOffboard.ps1`

| Function | Purpose |
|---|---|
| `Start-Offboard` | Interactive 12-step flow. |
| `Invoke-BulkOffboard -Path <csv>` | CSV-driven bulk. |
| `Invoke-OffboardOrchestration` | The 12-step state machine called by both above. |

### `GuestUsers.ps1`

| Function | Purpose |
|---|---|
| `Get-Guests` | Enumerate guest users. |
| `Get-StaleGuests` | Filter by `DaysSinceSignIn >= threshold`. |
| `Get-GuestsByInviter` / `Get-GuestsByDomain` | Pivots. |
| `Start-RecertCampaign` | Begin a recertification campaign. |
| `Show-PendingRecerts` | Walk pending decisions. |
| `Remove-Guest -UPN <u> -Reason <r>` | Full teardown (shares + groups + teams + delete). |
| `Invoke-BulkGuestRemoval -Path <csv>` | Bulk. |

## Lifecycle helpers

### `License.ps1` / `LicenseOptimizer.ps1`

| Function | Purpose |
|---|---|
| `Assign-License -User <u> -SkuId <id>` | Audited. |
| `Remove-License -User <u> -SkuId <id>` | Audited; reversible. |
| `Get-LicenseOptimizationReport` | Three-category findings. |
| `Invoke-BulkLicenseRemoval -Path <csv>` | Bulk remediation. |

### `Archive.ps1`

| Function | Purpose |
|---|---|
| `Enable-MailboxArchive -UPN <u>` / `Disable-MailboxArchive -UPN <u>` | Toggle. |

### `MFAManager.ps1`

| Function | Purpose |
|---|---|
| `Get-UserAuthMethods -User <u>` | Enumerate. |
| `Remove-AuthMethod -User <u> -MethodInfo <obj>` | Revoke one. |
| `Remove-AllAuthMethods -User <u>` | Revoke all. |
| `New-TemporaryAccessPass -User <u>` | Issue a TAP. |
| `Get-UsersWithNoMfa` / `Get-UsersWithOnlyPhoneMfa` / `Get-Fido2Users` | Compliance views. |

### `OneDriveManager.ps1`

| Function | Purpose |
|---|---|
| `Grant-OneDriveAccess -LeaverUPN <l> -TargetUPN <t>` | Add target as site owner. |
| `Revoke-OneDriveAccess -LeaverUPN <l> -TargetUPN <t>` | Remove. |
| `Invoke-OneDriveHandoff` | Orchestration (grant + retention + email). |
| `Get-OneDriveRecentFiles -UPN <u>` | Recent activity. |

### `TeamsManager.ps1`

| Function | Purpose |
|---|---|
| `Get-UserTeams -UPN <u>` | List teams + roles. |
| `Add-UserToTeam` / `Remove-UserFromTeam` | Audited; reversible. |
| `Set-TeamOwnership` (promote/demote) | Audited; reversible. |
| `Get-OrphanedTeams` / `Get-SingleOwnerTeams` | Reports. |
| `Invoke-TeamsOffboardTransfer -LeaverUPN <l>` | Sole-owner handoff. |
| `Invoke-BulkTeamsMembership -Path <csv>` | Bulk. |

### `SharePoint.ps1`

| Function | Purpose |
|---|---|
| `Add-SiteOwner` / `Remove-SiteOwner` | Audited; reversible. |
| `Get-UserOutboundShares -UPN <u>` | UAL-driven outbound list. |
| `Revoke-Share -ShareId <id>` | Single share. |
| `Invoke-SharePointOffboardCleanup -LeaverUPN <l>` | Bulk revoke for offboard. |
| `New-SiteFromTemplate -Template <t>` | Provisioning. |

### Other modules

`SecurityGroup.ps1`, `DistributionList.ps1`, `SharedMailbox.ps1`, `CalendarAccess.ps1`, `UserProfile.ps1`, `GroupManager.ps1`, `Reports.ps1`, `eDiscovery.ps1` follow the same `Add-* / Remove-* / Set-* / Get-*` pattern with `Invoke-Action` wrapping for mutations.

## Audit + undo

### `AuditViewer.ps1`

| Function | Purpose |
|---|---|
| `Read-AuditEntries [-Path]` | Load + parse JSONL. |
| `Filter-AuditEntries -Entries -Filter <hashtable>` | Filter by mode / event / actionType / user / target / date. |
| `ConvertFrom-AuditLine` | Single-line parser. |
| `Export-AuditEntriesCsv` / `Export-AuditEntriesHtml` | Exports. |

### `Undo.ps1`

| Function | Purpose |
|---|---|
| `Invoke-Undo -EntryId <id>` | Dispatch by `reverse.type`. |
| `Show-RecentUndoable` | Last N reversible operations. |
| `Get-UndoableEntries [-Filter] [-Since] [-Limit]` | Programmatic equivalent. |
| `Read-UndoState` / `Write-UndoState` | The sidecar state file. |

## Security + audit

### `SignInLookup.ps1`

| Function | Purpose |
|---|---|
| `Search-SignIns [-User] [-From] [-To] ...` | Graph `/auditLogs/signIns`. |

### `UnifiedAuditLog.ps1`

| Function | Purpose |
|---|---|
| `Search-UAL [-UserId] [-From] [-Operations] ...` | EXO `Search-UnifiedAuditLog` wrapper. |

### `BreakGlass.ps1` (Phase 4)

| Function | Purpose |
|---|---|
| `Register-BreakGlassAccount -UPN <u>` | Add to registry. |
| `Unregister-BreakGlassAccount -UPN <u>` | Remove. |
| `Get-BreakGlassAccounts` | List. |
| `Test-BreakGlassPosture` | Predicate checks (password age, recent sign-in). |

## Cost + health

### `Scheduler.ps1`

| Function | Purpose |
|---|---|
| `Register-ScheduledHealthCheck -CheckName <n> -Trigger <t>` | Create Task Scheduler entry. |
| `Unregister-ScheduledHealthCheck -CheckName <n>` | Remove. |
| `Get-ScheduledHealthChecks` | List. |
| `Test-ScheduledHealthCheck` | Verify credential decrypts. |
| `Register-SchedulerCredential` | Save DPAPI cred. |

## MSP (Phase 6)

### `MSPReports.ps1`

| Function | Purpose |
|---|---|
| `Invoke-AcrossTenants -Script <sb>` | Run scriptblock per tenant; restore prior context. |
| `Get-CrossTenantLicenseReport` | Rollup. |
| `Get-CrossTenantMfaReport` | Rollup. |
| `Get-CrossTenantStaleGuests` | Rollup. |
| `Get-CrossTenantOrphanedTeams` | Rollup. |
| `Get-CrossTenantBreakGlass` | Rollup. |

### `MSPDashboard.ps1`

| Function | Purpose |
|---|---|
| `Update-MSPDashboard` | Render portfolio HTML. |

## AI (Phase 5)

### `AIAssistant.ps1`

| Function | Purpose |
|---|---|
| `Start-AIAssistant` | The chat REPL (menu slot 99). |
| `Invoke-AIChat` | Single-turn provider call. |
| `Convert-ToSafePayload` / `Restore-FromSafePayload` | PII redaction. |
| `Reset-PrivacyMap` | Drop the session token map. |
| `Get-AIConfig` / `Save-AIConfig` | Config CRUD. |

### `AIToolDispatch.ps1`

| Function | Purpose |
|---|---|
| `Get-AIToolCatalog` | Load `ai-tools/*.json`. |
| `Test-AIToolInput` | JSON-Schema validate. |
| `Invoke-AIToolCall -ToolName <n> -Params <h>` | Dispatch. |

### `AIPlanner.ps1`

| Function | Purpose |
|---|---|
| `Invoke-AIPlanApprovalFlow -PlanInput <obj>` | Show plan + approve + execute. |
| `Set-AIPlanMode -Mode auto\|force\|skip` | Plan-mode latch for next prompt. |

### `AISessionStore.ps1`

| Function | Purpose |
|---|---|
| `Save-AISession -Title <t>` | Persist DPAPI-encrypted. |
| `Load-AISession -IdOrPrefix <id>` | Resume. |
| `Get-AISessionList` | Enumerate. |
| `Remove-AISession` / `Rename-AISession` | CRUD. |
| `Export-AISession -Id <id> -Path <p>` | Redacted JSON for sharing. |

### `AICostTracker.ps1`

| Function | Purpose |
|---|---|
| `Add-AICostEvent -Provider -Model -Usage` | Record one call. |
| `Get-AICostSummary` | Current session. |
| `Get-AICostHistory -Days N` | Rollup. |

## Incident response (Phase 7)

### `IncidentResponse.ps1`

| Function | Purpose |
|---|---|
| `Invoke-CompromisedAccountResponse -UPN -Severity` | The 13-step playbook. |
| `Get-IncidentSteps -Severity` | The step plan for a given severity. |
| `New-IncidentId` | `INC-YYYY-MM-DD-xxxx` generator. |
| `Get-IncidentDirectory -IncidentId <id>` | Per-tenant artifact dir. |

### `IncidentRegistry.ps1`

| Function | Purpose |
|---|---|
| `Get-Incident -Id <id>` | Fold all JSONL records for one id. |
| `Get-Incidents [-Status] [-Days] [-Severity]` | Filter list. |
| `Show-Incidents` | Compact table. |
| `Show-IncidentReport -Id <id>` | Open report.html. |
| `Close-Incident -Id -Resolution [-FalsePositive]` | Mark closed; optional undo walk. |
| `Undo-Incident -Id` | Walk reversible steps with confirmation. |
| `Export-Incident -Id -Path` | Compliance ZIP bundle. |
| `Get-IncidentTimeline` / `Get-IncidentList` / `Summarize-AuditEvents` | AI tool wrappers. |

### `IncidentBulk.ps1`

| Function | Purpose |
|---|---|
| `Invoke-BulkIncidentResponse -Path <csv>` | Multi-account incident response. |
| `Invoke-IncidentTabletop -ScenarioName <n>` | IR-readiness exercise. |
| `Get-TabletopScenarios` | Enumerate scenarios. |

### `IncidentTriggers.ps1`

| Function | Purpose |
|---|---|
| `Detect-AnomalousLocationSignIn -UPN` | Auto-trigger detector. |
| `Detect-ImpossibleTravel -UPN` | |
| `Detect-HighRiskSignIn -UPN` | |
| `Detect-MassFileDownload -UPN` | |
| `Detect-MassExternalShare -UPN` | |
| `Detect-SuspiciousInboxRule -UPN` | |
| `Detect-MFAFatigue -UPN` | |
| `Invoke-IncidentDetectors -UPNs [-All]` | Driver — runs every detector against the UPN set. |

## See also

- [`menu-map.md`](menu-map.md) — menu → function mapping.
- [`tool-catalog.md`](tool-catalog.md) — AI-callable subset.
- [`undo-handlers.md`](undo-handlers.md) — reverse dispatch table.
- [`../developer/architecture.md`](../developer/architecture.md) — module dependencies + design intent.
