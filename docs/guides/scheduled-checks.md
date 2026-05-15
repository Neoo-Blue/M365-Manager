# Scheduled health checks

`Scheduler.ps1` registers a M365 Manager health-check script with the host scheduler (Windows Task Scheduler on Windows, `crontab` on macOS / Linux), then keeps a local index at `<stateDir>/scheduled-checks.json` so the operator can list / test / remove entries through the in-app menu without re-typing the schedule.

## The non-interactive contract

Every script under `health-checks/` accepts `-NonInteractive`. The scheduler always passes it. `UI.ps1`'s `Set-NonInteractiveMode -Enabled $true` flips a session flag that:

- `Read-UserInput` returns the empty string.
- `Confirm-Action` returns `$false` **immediately** — a scheduled run must NEVER silently approve a destructive prompt.
- `Show-Menu` returns `-1` so a menu-wrapped flow falls back to its "Cancel" / "Back" branch.
- `Pause-ForUser` becomes a no-op.

If you write a new check, route every prompt through `Get-OperatorInput` and never call `Read-Host` directly.

## Credential model

A scheduled run can't trigger an interactive Connect-MgGraph browser flow. The operator runs `Register-SchedulerCredential` once (interactively) and the module stores `{ TenantId, AppId, encryptedSecret }` at `<stateDir>/scheduler-cred.xml`. `encryptedSecret` uses the same DPAPI / B64 wrapper as the AI API key (`Protect-Secret`).

Scheduled scripts that need Graph should:

1. Call `Get-SchedulerCredential` to retrieve the bundle.
2. `Connect-MgGraph -ClientSecretCredential` (or cert-based equivalent).

The shipped Phase 4 health checks all use `Connect-ForTask` which honors a pre-existing Graph context, so they work transparently once the operator has registered the credential.

## Schedule syntax

`New-ScheduledHealthCheck -Schedule <s>` accepts:

| Friendly form          | Effect |
|------------------------|--------|
| `Daily 09:00`          | Every day at 09:00 local. |
| `Weekly Mon 09:00`     | Every Monday at 09:00. Days: Mon, Tue, Wed, Thu, Fri, Sat, Sun. |
| `Monthly 1 09:00`      | The 1st of every month at 09:00. |
| `Hourly`               | Every hour on the hour. |
| `cron 0 9 * * *`       | Raw 5-field cron expression, used as-is on POSIX hosts. |

On Windows the parser maps each form to `New-ScheduledTaskTrigger`. On macOS / Linux the parser maps to a 5-field cron expression and edits the user's crontab; each entry is tagged with a `# m365mgr-<Name>` marker so `Remove-ScheduledHealthCheck` can find and remove it.

## Shipped health checks

Each emits a structured `health-result-<name>-<ts>.json` to the audit directory.

| Script                                       | What it checks |
|----------------------------------------------|----------------|
| `health-license-usage.ps1`                   | Runs Phase 4's `Get-LicenseUtilizationReport` (default 60-day inactivity threshold). |
| `health-mfa-gaps.ps1`                        | `Get-UsersWithNoMfa` + `Get-UsersWithOnlyPhoneMfa` (Phase 2). |
| `health-stale-guests.ps1`                    | `Get-StaleGuests -DaysSinceSignIn 90` (Phase 3). |
| `health-orphaned-teams.ps1`                  | `Get-OrphanedTeams` + `Get-SingleOwnerTeams`. |
| `health-conditional-access-conflicts.ps1`    | Three heuristics over `/identity/conditionalAccess/policies`: disabled critical policies, all-users-with-large-exclusion, missing legacy-auth block. |
| `health-breakglass-signins.ps1`              | Pulls every registered break-glass account's last-24-hour sign-ins (Phase 4 Commit C). Normal state is zero rows. |

Sample outputs for each are under `docs/samples/health-output/`.

## Default behavior

- **Output channel**: file. Add `email` or `teams` to also route via the Notifications framework (Commit D); recipient routing follows severity (Critical -> security team, Warning -> security + ops, Info -> ops).
- **Notify on**: findings — emails only fire when the check actually has something to report. Toggle to `always` for a heartbeat or `failure` to alert only on the check itself erroring.
