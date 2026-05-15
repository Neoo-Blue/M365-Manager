# MSP portfolio dashboard (Phase 6)

A single-page HTML dashboard summarizing the posture of every
registered tenant. Self-contained -- no external CSS / JS / CDNs,
so the file is safe to email or open from a USB stick.

## Generating

```powershell
Update-MSPDashboard                            # all registered tenants
Update-MSPDashboard -Tenants @('Acme','Contoso')   # subset
```

Or from the main menu: `Tenants...` -> `MSP portfolio dashboard`.

Each refresh writes:
- `<stateDir>/msp-dashboard/msp-dashboard-<yyyymmdd-hhmmss>.html`
  (historical snapshot)
- `<stateDir>/msp-dashboard/msp-dashboard-latest.html` (always
  points at the most recent run)

Bookmark the `latest.html` and the dashboard updates in place.

## What each card shows

One card per tenant:

```
+-----------------------------+
| * Contoso          abc-123  |   <- color dot = break-glass posture
+-----------------------------+
| 312 users      $1,840 /mo   |
| 96.4% MFA      7 stale gst  |
| 2 orphan teams ok break-gl  |
+-----------------------------+
| last sync: 2026-05-12 18:33 |
+-----------------------------+
```

Metric sources:
- **users** -- `Get-MgUser -All -Property Id | Count`
- **monthly USD** -- sum of `MonthlyCostUsd` from Phase 4
  `Get-LicenseUtilizationReport`. Picks up tenant-override
  `LicensePrices` when set.
- **MFA %** -- `1 - (no-mfa-count / total-users)` from Phase 2
  `Invoke-MfaComplianceScan`. Color thresholds: green >= 95%,
  yellow >= 80%, red below.
- **stale guests** -- Phase 3 `Get-StaleGuests` (honors
  `StaleGuestDays` tenant override).
- **orphan teams** -- Phase 3 `Get-OrphanedTeams`.
- **break-glass posture** -- Phase 4 `Test-BreakGlassPosture`.
  Status dot: green (`ok`), yellow (`warn`), red (`fail` / error
  during refresh), gray (`unknown`).

## Portfolio totals at the top

- Total users across all tenants.
- Total monthly USD.
- User-weighted MFA compliance %.

## Refresh cadence guidance

The dashboard pulls live data via the same Phase 2-4 functions
under the hood, so each refresh round-trips Graph (and EXO when
applicable) for every tenant. Realistic timings on a desktop
broadband link:

| Tenant size       | Time per tenant   |
|-------------------|-------------------|
| <100 users        | 2-4 sec           |
| 100-1,000 users   | 5-15 sec          |
| 1,000-10,000      | 20-60 sec         |

For a 25-tenant portfolio expect 5-15 minutes for a full refresh.
Run on a schedule (Phase 4 `Scheduler.ps1` supports cron-style
recurrence) so the operator opens an already-fresh dashboard
rather than waiting for the refresh interactively.

## Performance / scale notes

- **Sequential by design** -- see [multi-tenant.md](../concepts/multi-tenant.md)
  for why parallelism is unsafe with SDK singletons.
- **Failed tenants get rendered as red cards** rather than
  silently dropping. The Error field is captured but not currently
  shown in the HTML; check the audit log
  (`actionType=CrossTenantStepError`) for the stack.
- **No JavaScript at all** -- the dashboard is pure static HTML
  / inline CSS. Opens in any browser, including offline. This
  is a deliberate choice: an MSP carrying a dashboard around on a
  laptop should not depend on a CDN.
- **No deep-link into per-tenant detail reports yet.** The brief
  mentioned click-through into existing per-tenant HTML reports
  (Phase 2-4 outputs). Those reports don't have stable on-disk
  paths today; adding the deep-links is tracked as a follow-up.
  The portfolio overview is the immediate win.

## Audit fingerprint

Every refresh writes one `ActionType=MSPDashboardRefresh` entry
to the audit log of the originating tenant:

```json
{
  "event": "MSPDashboardRefresh",
  "actionType": "MSPDashboardRefresh",
  "target": { "tenantCount": 7, "totalSpendUsd": 12840.20 },
  "result": "ok",
  "tenant": { "name": "...originating...", "id": "..." }
}
```

Per-tenant fetches inside the run land under each *target*
tenant's audit log because `Switch-Tenant` calls
`Reset-AuditLogPath`. To reconstruct a refresh end-to-end, grep
all `<stateDir>/audit/*-*.log` files for the same MSPDashboardRefresh
entryId.
