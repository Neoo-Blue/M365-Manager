# Upgrade guide

Version-to-version migration. M365 Manager has shipped through 7 phases. Most upgrades are `git pull` and run — this doc calls out the few where format changes or operator action is needed.

## General upgrade procedure

```powershell
# 1. Stop the tool (close every operator's M365 Manager process).
# 2. Pull or repackage:
cd C:\Tools\M365Manager
git pull --ff-only

# 3. Sanity check on one workstation:
Invoke-Pester ./tests/                     # 224/224 expected
.\Launch.bat                               # menu opens, tenant connects

# 4. Roll out to the fleet (re-deploy the share / package).
```

The audit log format, tenant-profile format, AI session format, and incident artifact format are all forward-compatible — meaning a v1.6 file is readable by v1.7 of the tool. Backward-compat is NOT a goal: don't expect a v1.7 file to be readable by older versions.

## Phase-by-phase upgrade notes

### Phase 0 → Phase 0.5 (security hardening)

- `ai_config.json`'s `ApiKey` field becomes DPAPI-encrypted on first read. Plaintext keys get encrypted in place; **the original plaintext is unrecoverable** after the migration. Operators should save their plaintext key elsewhere before launching v0.5 the first time.
- New `Privacy` config block. Defaults: `ExternalRedaction=Enabled`, `RedactInAuditLog=Disabled`. No operator action required unless you want non-default behavior.

### Phase 0.5 → Phase 1 (bulk + templates + dryrun)

- New menu slots 13 / 14 (Bulk Onboard / Bulk Offboard).
- PREVIEW mode picker on launch. Existing scripts that call `Start-Onboard` etc. work unchanged — PREVIEW is an opt-in flip via `Set-PreviewMode -Enabled $true`.
- `templates/role-*.json` files ship. Drop new ones to match your roles.

### Phase 1 → Phase 2 (audit + MFA + lookups)

- **Audit log format changes from per-line text to JSONL.** Old logs (`session-<ts>.log` with `[ts] [event] [MODE=X] detail` lines) are still readable by `ConvertFrom-AuditLine` — but new entries are JSON. Mixed logs work; downstream readers should handle both.
- New `<stateDir>\undo-state.json` sidecar. Created on first `Invoke-Undo`.
- MFA module + sign-in lookup + UAL + audit viewer all land — no operator migration needed.

### Phase 2 → Phase 3 (lifecycle completeness)

- OneDrive / Teams / SharePoint / Guest modules land.
- Bulk-offboard CSVs gain new optional columns (`HandoffOneDriveTo`, `ConvertToShared`, etc.). Old CSVs (without these columns) still parse — defaults applied.
- New menu slots 16 (Teams), 17 (SharePoint), 18 (Guests).

### Phase 3 → Phase 4 (cost + health)

- New menu slot 19 (License optimizer) + 20 (Scheduled checks) + Break-glass + Notifications.
- `ai_config.json` gains `Notifications` block. Defaults are empty arrays; nothing fires until configured.
- Scheduler creates Windows Task Scheduler entries under `\M365Manager\` task path. Cleanup on uninstall: `Get-ScheduledTask -TaskPath '\M365Manager\' | Unregister-ScheduledTask -Confirm:$false`.

### Phase 4 → Phase 5 (AI v2)

- The AI assistant was previously a regex-driven `RUN:` extractor. Phase 5 introduces:
  - Native tool calling (`ai-tools/*.json` catalog).
  - Multi-step plans (`submit_plan` meta-tool).
  - Persistent encrypted sessions (`<stateDir>\chat-sessions\`).
  - Cost tracking (`<stateDir>\ai-costs.jsonl`).
- **The regex path is preserved as a fallback** for Ollama models that don't support tool calling. It emits a one-time deprecation warning when fired. No operator migration needed.
- New chat commands: `/plan`, `/noplan`, `/save`, `/load`, `/list`, `/rename`, `/delete`, `/ephemeral`, `/export`, `/cost`, `/costs`, `/about`, `/dryrun`, `/tools`. Old commands (`/clear`, `/config`, `/privacy`) preserved.

### Phase 5 → Phase 6 (multi-tenant)

- **Audit log filename changes** from `session-<ts>-<pid>.log` to `session-<ts>-<pid>-<tenant-slug>.log`. Old log files are not renamed — they keep their old names. Audit viewer reads both.
- **Tenant field in audit JSONL changes** from a string to a structured `{name, id, domain, mode}` hashtable. The audit-viewer filter handles both shapes via fallback.
- **New `<stateDir>\tenants.json` + `<stateDir>\secrets\tenant-<name>.dat`.** Created lazily on `Register-Tenant`.
- **`Select-TenantMode` now calls `Reset-AuditLogPath`** so the new tenant gets a new log file. No operator action.
- New chat command `/tenant <name>` and menu slot 21 (Tenants).
- `tenant-overrides/` directory is recognized; only 3 keys (`AI.MonthlyBudgetUsd`, `AI.AlertAtPct`, `StaleGuestDays`) actually route through `Get-EffectiveConfig` today. See [`pre-merge-review.md`](pre-merge-review.md).

### Phase 6 → Phase 7 (incident response)

- New `<stateDir>\<tenant-slug>\incidents\` directory structure. Tenant-scoped from the start.
- New menu slot 22 (Incident Response).
- New chat command `/incident <upn> [severity]`.
- New `IncidentResponse` config block in `ai_config.json`. Default `AutoExecuteOnSeverity=None` — no auto-escalation. **Operators should not raise this without a documented review.**
- New scheduled-check `health-incident-triggers.ps1`. Register it via the scheduler menu if you want the 15-minute detector sweep.
- `templates/tabletop-scenarios/` ships four scenarios. Add more as needed.

## Backwards-compatibility contract

Things we promise to keep working:

- **Audit log readability.** Old JSONL lines + pre-Phase-2 text lines remain readable by `Read-AuditEntries`.
- **CSV column names.** New columns may be added; old required columns stay.
- **Function signatures.** Public functions (the ones in [`../reference/cmdlets.md`](../reference/cmdlets.md)) keep their parameter names. New parameters get defaults so old call sites work.
- **Config key names.** Keys may move between blocks via `Get-EffectiveConfig` resolution, but the JSON paths stay valid.

Things we DON'T promise:

- **Forward-only artifact reading.** A new tool version may write artifacts that older tools can't read.
- **Module names.** Files may be renamed or split. Always launch via `Launch.bat`, never reference specific `.ps1` paths from external scripts.
- **Internal scriptblock signatures.** Functions starting with a verb that isn't in [`cmdlets.md`](../reference/cmdlets.md) are internal — they may change without notice.

## Breaking-change log

Maintained as a placeholder until v1.0 reaches a wider audience. Material breaking changes will land in `CHANGELOG.md` at the repo root.

| Date | Phase | Change | Migration |
|---|---|---|---|
| (none yet) | — | First public release. | — |

## See also

- [`deployment.md`](deployment.md) — fleet rollout patterns.
- [`validation-runbook.md`](validation-runbook.md) — live PREVIEW smoke for verifying an upgrade against a tenant.
- [`pre-merge-review.md`](pre-merge-review.md) — known deferred items.
