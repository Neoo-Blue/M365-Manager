# Config keys

Every config key the tool reads. Single sortable reference. The overview guide is at [`../getting-started/configuration.md`](../getting-started/configuration.md); this is the comprehensive list.

Scopes:
- **Global** — `ai_config.json` (per machine, per user).
- **Tenant** — `<stateDir>\tenant-overrides\<tenant>.json` (per tenant, overrides global).
- **Env** — `$env:M365MGR_<KEY>` (uppercase, dots → underscore).
- **CLI** — function parameter at call time.

Resolution order: `Global → Tenant → Env → CLI`. Last wins. Today only 3 keys are wired through the `Get-EffectiveConfig` resolver — the rest read `$Config` directly.

## Top-level

| Key | Type | Default | Scope | Notes |
|---|---|---|---|---|
| `Provider` | string | `"Anthropic"` | Global | One of: `Ollama` / `Anthropic` / `OpenAI` / `AzureOpenAI` / `Custom`. |
| `Endpoint` | string | provider-specific | Global | Full URL for chat completions. |
| `Model` | string | provider-specific | Global | Model name. |
| `ApiKey` | string | `""` | Global | DPAPI-encrypted on first save. |

## Privacy

| Key | Type | Default | Scope | Notes |
|---|---|---|---|---|
| `Privacy.ExternalRedaction` | `Enabled` / `Disabled` | `Enabled` | Global | Tokenize PII for non-local providers. |
| `Privacy.RedactInAuditLog` | `Enabled` / `Disabled` | `Disabled` | Global | When Enabled, the audit log uses tokens too. |
| `Privacy.ExternalPayloadCapBytes` | int | `8192` | Global | `0` = no cap. |
| `Privacy.TrustedProviders` | string[] | `[]` | Global | Lowercase provider names — treat like localhost. |

## Notifications

| Key | Type | Default | Scope | Notes |
|---|---|---|---|---|
| `Notifications.DefaultEmailFrom` | string | `""` | Global, Tenant | UPN to send from. Empty = operator's mailbox. |
| `Notifications.SecurityTeamRecipients` | string[] | `[]` | Global, Tenant | For `Critical`. |
| `Notifications.OperationsTeamRecipients` | string[] | `[]` | Global, Tenant | For `Warning` / `Info`. |
| `Notifications.TeamsWebhookSecurity` | string | `""` | Global, Tenant | DPAPI-encrypted at rest. |
| `Notifications.TeamsWebhookOperations` | string | `""` | Global, Tenant | DPAPI-encrypted. |
| `Notifications.DryRunNotifications` | bool | `false` | Global, Tenant | Log without sending. |

## AI assistant

| Key | Type | Default | Scope | Notes |
|---|---|---|---|---|
| `AI.MonthlyBudgetUsd` | double | `100.00` | Global, Tenant, Env | **Wired through `Get-EffectiveConfig`.** |
| `AI.AlertAtPct` | int | `80` | Global, Tenant, Env | **Wired.** Warn when monthly spend crosses this percentage. |
| `AI.AutoPlanThreshold` | int | `3` | Global | Auto-plan-mode kicks in at this many tool calls. |
| `AI.MaxTurns` | int | `25` | Global | Cap on conversation turns before forced summarization. |
| `AI.SessionRetentionDays` | int | `90` | Global | Auto-prune saved chats older than this. (Not yet wired; see follow-up.) |

## Incident response (Phase 7)

| Key | Type | Default | Scope | Notes |
|---|---|---|---|---|
| `IncidentResponse.AutoExecuteOnSeverity` | enum | `"None"` | Global, Tenant | `None` / `Critical` / `HighAndCritical`. **The most consequential key in the tool.** |
| `IncidentResponse.UseAIForNarrative` | enum | `"Disabled"` | Global, Tenant | `Enabled` allows `Summarize-AuditEvents` to send (redacted) audit data to the AI provider. |
| `IncidentResponse.SnapshotRetentionDays` | int | `365` | Global, Tenant | How long to keep incident artifacts on disk. |
| `IncidentResponse.DetectorIntervalMinutes` | int | `15` | Global, Tenant | Frequency the scheduler runs the detector sweep. |
| `IncidentResponse.ImpossibleTravelMaxKmPerHour` | int | `900` | Global, Tenant | Threshold for `Detect-ImpossibleTravel`. |
| `IncidentResponse.MassDownloadFileCount` | int | `50` | Global, Tenant | `Detect-MassFileDownload` threshold. |
| `IncidentResponse.MassDownloadWindowMinutes` | int | `5` | Global, Tenant | Window for above. |
| `IncidentResponse.MassShareCount` | int | `20` | Global, Tenant | `Detect-MassExternalShare` threshold. |
| `IncidentResponse.MassShareWindowMinutes` | int | `60` | Global, Tenant | Window for above. |
| `IncidentResponse.MFAFatigueRejectCount` | int | `10` | Global, Tenant | `Detect-MFAFatigue` threshold. |
| `IncidentResponse.MFAFatigueWindowMinutes` | int | `60` | Global, Tenant | Window for above. |
| `IncidentResponse.AnomalousLocationLookbackDays` | int | `90` | Global, Tenant | `Detect-AnomalousLocationSignIn` baseline window. |
| `IncidentResponse.TabletopUPN` | string | `""` | Global, Tenant | Sandbox account for `Invoke-IncidentTabletop`. |

## Other

| Key | Type | Default | Scope | Notes |
|---|---|---|---|---|
| `StaleGuestDays` | int | `90` | Global, Tenant, Env | **Wired through `Get-EffectiveConfig`.** Threshold for `Get-StaleGuests`. |
| `BreakGlassPasswordAgeWarnDays` | int | `90` | Global | Warn if break-glass password not rotated within this many days. |
| `BreakGlassSignInWarnDays` | int | `30` | Global | Warn if break-glass account signed in recently (suspicious). |

## Wired vs not

Today only `AI.MonthlyBudgetUsd`, `AI.AlertAtPct`, and `StaleGuestDays` route through the `Get-EffectiveConfig` resolver (which honors the full `Global → Tenant → Env → CLI` chain). Every other key in this list reads `$Config` directly from the global file — meaning tenant overrides + env vars are accepted by the resolver but ignored by the consuming module.

This is mechanical conversion work tracked in the deferred-items section of [`../operations/pre-merge-review.md`](../operations/pre-merge-review.md). The override file format is forward-compatible; the unwired keys will become wired without breaking change.

## See also

- [`../getting-started/configuration.md`](../getting-started/configuration.md) — overview + per-block reference.
- [`../concepts/multi-tenant.md`](../concepts/multi-tenant.md) — tenant override mechanics.
- [`template-schema.md`](template-schema.md) — template JSON schemas.
