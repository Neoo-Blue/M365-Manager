# Configuration

Every config key the tool reads, what it does, the default, and where it lives.

## Config files

| File | Scope | Encrypted at rest? | Gitignored? |
|---|---|---|---|
| `ai_config.json` | Global (per machine, per user) | Mixed — API keys + webhook URLs DPAPI-encrypted; everything else plaintext | YES |
| `ai_config.example.json` | Template | n/a | NO |
| `<stateDir>\tenants.json` | Per machine (registry of tenants) | Metadata plaintext; tenant secrets in sidecar `.dat` files DPAPI-encrypted | n/a |
| `<stateDir>\tenant-overrides\<tenant>.json` | Per tenant | Plaintext | n/a |

`<stateDir>` resolves to `%LOCALAPPDATA%\M365Manager\state` on Windows or `~/.m365manager/state` (chmod 700) on POSIX.

## Resolution order

For any single key, the effective value is resolved through (last wins):

1. **Global** — `ai_config.json`
2. **Tenant override** — `<stateDir>\tenant-overrides\<current-tenant>.json`
3. **Environment variable** — `$env:M365MGR_<KEY>` (uppercase, dots → underscore)
4. **CLI flag** — passed at function call time

Use `Get-EffectiveConfig -Key 'AI.MonthlyBudgetUsd'` to see what the resolver returns for the current tenant.

Today, 3 of 13 catalog'd keys actually route through `Get-EffectiveConfig` (`AI.MonthlyBudgetUsd`, `AI.AlertAtPct`, `StaleGuestDays`); the rest of the codebase still reads `$Config` directly. The resolver is built; mechanical conversion is follow-up. See [`../operations/pre-merge-review.md`](../operations/pre-merge-review.md) for the deferred items.

## Top-level keys (ai_config.json)

```jsonc
{
  "Provider": "Anthropic",          // Ollama | Anthropic | OpenAI | AzureOpenAI | Custom
  "Endpoint": "https://api.anthropic.com/v1/messages",
  "Model":    "claude-sonnet-4-20250514",
  "ApiKey":   "DPAPI:..."            // DPAPI-encrypted after first save
}
```

API keys are encrypted DPAPI on first save. If you paste a plaintext key into the JSON file manually and launch the AI assistant, the first read encrypts it in place. **The encrypted form is not portable** across users or machines.

## Privacy block

```jsonc
"Privacy": {
  "ExternalRedaction":       "Enabled",   // Enabled | Disabled
  "RedactInAuditLog":        "Disabled",  // Enabled | Disabled
  "ExternalPayloadCapBytes": 8192,        // 0 = no cap
  "TrustedProviders":        []           // ["azure-openai", ...] -- skip PII tokenization
}
```

| Key | Default | Notes |
|---|---|---|
| `ExternalRedaction` | `Enabled` | Tokenize UPNs / GUIDs / tenant IDs / display names before sending to non-local providers. Set Disabled only if you have a redaction layer upstream. **Secrets (JWT / `sk-…` / cert thumbprints) are always tokenized regardless.** |
| `RedactInAuditLog` | `Disabled` | Disabled writes raw values to the audit log (better for forensics). Enabled tokenizes audit values too. Either way, secret-bearing params are scrubbed. |
| `ExternalPayloadCapBytes` | `8192` | After redaction, outbound message content truncated past this byte count. Applies to external providers only. |
| `TrustedProviders` | `[]` | Lowercase provider names treated like localhost (no PII redaction; secrets still scrubbed). Example: `["azure-openai"]` for an in-tenant Azure deployment. |

## IncidentResponse block (Phase 7)

```jsonc
"IncidentResponse": {
  "AutoExecuteOnSeverity":         "None",      // None | Critical | HighAndCritical
  "UseAIForNarrative":             "Disabled",
  "SnapshotRetentionDays":         365,
  "DetectorIntervalMinutes":       15,
  "ImpossibleTravelMaxKmPerHour":  900,
  "MassDownloadFileCount":         50,
  "MassDownloadWindowMinutes":     5,
  "MassShareCount":                20,
  "MassShareWindowMinutes":        60,
  "MFAFatigueRejectCount":         10,
  "MFAFatigueWindowMinutes":       60,
  "AnomalousLocationLookbackDays": 90,
  "TabletopUPN":                   ""
}
```

`AutoExecuteOnSeverity` is the single most consequential key in the tool. Default `None` — every finding opens a forensic incident + alerts the team, but escalation to containment is operator-driven. See [`../playbooks/incident-triggers.md`](../playbooks/incident-triggers.md) for the trust model + tuning guide.

## Notifications block

```jsonc
"Notifications": {
  "DefaultEmailFrom":         "",
  "SecurityTeamRecipients":   [],
  "OperationsTeamRecipients": [],
  "TeamsWebhookSecurity":     "",       // DPAPI:... after first save
  "TeamsWebhookOperations":   "",       // DPAPI:... after first save
  "DryRunNotifications":      false
}
```

| Key | Default | Notes |
|---|---|---|
| `DefaultEmailFrom` | `""` (use the operator's mailbox) | Sender for `Send-Email`. |
| `SecurityTeamRecipients` | `[]` | Routed to for `-Severity Critical` notifications + incident-response alerts. |
| `OperationsTeamRecipients` | `[]` | Routed to for `-Severity Warning` / `Info` notifications. |
| `TeamsWebhookSecurity` | `""` | Posts Critical alerts. DPAPI-encrypted after first save. |
| `TeamsWebhookOperations` | `""` | Posts Warning / Info alerts. DPAPI-encrypted after first save. |
| `DryRunNotifications` | `false` | When true, every send writes an audit line but doesn't actually deliver. Useful for testing health checks. |

## Per-tenant overrides

Override any key for a specific tenant by writing `<stateDir>\tenant-overrides\<tenant-slug>.json`:

```jsonc
// E.g. <stateDir>\tenant-overrides\research.json
{
  "IncidentResponse": {
    "MassDownloadFileCount":     500,     // researchers pull large datasets routinely
    "MassDownloadWindowMinutes": 60
  },
  "AI": {
    "MonthlyBudgetUsd": 50.00
  }
}
```

Resolution rule: any key absent here falls back to the global value in `ai_config.json`.

## Environment variables

A few keys are also readable as env vars (uppercase, dots → underscore):

```powershell
$env:M365MGR_AI_MONTHLYBUDGETUSD = "100"
$env:M365MGR_STALEGUESTDAYS      = "120"
```

Env vars override tenant files but lose to explicit CLI flags.

## Catalog'd overridable keys

`$script:TenantOverridableKeys` in `TenantOverrides.ps1` lists the 13 keys the resolver knows about. Adding a new override-able key is a two-line change in that file. See [`../developer/adding-a-module.md`](../developer/adding-a-module.md) for the pattern.

## See also

- [`../reference/config-keys.md`](../reference/config-keys.md) — every config key in a single table, sortable.
- [`../concepts/security-model.md`](../concepts/security-model.md) — DPAPI key encryption details.
- [`../guides/notifications.md`](../guides/notifications.md) — Notifications setup walkthrough.
