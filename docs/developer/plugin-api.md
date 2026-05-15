# Plugin API — placeholder

**Status: not implemented.** This doc reserves the design space + describes the contract a future plugin API would expose.

M365 Manager today is a single dot-sourced codebase. Adding new functionality means editing the repo + landing PRs against `main`. A plugin API would let third parties extend the tool without modifying the core repo — useful for:

- Per-customer customizations an MSP wants to keep private.
- Vendor integrations (PIM tools, SIEM forwarders, ticket-tracker bridges) that wouldn't belong in the core repo.
- Pre-release experimental features that the maintainer wants to ship under a flag.

This doc captures the design intent so a contributor implementing the plugin API has a target. It is not a binding contract — early implementation feedback may change it.

## Design goals

- **Drop-in extension.** Plugins live in their own dir; loaded by enumerating that dir; no edits to `Main.ps1`'s module list.
- **No core elevation.** Plugins run with the same trust the operator has; they can't escalate.
- **Auditable.** Plugin operations go through `Invoke-Action` just like core operations. No bypass.
- **Versioned.** Plugins declare a minimum core version they need + a manifest version.
- **Signed.** Plugins should be signable separately from the core. Production deploys can enforce `Trusted Publishers`.

## Proposed shape

### Plugin directory

```
<stateDir>/plugins/<plugin-name>/
├── manifest.json
├── <plugin-name>.ps1            # main entry
├── ai-tools/                    # optional: catalog additions
│   └── <name>.json
├── templates/                   # optional
│   └── ...
└── README.md
```

### manifest.json

```jsonc
{
  "name":                 "ServiceNowBridge",
  "version":              "0.4.0",
  "description":          "Open ServiceNow tickets from incident-response operations.",
  "author":               "Acme Corp",
  "minimumCoreVersion":   "1.0.0",
  "load": {
    "psFiles":            ["ServiceNowBridge.ps1"],
    "aiToolFiles":        ["ai-tools/servicenow.json"],
    "templateFiles":      []
  },
  "menuEntries": [
    { "submenu": "Incident Response", "label": "Open ServiceNow ticket", "function": "Start-SNOWTicketFlow" }
  ],
  "config": {
    "block": "ServiceNowBridge",
    "keys": [
      { "name": "InstanceUrl",    "type": "string", "default": "" },
      { "name": "ApiKey",         "type": "string", "default": "", "encrypted": true }
    ]
  },
  "permissions": ["Notifications.Send","AuditLog.Write","Network.External"]
}
```

### Loading

`Main.ps1` would, after dot-sourcing the core modules, walk `<stateDir>/plugins/*/manifest.json`, validate each, and load matched plugins:

```powershell
foreach ($manifest in (Get-ChildItem "<stateDir>/plugins/*/manifest.json")) {
    $m = Get-Content $manifest -Raw | ConvertFrom-Json
    if (-not (Test-PluginCompat -Manifest $m -CoreVersion $script:CoreVersion)) {
        Write-Warn "Plugin '$($m.name)' requires core $($m.minimumCoreVersion); skipped."
        continue
    }
    foreach ($psFile in $m.load.psFiles) {
        . (Join-Path (Split-Path $manifest) $psFile)
    }
    foreach ($toolFile in $m.load.aiToolFiles) {
        Register-AIToolFromFile -Path (Join-Path (Split-Path $manifest) $toolFile)
    }
    # ... etc
}
```

### Menu integration

Plugins can:

- Add a top-level slot (rare; reserved for substantial features).
- Add a submenu option (common).
- Extend an existing flow via a hook (e.g. "after offboard completes, call Start-SNOWTicketFlow").

The exact hook surface is TBD; candidates:

- `On-OffboardComplete <upn>`
- `On-IncidentClose <id>`
- `On-LicenseRemediation <upn> <sku>`

### AI tool integration

Plugins can register catalog entries by dropping JSON files in their `ai-tools/` subdir. The loader merges them into the global catalog.

The same destructive / reverse / explicit-approval rules apply — see [`adding-an-ai-tool.md`](adding-an-ai-tool.md).

### Config integration

Plugins declare config keys in their manifest. The loader registers them with `Get-EffectiveConfig` so the standard resolution order (global → tenant → env → CLI) applies. Encrypted keys (`"encrypted": true`) go through `Protect-Secret`.

### Permissions

Plugins declare what they need:

| Permission | Grants access to |
|---|---|
| `Notifications.Send` | Calling `Send-Notification`. |
| `AuditLog.Write` | Writing audit entries (most plugins want this). |
| `AuditLog.Read` | Reading audit entries (analytics plugins). |
| `Network.External` | Making outbound HTTP calls (3rd-party integrations). |
| `State.Write` | Writing to `<stateDir>` outside the plugin's own subdir. |
| `Tenant.Switch` | Calling `Switch-Tenant`. Most plugins shouldn't need. |

The core verifies permissions at load time + refuses to register plugins requesting permissions they didn't declare.

## Backwards compatibility

Adding the plugin API is additive — existing modules under the repo root continue to work. The plugin directory is opt-in; tenants that don't use plugins see no change.

## Not yet implemented

The above is the **design**. Today, the only way to extend M365 Manager is via PRs against the core repo. If you're a contributor reading this and you'd like to drive the plugin API:

1. File an issue with your use case (what plugin would you ship?). Concrete uses guide the design.
2. Implement against the proposed shape above. Or argue why the shape should change.
3. Land a minimum-viable plugin loader (no menu hooks, no AI integration — just dot-source PS files from `<stateDir>/plugins/`) so the API can ship incrementally.

## See also

- [`architecture.md`](architecture.md) — core architecture.
- [`adding-a-module.md`](adding-a-module.md) — current contribution pattern (in-tree).
- [`adding-an-ai-tool.md`](adding-an-ai-tool.md) — current AI tool addition (in-tree).
- [`testing.md`](testing.md) — Pester patterns the plugin API would inherit.
