# M365 Manager — Documentation

Comprehensive documentation for the M365 Manager tool. Organized into seven sections by audience and intent:

- **[getting-started/](getting-started/)** — installation, first-run, configuration. **Start here if you're new to the tool.**
- **[concepts/](concepts/)** — how the tool is architected. Security model, multi-tenant model, AI assistant model.
- **[guides/](guides/)** — task-oriented walkthroughs for each feature area (onboarding, offboarding, incident response, etc.).
- **[reference/](reference/)** — schema and format references (audit JSONL, CSV columns, AI tool catalog, undo handlers, config keys).
- **[operations/](operations/)** — running M365 Manager in production. Deployment, upgrades, troubleshooting, permissions.
- **[playbooks/](playbooks/)** — ready-to-execute IR-style runbooks for specific scenarios.
- **[developer/](developer/)** — contributing. Architecture deep-dive, adding modules / AI tools / detectors.
- **[samples/](samples/)** — example inputs and outputs (plans, sessions, incidents, dashboards, audit lines).

## Quick links

| If you want to… | Read this |
|---|---|
| Get the tool running on your workstation today | [`getting-started/installation.md`](getting-started/installation.md) |
| Take a 10-minute guided tour | [`getting-started/quickstart.md`](getting-started/quickstart.md) |
| Understand the security model | [`concepts/security-model.md`](concepts/security-model.md) |
| Offboard a user | [`guides/offboarding.md`](guides/offboarding.md) |
| Respond to a compromised account | [`guides/incident-response.md`](guides/incident-response.md) |
| Run the same operation across many tenants | [`concepts/multi-tenant.md`](concepts/multi-tenant.md) |
| Talk to the AI assistant ("Mark") | [`concepts/ai-assistant.md`](concepts/ai-assistant.md) |
| Diagnose an error message | [`operations/troubleshooting.md`](operations/troubleshooting.md) |
| Roll out the tool to your team | [`operations/deployment.md`](operations/deployment.md) |
| Add a new feature module | [`developer/adding-a-module.md`](developer/adding-a-module.md) |

## What's where

### getting-started/
- [`installation.md`](getting-started/installation.md) — prereqs (PowerShell 7+, Graph SDK + EXO + SPO modules, gh CLI), install steps, permissions matrix, first-run.
- [`quickstart.md`](getting-started/quickstart.md) — 10-minute hands-on tour. Connect, register a tenant, run one report, run one PREVIEW-mode mutation.
- [`tenant-setup.md`](getting-started/tenant-setup.md) — connecting your first tenant, registering a profile (Phase 6), partner / GDAP flow.
- [`configuration.md`](getting-started/configuration.md) — every config key, what it does, default, scope (global / tenant override / env / CLI).

### concepts/
- [`architecture.md`](concepts/architecture.md) — module map, dot-source order, how `Invoke-Action` threads through every mutation.
- [`security-model.md`](concepts/security-model.md) — DPAPI key encryption, AST allow-list, redaction, audit log, undo.
- [`multi-tenant.md`](concepts/multi-tenant.md) — tenant profile registry, switching, MSP mode (Phase 6).
- [`tenant-overrides.md`](concepts/tenant-overrides.md) — per-tenant config keys, the `Get-EffectiveConfig` resolver.
- [`ai-assistant.md`](concepts/ai-assistant.md) — Mark conceptually: tools, plans, costs, sessions.

### guides/
Task-oriented walkthroughs. One per feature area.
- [`onboarding.md`](guides/onboarding.md) · [`offboarding.md`](guides/offboarding.md) · [`incident-response.md`](guides/incident-response.md)
- [`license-optimization.md`](guides/license-optimization.md) · [`mfa-management.md`](guides/mfa-management.md) · [`sharepoint-management.md`](guides/sharepoint-management.md) · [`teams-management.md`](guides/teams-management.md) · [`onedrive-handoff.md`](guides/onedrive-handoff.md) · [`guest-lifecycle.md`](guides/guest-lifecycle.md)
- [`audit-and-undo.md`](guides/audit-and-undo.md) · [`sign-in-lookup.md`](guides/sign-in-lookup.md) · [`unified-audit-log.md`](guides/unified-audit-log.md)
- [`breakglass-accounts.md`](guides/breakglass-accounts.md) · [`scheduled-checks.md`](guides/scheduled-checks.md) · [`notifications.md`](guides/notifications.md) · [`tabletop-exercises.md`](guides/tabletop-exercises.md)
- AI: [`ai-tools-overview.md`](guides/ai-tools-overview.md) · [`ai-planning.md`](guides/ai-planning.md) · [`ai-sessions.md`](guides/ai-sessions.md) · [`ai-costs.md`](guides/ai-costs.md)
- MSP: [`msp-dashboard.md`](guides/msp-dashboard.md)

### reference/
- [`menu-map.md`](reference/menu-map.md) — every menu / submenu / slot number.
- [`chat-commands.md`](reference/chat-commands.md) — every AI chat slash command + behavior.
- [`csv-formats.md`](reference/csv-formats.md) — every bulk CSV schema in one place.
- [`template-schema.md`](reference/template-schema.md) — role templates, site templates, scheduled-check templates.
- [`audit-format.md`](reference/audit-format.md) — JSONL audit record reference.
- [`tool-catalog.md`](reference/tool-catalog.md) — AI tool catalog (every tool, signature, destructive/reverse flags).
- [`undo-handlers.md`](reference/undo-handlers.md) — reverse handler dispatch table.
- [`config-keys.md`](reference/config-keys.md) — every config key with scope + default.
- [`cmdlets.md`](reference/cmdlets.md) — public function reference, grouped by module.

### operations/
- [`deployment.md`](operations/deployment.md) — rolling out on operator workstations, exec policy, signed-module guidance.
- [`upgrade-guide.md`](operations/upgrade-guide.md) — upgrading between phases, breaking changes, migration steps.
- [`troubleshooting.md`](operations/troubleshooting.md) — common error messages → likely cause → fix.
- [`permissions.md`](operations/permissions.md) — Graph scopes + EXO/SPO/SCC role matrix per feature.
- [`validation-runbook.md`](operations/validation-runbook.md) — live PREVIEW-mode smoke for the destructive flows.
- [`pre-merge-review.md`](operations/pre-merge-review.md) — pre-merge self-critique from the v1 ship.

### playbooks/
Scenario-specific runbooks. Each is a full walkthrough — when does this apply, what to do, expected outcome.
- [`compromised-account.md`](playbooks/compromised-account.md) — paper companion to the incident-response tool.
- [`mass-phishing-response.md`](playbooks/mass-phishing-response.md)
- [`insider-departure.md`](playbooks/insider-departure.md)
- [`audit-prep.md`](playbooks/audit-prep.md)
- [`license-true-up.md`](playbooks/license-true-up.md)
- [`tenant-migration.md`](playbooks/tenant-migration.md)
- [`incident-triggers.md`](playbooks/incident-triggers.md) — auto-detection framework.

### developer/
- [`architecture.md`](developer/architecture.md) — deep dive into module dependencies + life of a request.
- [`adding-a-module.md`](developer/adding-a-module.md) — end-to-end walkthrough of adding a feature.
- [`adding-an-ai-tool.md`](developer/adding-an-ai-tool.md) — tool catalog JSON schema, reverse-tool pairing.
- [`adding-a-detector.md`](developer/adding-a-detector.md) — wiring a new incident-response trigger.
- [`testing.md`](developer/testing.md) — Pester patterns, mocking Graph, smoke checklist.
- [`plugin-api.md`](developer/plugin-api.md) — placeholder stub for the future plugin API.

### samples/
- [`incident/`](samples/incident/) — anonymized example incident folder (snapshot + audits + report).
- [`msp-dashboard/`](samples/msp-dashboard/) — pre-rendered HTML for a 3-tenant portfolio.
- [`health-output/`](samples/health-output/) — five scheduled-check result JSONs.
- [`sample-plan.json`](samples/sample-plan.json) · [`sample-session-export.json`](samples/sample-session-export.json) — AI assistant samples.
- [`preview-session.log`](samples/preview-session.log) — sample PREVIEW-mode audit JSONL.

## See also

- Top-level [`README.md`](../README.md) — project landing page.
- [`CHANGELOG.md`](../CHANGELOG.md) — *(future)* version-by-version delta.
