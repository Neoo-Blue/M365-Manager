# Repository relationships

M365-Manager exists across three repositories with strict, one-directional dependencies. This document explains the architecture, what lives where, and how to contribute.

## The three repos

| Repo | Visibility | Purpose |
|---|---|---|
| [`Neoo-Blue/M365-Manager`](https://github.com/Neoo-Blue/M365-Manager) (this one) | **Public** | The PowerShell tool. Modules, TUI, AI catalog, audit / undo / redaction primitives, tests, operator documentation. |
| [`Neoo-Blue/m365-manager-server`](https://github.com/Neoo-Blue/m365-manager-server) | Private | Integration layer. MCP server + Microsoft Teams bot + per-user permission scoping + escalation flow + workflow learning. |
| [`Neoo-Blue/m365-manager-cloud`](https://github.com/Neoo-Blue/m365-manager-cloud) | Private | SaaS control plane. Multi-tenant orchestration, customer portal, billing, observability, compliance controls. Currently a placeholder. |

## Dependency direction

```
+-------------------------------+
|  Neoo-Blue/M365-Manager       |   PowerShell module + TUI + docs
|  (public)                     |
+---------------+---------------+
                |
                | depends on (PSGallery / direct install)
                v
+-------------------------------+
|  Neoo-Blue/m365-manager-server|   MCP server + Teams bot
|  (private)                    |
+---------------+---------------+
                |
                | depends on (runtime + hosted-as)
                v
+-------------------------------+
|  Neoo-Blue/m365-manager-cloud |   SaaS control plane
|  (private, placeholder)       |
+-------------------------------+
```

**Strictly one-way down the stack.** The public repo never imports from the private repos. The server never imports from the cloud. Changes flow upstream: a bug fix in the public PowerShell module benefits everyone; a bug fix in the server only benefits hosted customers.

## What lives where

### `M365-Manager` (this repo, public)

- The PowerShell module: ~40 `.ps1` files dot-sourced via `Main.ps1`.
- The TUI: blue background, menus, confirmation prompts, mode banner.
- The AI tool catalog (`ai-tools/*.json`).
- The audit log + undo dispatch + PREVIEW mode + redaction layer.
- The 12-step canonical offboard flow.
- The 13-step compromised-account incident-response playbook.
- The 7-detector auto-triggers for incident response.
- Pester test suite.
- Comprehensive operator + contributor documentation under `docs/`.

If you're solving a problem at the PowerShell-or-tenant-administration layer, it goes here.

### `m365-manager-server` (private)

- The MCP server implementation (Python). Exposes the public repo's AI tool catalog as MCP tools.
- The Microsoft Teams bot (Bot Framework). Authenticates each user on-behalf-of, dispatches via MCP, replies with Adaptive Cards.
- The per-user permission catalog (`permissions/role-tool-map.json`).
- The escalation / approval queue (with Teams DM approver routing, time-bound approvals, full audit chain).
- The workflow / skill learning store (per-tenant, tokenized).
- Deployment artifacts (Docker images, Bicep templates) for single-tenant + multi-tenant Azure-hosted deployments.

If you're solving a problem at the integration layer — exposing the PowerShell module to non-PowerShell users via Teams or MCP-aware AI assistants — it goes here.

### `m365-manager-cloud` (private, placeholder)

- Multi-tenant orchestration of `m365-manager-server` instances.
- Customer signup + onboarding flow.
- Billing integration (Stripe).
- Customer-facing portal (non-PowerShell UI).
- Cross-tenant observability (metrics, logs, traces, alerting).
- SOC 2 / ISO 27001 controls + evidence collection.

If you're solving a problem at the "Neoo-Blue runs this for the customer" layer, it goes here. Currently nothing is built — the repo exists to reserve the shape.

## How a feature lands

Most features start in the public repo. The path:

1. **Need identified.** Either by Neoo-Blue, by a contributor, or by a customer.
2. **Implement in the public repo first.** A new tool category (e.g. "Power Platform admin"), a new AI tool, a new playbook — all of these belong in the public repo's PowerShell module.
3. **Surface in the server if needed.** If non-PowerShell users (Teams users, MCP-aware AI assistants) should have access, the feature appears in the server repo as a thin MCP / bot integration. The server doesn't reimplement; it dispatches.
4. **Roll out via the cloud if needed.** Hosted customers get the feature when the cloud control plane provisions an updated server-image.

There are exceptions:

- **MCP-specific concerns** (MCP protocol features, Bot Framework integration patterns, Adaptive Card design) live in the server repo from the start — they have no PowerShell analog.
- **SaaS-specific concerns** (billing, customer signup, cross-tenant dashboards) live in the cloud repo from the start.

## What gets backported

When a fix or feature lands in the private repos, sometimes it needs to backport to the public repo:

- **Backport when** the underlying issue exists in the public PowerShell module and would affect any user (operator, server, cloud).
- **Don't backport when** the fix is integration-layer specific (e.g. a Bot Framework adapter quirk, an MCP protocol edge case, a multi-tenant orchestration concern).

Example backport: while building the Teams bot, the contributor finds that `Invoke-CompromisedAccountResponse` has a subtle bug where the AI dispatcher's default branch bypasses `Invoke-Action` for some SDK cmdlets. That's a public-repo bug; backport. The fix file lives in the public repo and is consumed by the server.

Example NOT to backport: the bot needs an "approve / reject / modify" Adaptive Card. That's bot-specific; no analog in the PowerShell TUI; stays in the server repo.

## How to contribute

- **Public repo**: open pull requests against `Neoo-Blue/M365-Manager`. The contribution guide is at `docs/developer/` (under the reorganized docs structure landing in `chore/docs-reorganization`).
- **Server repo**: currently a closed-source private repo. Contribution policy will be established as the project matures. If you want to contribute and you've found this doc somehow, reach out via a GitHub issue on the public repo.
- **Cloud repo**: same as the server repo. Currently a placeholder.

## Why this split

The decision to keep the MCP server + Teams bot private (rather than open-source alongside the PowerShell module) is intentional:

1. **The PowerShell module is the operator's tool.** It's general-purpose; runs on any operator's workstation; benefits from open contribution.
2. **The server is an integration product.** It targets a smaller audience (IT shops + MSPs that want chat-based access to the module's surface). Closed-source preserves room for paid support / hosted offerings without competing with a community-maintained fork.
3. **The cloud is a service.** It's not source-distributable in any meaningful sense — the value is operating it, not the code.

Open-sourcing the public repo while keeping the integration + service layers private is a deliberate choice to balance contribution-friendly engineering practices with sustainable business model design.

## Versioning

Each repo versions independently using [Semantic Versioning](https://semver.org/spec/v2.0.0.html). Inter-repo compatibility is managed via the server repo's `minimumCoreVersion` field in its dependency manifest.

- Public repo: ships features on phase-based feature branches (`feature/<area>`).
- Server repo: ships features on a single feature branch per phase (`feature/phase-8-mcp`).
- Cloud repo: not actively developed.

## See also

- [`README.md`](README.md) — public repo project overview.
- [`docs/README.md`](docs/README.md) — full documentation tree.
- Server repo: https://github.com/Neoo-Blue/m365-manager-server
- Cloud repo: https://github.com/Neoo-Blue/m365-manager-cloud
