# M365 Administration Tool

A modular PowerShell TUI for Microsoft 365 administration. Blue background, multi-color menus, confirmation prompts on every change, browser-based OAuth, full audit log, undo system, AI assistant, multi-tenant / MSP mode, and a compromised-account incident-response playbook.

```text
+================================================================+
|  M365 Admin                       [LIVE]                       |
+================================================================+

  Tenant: GDAP Contoso (contoso.onmicrosoft.com)
    Graph: OK     EXO: OK     SCC: OK

  [ LIVE MODE -- changes apply to the tenant ]

   1.  Onboard New User                12. Bulk Onboard from CSV...
   2.  Offboard User                   13. Bulk Offboard from CSV...
   3.  Add / Remove License            14. Audit & Reporting...
   4.  Mailbox Archiving               15. MFA & Authentication...
   ...                                 22. Incident Response...
```

## Install in one line

```powershell
git clone https://github.com/Neoo-Blue/M365-Manager.git ; cd M365-Manager ; .\Launch.bat
```

Full prerequisites, module install, and first-run walkthrough: **[`docs/getting-started/installation.md`](docs/getting-started/installation.md)**

## Take a 10-minute tour

**[`docs/getting-started/quickstart.md`](docs/getting-started/quickstart.md)** — connect, register a tenant profile, run one report, run one PREVIEW-mode mutation, view the audit log. Verbatim-runnable end to end.

## What it does

| Area | Key flows |
|---|---|
| **Identity lifecycle** | Onboard (single + bulk + role templates) · Offboard (12-step canonical flow + bulk) · Group / DL / shared-mailbox / calendar / user-profile management |
| **Security** | MFA management + TAP issuance · Sign-in lookup · Unified audit log search · Break-glass account registry · Per-tenant audit log with undo system |
| **Lifecycle** | OneDrive handoff · Teams ownership transfer · SharePoint share cleanup · Guest user lifecycle with recertification |
| **Cost & health** | License optimizer with savings math · Scheduled health checks via Task Scheduler · Notifications (email + Teams webhook) |
| **MSP mode** | Tenant profile registry · Fast switch with per-tenant audit log · Cross-tenant rollups · Single-page HTML portfolio dashboard |
| **AI assistant** | Native tool calling · Multi-step plan approval · Persistent encrypted chat sessions · Cost tracking · PII redaction |
| **Incident response** | 13-step compromised-account playbook · Severity gating · Forensic snapshot before mutation · Auto-detection framework with 7 triggers · Bulk + tabletop modes |

Each row is its own guide under **[`docs/guides/`](docs/guides/)**.

## Documentation map

| Section | What's there |
|---|---|
| **[getting-started/](docs/getting-started/)** | Install, quickstart, tenant setup, configuration |
| **[concepts/](docs/concepts/)** | Architecture, security model, multi-tenant, AI assistant |
| **[guides/](docs/guides/)** | One walkthrough per feature area (16 docs) |
| **[reference/](docs/reference/)** | Audit format, CSV schemas, AI tool catalog, config keys, public function reference |
| **[operations/](docs/operations/)** | Deployment, upgrade, troubleshooting, permissions, validation runbook |
| **[playbooks/](docs/playbooks/)** | Ready-to-execute IR-style runbooks for specific scenarios |
| **[developer/](docs/developer/)** | Contributing — module / AI tool / detector authoring |
| **[samples/](docs/samples/)** | Anonymized example incidents, dashboards, plans, sessions |

The master TOC lives at **[`docs/README.md`](docs/README.md)**.

## Highlights

- **Every mutation is auditable + reversible** where reversal is possible. The audit log is JSON-per-line with `entryId` correlation, structured `target` hashtable, and `reverse` recipe; the undo system dispatches by `actionType`. Filter / view / undo via menu option 14 (Audit & Reporting). See [`docs/reference/audit-format.md`](docs/reference/audit-format.md).
- **Two-mode operation.** On launch, the operator picks `LIVE` or `PREVIEW`. PREVIEW logs intent without touching the tenant. Switch in chat with `/dryrun`.
- **Defense in depth on secrets.** API keys, webhook URLs, scheduler credentials, and tenant client-secrets are all DPAPI-encrypted at rest (per-user, per-machine — not portable). Run option 99 (AI Assistant) → `/config` to seed; the plaintext key is encrypted in place on first save. See [`docs/concepts/security-model.md`](docs/concepts/security-model.md).
- **PII redaction by default.** External-LLM payloads are tokenized (`<UPN_1>`, `<GUID_3>`, `<TENANT>`, etc.) before send and restored on response. Secrets (`sk-…`, JWTs, cert thumbprints) are always tokenized regardless of provider. Configure via `/privacy` in the chat.
- **Multi-tenant from one console.** Register N tenants, switch with `Switch-Tenant -Name <n>` (or `/tenant <name>` in chat), run cross-tenant reports via `Invoke-AcrossTenants`. The audit log filename includes the tenant slug so cross-tenant operations stay distinct. See [`docs/concepts/multi-tenant.md`](docs/concepts/multi-tenant.md).

## Tests

```powershell
Invoke-Pester ./tests/
```

224 tests across 22 suites. No live Graph / EXO / SPO calls — every assertion uses canned data or mocked SDK calls. See [`docs/developer/testing.md`](docs/developer/testing.md).

## Authentication

Browser-based OAuth via the Microsoft Graph PowerShell SDK + EXO/SPO modules. Three modes:

- **Direct admin** — your own organization, interactive sign-in.
- **GDAP partner** — customer tenants the operator has delegated admin rights to.
- **App-only via tenant profile** — for unattended automation. Cert-thumbprint preferred; client-secret supported with a warning. See [`docs/concepts/multi-tenant.md`](docs/concepts/multi-tenant.md).

## Help / feedback

- Issues: https://github.com/Neoo-Blue/M365-Manager/issues
- Look in [`docs/operations/troubleshooting.md`](docs/operations/troubleshooting.md) for known errors → fix mappings.

## Design principles

- **Confirmation first.** Every destructive operation prompts. Bulk flows validate before any mutation.
- **Audit everything.** A scheduled run that operates without an operator is still answerable to a forensic question six months later.
- **Reverse where possible.** Most mutations have a curated reverse recipe; the rest are flagged with an explicit `NoUndoReason`.
- **Make the dangerous things obvious.** Red `[LIVE]` banner. Yellow `[PREVIEW]` banner. Per-tenant color rotation. `EXPLICIT APPROVAL REQUIRED` warning on the highest-blast-radius tool.
- **Don't surprise the operator.** Default-off auto-execution. Default-on PII redaction. Default-on undo recipes. Default to asking.

## License

[Add your license file here]
