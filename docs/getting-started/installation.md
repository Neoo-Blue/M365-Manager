# Installation

End-to-end install steps for a new operator workstation. Total time: about 10 minutes if PowerShell + the modules are already installed; 30 minutes from a clean Windows box.

## Prerequisites

### PowerShell

The tool targets **PowerShell 7+**. Earlier 5.1 mostly works too — the codebase started life on 5.1 — but PS 7 is the version Pester runs against and the version the docs assume.

```powershell
$PSVersionTable.PSVersion
# Want 7.x or higher. If you see 5.1, install PS 7 separately:
#   winget install Microsoft.PowerShell
# Or on Mac: brew install powershell
```

### Microsoft modules

The launcher auto-installs these on first run, but if your machine has `Constrained Language Mode`, `AllSigned` execution policy, or no internet at run-time, pre-install:

```powershell
Install-Module Microsoft.Graph.Authentication             -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users                      -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users.Actions              -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Groups                     -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement                   -Scope CurrentUser -Force
Install-Module Microsoft.Online.SharePoint.PowerShell     -Scope CurrentUser -Force
```

The Security & Compliance Center cmdlets (used by `eDiscovery.ps1` and Phase 7's `New-ComplianceSearch` purge step) load via `Connect-IPPSSession`, which ships with `ExchangeOnlineManagement` — no separate install.

### gh CLI (optional)

Only needed if you'll be filing issues / PRs against this repo. `winget install GitHub.cli` on Windows, `brew install gh` on Mac.

### Account permissions

The connecting account needs one of:

| Mode | Roles |
|---|---|
| **Full** | Global Administrator |
| **Granular (least privilege)** | User Administrator + License Administrator + Exchange Administrator + SharePoint Administrator + Teams Administrator + Compliance Administrator + Security Administrator |
| **Read-only** | Global Reader + Security Reader (works for reports and PREVIEW; mutations need write scopes) |

GDAP partner access works too — the operator connects to the partner tenant first, then picks the customer tenant. See [tenant-setup.md](tenant-setup.md).

## Install steps

### 1. Clone the repo

```cmd
git clone https://github.com/Neoo-Blue/M365-Manager.git
cd M365-Manager
```

### 2. Allow script execution

If your machine has `Restricted` execution policy (Win 11 default), the tool won't run. Two options:

```powershell
# Option A (recommended): just for the current user
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Option B: just for the single launch (use Launch.bat which sets this)
# -- nothing to do here, Launch.bat handles it.
```

### 3. First launch

```cmd
Launch.bat
```

The launcher does three things:
- Sets `MSAL_BROKER_ENABLED=0` (prevents a DLL conflict between Graph and Exchange Online modules).
- Sets `M365ADMIN_ROOT` so the tool can find itself.
- Runs `pwsh -ExecutionPolicy Bypass -NoProfile -File Main.ps1`.

On non-Windows (Mac / Linux):

```bash
$env:MSAL_BROKER_ENABLED = "0"
pwsh -ExecutionPolicy Bypass -NoProfile -File ./Main.ps1
```

### 4. What you'll see

1. **Banner** — colored title + version stamp.
2. **Module check** — every required module is verified. Missing modules trigger an auto-install prompt.
3. **Cache cleanup** — any leftover Graph / EXO sessions from a prior run are disconnected so you start clean.
4. **Operating-mode picker** — `LIVE` (changes apply) vs `PREVIEW` (dry-run; logs intent only). **Pick PREVIEW for your first session** — you'll get a feel for the audit format without risk.
5. **Tenant selection** — first run shows the partner-vs-direct picker. After you register at least one tenant profile (Phase 6), this becomes a list of registered tenants. See [tenant-setup.md](tenant-setup.md).
6. **Connection status bar** — `Graph: OK`, `EXO: OK`, `SCC: OK` on success.
7. **Main menu**.

## Configuration after first launch

Three things to configure:

1. **AI assistant** (optional but recommended) — option 99 (hidden) → walks through provider / model / key setup. Stored DPAPI-encrypted in `ai_config.json` (gitignored). See [configuration.md](configuration.md).
2. **Tenant profiles** — register your tenant(s) for fast switching. See [tenant-setup.md](tenant-setup.md).
3. **Notifications** — if you want email / Teams alerts from health checks, configure recipients + webhook URLs. See [`../guides/notifications.md`](../guides/notifications.md).

## Verify the install

From the tool's main menu:

1. **Run Pester** (one-time sanity):
   ```powershell
   Invoke-Pester ./tests/
   ```
   Expected: `224 / 0 / 0` (passed / failed / skipped). If any test fails, file an issue with the output — that's a regression that shouldn't have made it to `main`.

2. **Connect + run a read-only report**: option 11 → Reports → Tenant overview. Should print user count, license SKUs, group count.

3. **Check the audit log** is being written: option 14 → Audit & Reporting → Open audit log viewer. The session-start line should be visible.

## Troubleshooting first-run failures

See [`../operations/troubleshooting.md`](../operations/troubleshooting.md) for the full error-message → fix index. The four most common first-run snags:

| Symptom | Fix |
|---|---|
| `Cannot find module Microsoft.Graph.Authentication` | Run `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force`. Restart pwsh. |
| `Module conflict: Microsoft.Graph.Identity vs Exchange Online` | You launched without `Launch.bat`. Set `$env:MSAL_BROKER_ENABLED = "0"` and relaunch. |
| `AADSTS50034: The user account does not exist in tenant` | You picked the wrong tenant at the OAuth prompt. Quit (`Q`), relaunch, pick a different tenant. |
| Browser opens to OAuth but tool says "no consent given" | Click the consent dialog through to the end; you must consent to every scope the tool requests on first sign-in. |

## See also

- [quickstart.md](quickstart.md) — 10-minute hands-on tour.
- [tenant-setup.md](tenant-setup.md) — registering tenant profiles.
- [configuration.md](configuration.md) — every config key.
- [`../operations/permissions.md`](../operations/permissions.md) — full Graph scope + EXO/SPO/SCC role matrix per feature.
