# Deployment

Rolling out M365 Manager on operator workstations. Covers single-machine deploy, signed-module distribution, execution policy, and team / fleet rollout.

## Single-machine deploy

The base case. See [`../getting-started/installation.md`](../getting-started/installation.md) — clone the repo, install modules, run `Launch.bat`.

This works for solo operators and consultants. For fleet deployment (an MSP rolling out to a tech team, or an internal IT shop with multiple admins), use the patterns below.

## Fleet deploy

### Option A — shared network share

```text
\\fileserver\IT\Tools\M365Manager\
├── Launch.bat
├── Main.ps1
├── ... (every .ps1)
├── ai-tools\
├── templates\
├── docs\
└── tests\
```

Every operator runs `Launch.bat` from the share. Pros:

- One copy to maintain.
- Updates roll out by replacing the share contents.

Cons:

- Network share must be online when the operator launches.
- DPAPI-encrypted `ai_config.json` and per-tenant secrets land on each operator's `%LOCALAPPDATA%`, not the share — operators need to configure once per workstation.
- Versions drift if an operator pins the share to a local mirror.

### Option B — per-workstation clone + scheduled pull

```powershell
# In a logon script or scheduled task on each workstation:
$repoPath = "C:\Tools\M365Manager"
if (-not (Test-Path $repoPath)) {
    git clone https://github.com/Neoo-Blue/M365-Manager $repoPath
} else {
    git -C $repoPath pull --ff-only
}
```

Pros: every workstation has a local copy + version pin via tag. Cons: each operator must have `git` + network access.

### Option C — packaged install

For the highest reliability deploy, package the tool into an MSI or Chocolatey package:

```powershell
# Chocolatey package skeleton (chocolateyInstall.ps1):
$packageName = 'm365manager'
$installPath = "$env:ProgramData\M365Manager"
Copy-Item "$PSScriptRoot\tools\*" $installPath -Recurse -Force
Install-ChocolateyShortcut -shortcutFilePath "$env:Public\Desktop\M365 Manager.lnk" `
    -targetPath "$installPath\Launch.bat"
```

Pros: standard Windows software lifecycle (install / upgrade / uninstall). Cons: requires CI to build packages on each release.

## Execution policy

The tool needs `RemoteSigned` at minimum.

```powershell
# Per-user (recommended; non-admin):
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Per-machine (admin):
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
```

`Launch.bat` passes `-ExecutionPolicy Bypass` for the launched process, so even a `Restricted` machine can run the tool from `Launch.bat`. But this means: anyone with file-write access to `Launch.bat` can run anything they want under the operator's identity. Production deploys should:

1. Run from a read-only network share OR a signed package directory.
2. Set machine policy to `AllSigned` and sign every `.ps1` in the repo (see "Signed modules" below).
3. Drop the `-Bypass` flag from `Launch.bat`.

## Signed modules

For shops with strict execution-policy posture:

### Generate a code-signing cert

Either purchase one from a public CA, or generate an internal one from your enterprise CA:

```powershell
# Self-signed for testing:
$cert = New-SelfSignedCertificate -CertStoreLocation Cert:\CurrentUser\My `
    -Subject "CN=M365 Manager Code Signing" `
    -KeyUsage DigitalSignature `
    -Type CodeSigningCert `
    -NotAfter (Get-Date).AddYears(2)
```

### Sign every script

```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My\<thumbprint>
Get-ChildItem -Path .\*.ps1 -Recurse | ForEach-Object {
    Set-AuthenticodeSignature -FilePath $_.FullName -Certificate $cert -TimestampServer "http://timestamp.digicert.com"
}
```

Run this in the build / release pipeline; commit the signed versions, OR sign at package time.

### Verify on the operator workstation

```powershell
Get-AuthenticodeSignature .\Main.ps1
# Status should be Valid; signer should be your CN.
```

With `AllSigned` policy, only signed scripts run. Without the right cert in `Trusted Publishers`, PowerShell prompts on every script execution — chain the cert into the trust store via GPO or `Import-Certificate`.

## Per-operator first-run

After the tool is on the workstation, each operator runs the first-run flow once:

1. Launch via `Launch.bat`.
2. Configure AI provider (option 99 → `/config`). The API key gets DPAPI-encrypted per-user. **Each operator needs their own AI key** — sharing one is awkward because DPAPI binds to the user.
3. Register tenant profile(s) (option 22 → Tenants → Register). Same DPAPI scope.
4. Configure notifications (option 20 → Notifications setup).

See [`../getting-started/installation.md`](../getting-started/installation.md) for the post-install verification checklist.

## State directory

`<stateDir>` (`%LOCALAPPDATA%\M365Manager\state`) is **per-operator**. Don't try to roam it. The DPAPI-encrypted blobs in `secrets/` won't decrypt on a different user account; chat sessions and incident artifacts are scoped to the operator who created them.

For shared state (e.g. shared incident registry across an IR team), the future plan is a backing store outside `<stateDir>` — not implemented today. Track in [`pre-merge-review.md`](pre-merge-review.md).

## Upgrades

See [`upgrade-guide.md`](upgrade-guide.md). TL;DR:

- `git pull` (or update the package) gets new functionality.
- The audit log format and tenant-profile format are versioned and backward-compatible — no migration needed for routine updates.
- Breaking changes will be called out in `CHANGELOG.md` (not yet shipped).

## CI / automation deploy

For the headless / scheduled use case (running the tool from a CI runner or Azure Automation Runbook):

1. Clone the repo on the agent.
2. Pre-install all PowerShell modules in the agent image.
3. Use a CertThumbprint tenant profile (see [`../getting-started/tenant-setup.md`](../getting-started/tenant-setup.md)) — interactive auth doesn't work on headless agents.
4. Set `$env:M365MGR_NONINTERACTIVE = "1"` so every prompt declines.
5. Drive specific functions directly (`Invoke-BulkOffboard`, `Invoke-CompromisedAccountResponse`, etc.) — don't launch `Main.ps1`.

The Phase 4 scheduler module already does this for health checks; the same pattern works for ad-hoc automation.

## See also

- [`../getting-started/installation.md`](../getting-started/installation.md) — single-machine install.
- [`upgrade-guide.md`](upgrade-guide.md) — version-to-version migration.
- [`permissions.md`](permissions.md) — what to grant the connecting account.
- [`troubleshooting.md`](troubleshooting.md) — common deploy-time issues.
