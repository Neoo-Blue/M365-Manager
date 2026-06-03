# ============================================================
#  Auth.ps1 - Authentication, Dependencies & Session Management
# ============================================================

# Disable MSAL WAM broker globally for this process.
# Prevents DLL version conflicts between MS Graph SDK and EXO module.
# Forces all connections to use standard browser auth instead.
[System.Environment]::SetEnvironmentVariable("MSAL_BROKER_ENABLED", "0", "Process")

$script:SessionState = @{
    MgGraph            = $false
    ExchangeOnline     = $false
    ComplianceCenter   = $false
    SharePointOnline   = $false
    SharePointAdminUrl = $null
    TenantMode         = "Direct"
    TenantId           = $null
    TenantName         = $null
    TenantDomain       = $null
    PartnerConnected   = $false
}

$script:MgScopes = @(
    "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All",
    "Organization.Read.All","UserAuthenticationMethod.ReadWrite.All",
    "AuditLog.Read.All","Reports.Read.All","Mail.Send",
    "Sites.FullControl.All",
    "TeamMember.ReadWrite.All","TeamSettings.ReadWrite.All",
    "Channel.ReadBasic.All","ChannelMember.ReadWrite.All",
    "Policy.Read.All"
)
$script:MgPartnerScopes = @("Directory.Read.All","Contract.Read.All")

# ============================================================
#  Dependency Management - install, import, verify
# ============================================================

function Test-IsAdmin {
    <#
        True when the current PowerShell process is elevated.
        On non-Windows we have no UAC, so this is always $false
        (Install-Module -Scope CurrentUser is always safe there).
    #>
    if ($IsWindows -or ($PSVersionTable.PSVersion.Major -lt 6)) {
        try {
            $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
            $pr = [System.Security.Principal.WindowsPrincipal]::new($id)
            return $pr.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
        } catch { return $false }
    }
    return $false
}

function Request-AdminElevation {
    <#
        Re-launch PowerShell elevated to install one or more modules
        for AllUsers, then exit. Returns $false if the operator
        declines or we're not on Windows.
    #>
    param([string[]]$ModulesToInstall)
    if (-not ($IsWindows -or ($PSVersionTable.PSVersion.Major -lt 6))) {
        Write-Warn "Elevation only applies on Windows. Skipping."
        return $false
    }
    Write-Host ""
    Write-Warn "Admin privileges are recommended to install for AllUsers:"
    foreach ($m in $ModulesToInstall) { Write-Host "    - $m" -ForegroundColor White }
    Write-Host ""
    $ans = Read-UserInput "Relaunch elevated and install now? (y/n)"
    if ($ans -notmatch '^[Yy]') { return $false }

    $installCmd = ($ModulesToInstall | ForEach-Object {
        "Install-Module $_ -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Continue"
    }) -join '; '
    $bootstrap = "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force; " +
                 "[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12; " +
                 "if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) { Install-PackageProvider -Name NuGet -Force -ForceBootstrap | Out-Null }; " +
                 "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue; " +
                 "$installCmd; Write-Host ''; Write-Host 'Done. You may close this elevated window.' -ForegroundColor Green; Read-Host 'Press Enter to close'"

    $psExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    try {
        Start-Process -FilePath $psExe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-Command', $bootstrap) -Verb RunAs -ErrorAction Stop -Wait
        Write-Success "Elevated install finished. Re-checking modules..."
        return $true
    } catch {
        Write-ErrorMsg "Could not start elevated PowerShell: $_"
        return $false
    }
}

function Test-ModuleSmoke {
    <#
        Verify a module is loaded AND its key cmdlet is callable.
        Returns @{ Ok = $bool; Reason = '...' }.
    #>
    param([string]$ModuleName, [string]$TestCmd)
    $loaded = Get-Module -Name $ModuleName
    if (-not $loaded) { return @{ Ok = $false; Reason = "module not in session" } }
    if ($TestCmd) {
        $cmd = Get-Command $TestCmd -ErrorAction SilentlyContinue
        if (-not $cmd) { return @{ Ok = $false; Reason = "cmdlet $TestCmd not exported" } }
    }
    return @{ Ok = $true; Reason = "v$($loaded.Version)" }
}

function Initialize-PSGalleryBootstrap {
    <#
        One-time-per-session bootstrap of the things PSGallery needs to
        work on Windows PowerShell 5.1 from a clean machine:
          1. Force TLS 1.2. PS 5.1 defaults to TLS 1.0/1.1; PSGallery
             dropped sub-1.2 ages ago. Without this, Install-Module
             silently fails with "no module found" style errors.
          2. Bootstrap the NuGet PackageProvider. Without it, the
             first Install-Module prompts interactively (blocks
             unattended runs) or returns an opaque PackageProvider
             error.
          3. Trust PSGallery for this session so Install-Module doesn't
             stop on an "Untrusted repository" Y/N prompt.
        All steps are idempotent and safe to re-run.
    #>
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch {}

    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-InfoMsg "Bootstrapping NuGet PackageProvider..."
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ForceBootstrap -Scope CurrentUser -ErrorAction Stop | Out-Null
        } catch {
            Write-Warn "NuGet bootstrap failed: $_"
        }
    }

    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }
    } catch {}
}

# ---- Microsoft.Graph version policy ---------------------------------
# Baseline: the version we treat as "good" and aim to install.
# Known-bad: versions we explicitly refuse to accept even when nothing
# else looks wrong (e.g. v2.37.0 ships a broken Authentication.Core.dll
# with no GetTokenAsync impl, which makes every import fail).
$script:MgGraphBaselineVersion = '2.25.0'
$script:MgGraphKnownBadVersions = @('2.37.0')

function Test-MgGraphVersionCoherence {
    <#
        Returns @() when the on-disk state is acceptable, otherwise an
        array of human-readable problem lines. "Acceptable" means
        every sibling resolves to the same NON-broken version. That
        version doesn't have to be the baseline -- if the operator
        deliberately pinned to 2.30.0 and everything matches, we stay
        quiet.

        We do flag:
          (a) Sibling drift -- different siblings have different
              top-installed versions.
          (b) Any sibling whose top-installed version appears in
              $MgGraphKnownBadVersions.
          (c) Multiple DISTINCT versions of the same sibling on disk
              (an old copy lurking can load first and trigger
              "Could not load file or assembly").

        Same version installed in TWO scopes (CurrentUser + AllUsers
        at v2.25.0) is NOT flagged -- it's harmless.
    #>
    $all = @(Get-Module -ListAvailable -Name 'Microsoft.Graph.*' |
             Where-Object { $_.Name -ne 'Microsoft.Graph' })
    if ($all.Count -lt 2) { return @() }
    $msgs = @()

    # (a) Sibling drift.
    $tops = @($all | Group-Object Name | ForEach-Object {
        $_.Group | Sort-Object Version -Descending | Select-Object -First 1
    })
    $topVersions = $tops | ForEach-Object { [string]$_.Version } | Sort-Object -Unique
    if ($topVersions.Count -gt 1) {
        $msgs += "Microsoft.Graph.* sibling drift: $($topVersions -join ', ')"
        foreach ($t in $tops) { $msgs += "    $($t.Name) v$($t.Version)" }
    }

    # (b) Known-bad version detected on a top-installed sibling.
    $badHits = @($tops | Where-Object { $script:MgGraphKnownBadVersions -contains [string]$_.Version })
    if ($badHits.Count -gt 0) {
        $msgs += "Known-broken Microsoft.Graph version on disk:"
        foreach ($b in $badHits) {
            $msgs += "    $($b.Name) v$($b.Version)  <-- known-broken release"
        }
        $msgs += "    (Pin to v$($script:MgGraphBaselineVersion) via option 98.)"
    }

    # (c) Multiple installed VERSIONS of the same sibling.
    $multi = @()
    foreach ($g in @($all | Group-Object Name)) {
        $distinctVersions = @($g.Group | ForEach-Object { [string]$_.Version } | Sort-Object -Unique)
        if ($distinctVersions.Count -gt 1) {
            $multi += [PSCustomObject]@{ Name = $g.Name; Versions = $distinctVersions }
        }
    }
    if ($multi.Count -gt 0) {
        $msgs += "Multiple installed versions of the same Microsoft.Graph.* module(s):"
        foreach ($g in $multi) {
            $msgs += "    $($g.Name) -> $($g.Versions -join ', ')"
        }
        $msgs += "    (PowerShell may load an old copy first. Run option 98 to clean up.)"
    }

    return $msgs
}

function Repair-MgGraphSiblings {
    <#
        Bring every Microsoft.Graph.* sibling we use to the same
        version, AND uninstall the older copies that are still on
        disk. Without the uninstall step, Install-Module just lays
        the new version next to the old one, so the next startup
        coherence check keeps flagging "Multiple installed versions
        of the same Microsoft.Graph.* module(s)".

        IMPORTANT: an in-process Install-Module CANNOT replace a sibling
        whose DLLs are already loaded into this session. If we detect
        any Microsoft.Graph.* in Get-Module (loaded), we refuse the
        in-place repair and steer the operator to Invoke-MgGraphFullRepair
        (hidden main-menu option 98) which spawns an elevated, fresh
        PowerShell where nothing is loaded yet.

        Uninstalling AllUsers-scoped modules requires admin. When we
        hit that, we surface a clear message pointing at option 98.
    #>
    param([string[]]$Names)
    $loaded = @(Get-Module -Name 'Microsoft.Graph.*' | Where-Object { $_.Name -ne 'Microsoft.Graph' })
    if ($loaded.Count -gt 0) {
        Write-Host ""
        Write-Warn ("Already loaded into this session: {0}" -f (($loaded | ForEach-Object { "$($_.Name) v$($_.Version)" }) -join ', '))
        Write-Host "  In-process repair cannot replace DLLs that are already loaded." -ForegroundColor Yellow
        Write-Host "  At the main menu, type " -NoNewline -ForegroundColor Yellow
        Write-Host "98" -NoNewline -ForegroundColor Cyan
        Write-Host " ('Repair Microsoft.Graph modules' hidden option)." -ForegroundColor Yellow
        Write-Host "  That spawns an elevated, fresh PowerShell window which CAN replace the modules." -ForegroundColor Yellow
        Write-Host ""
        return
    }

    $authMod = Get-Module -ListAvailable Microsoft.Graph.Authentication |
               Sort-Object Version -Descending | Select-Object -First 1
    if (-not $authMod) {
        Write-Warn "Microsoft.Graph.Authentication isn't installed yet; nothing to align to."
        return
    }
    # Aim at the BASELINE version, not the highest installed. The
    # highest is often the broken v2.37.0; defaulting to the baseline
    # makes a "press Y at the prompt" do the right thing without
    # asking the operator to type a version. If $authMod itself is
    # already at the baseline, this is a no-op.
    $target = [System.Version]$script:MgGraphBaselineVersion
    Write-InfoMsg "Aligning Microsoft.Graph.* siblings to baseline v$target..."

    # ---- Pass 1: uninstall every non-target version on disk ----
    $needElevation = $false
    foreach ($n in $Names) {
        $olds = @(Get-Module -ListAvailable -Name $n | Where-Object { $_.Version -ne $target })
        foreach ($o in $olds) {
            try {
                Uninstall-Module -Name $n -RequiredVersion $o.Version -Force -ErrorAction Stop
                Write-Success "  Uninstalled $n v$($o.Version)"
            } catch {
                $msg = "$_"
                if ($msg -match 'Administrator|elevation|Access.*denied|Unauthorized') {
                    Write-Warn "  Cannot uninstall $n v$($o.Version) without elevation (likely AllUsers scope)."
                    $needElevation = $true
                } else {
                    Write-Warn "  Uninstall $n v$($o.Version) failed: $msg"
                }
            }
        }
    }

    if ($needElevation) {
        Write-Host ""
        Write-Warn "Old Microsoft.Graph.* versions are installed AllUsers-scope and need admin to remove."
        Write-Host "  At the main menu, type " -NoNewline -ForegroundColor Yellow
        Write-Host "98" -NoNewline -ForegroundColor Cyan
        Write-Host " to run the elevated uninstall + reinstall." -ForegroundColor Yellow
        Write-Host ""
        return
    }

    # ---- Pass 2: install target version where missing ----
    $inUseCount = 0
    foreach ($n in $Names) {
        $haveTarget = Get-Module -ListAvailable -Name $n | Where-Object { $_.Version -eq $target }
        if ($haveTarget) {
            Write-InfoMsg "  $n v$target already installed."
            continue
        }
        $warned = $null
        try {
            Install-Module $n -RequiredVersion $target -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -WarningVariable warned -ErrorAction Stop
            Write-Success "  $n -> v$target"
            if ($warned -and (($warned | Out-String) -match 'in use|currently in use')) { $inUseCount++ }
        } catch {
            $emsg = "$_"
            Write-Warn "  $n : $emsg"
            if ($emsg -match 'in use|currently in use') { $inUseCount++ }
        }
    }

    # ---- Pass 3: re-check coherence and report ----
    $stillDrifted = Test-MgGraphVersionCoherence
    if ($stillDrifted.Count -eq 0) {
        Write-Success "Microsoft.Graph.* siblings are now coherent at v$target."
    } else {
        Write-Warn "Coherence check still reports issues:"
        foreach ($l in $stillDrifted) { Write-Host "    $l" -ForegroundColor Yellow }
    }

    if ($inUseCount -gt 0) {
        Write-Host ""
        Write-Warn ("{0} module(s) reported 'currently in use'. This is NOT a failure --" -f $inUseCount)
        Write-Host "  the v$target files ARE on disk now. The current PowerShell session" -ForegroundColor Yellow
        Write-Host "  still has the old DLLs loaded and won't pick up the new ones until" -ForegroundColor Yellow
        Write-Host "  you close this window and re-launch the tool." -ForegroundColor Yellow
        Write-Host ""
    }
}

function Write-MgGraphAssemblyMismatchHelp {
    <#
        If $Message looks like a "Could not load file or assembly
        Microsoft.Graph.* Version=X.Y.Z" error, print a clear
        remediation block and OFFER to run the elevated repair now.
        Returns $true so the caller can short-circuit further error
        chatter; $false when the message doesn't match.
    #>
    param([Parameter(Mandatory)][string]$Message)
    if ($Message -notmatch 'Could not load file or assembly.*Microsoft\.Graph') { return $false }
    $version = ''
    if ($Message -match 'Version=([\d\.]+)') { $version = $Matches[1] }

    Write-Host ""
    Write-ErrorMsg "Microsoft.Graph module version mismatch / incomplete install."
    if ($version) {
        Write-Host ("    Looking for Microsoft.Graph.Authentication.Core v{0} but it isn't on disk" -f $version) -ForegroundColor Yellow
        Write-Host  "    (or an older sibling version got loaded first and is masking it)." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "  Why a plain re-launch didn't fix it: PowerShell can load an OLDER" -ForegroundColor Yellow
    Write-Host "  sibling first when MULTIPLE versions are installed on disk. The" -ForegroundColor Yellow
    Write-Host "  reliable cure is to uninstall ALL Microsoft.Graph.* versions and" -ForegroundColor Yellow
    Write-Host "  reinstall a clean coherent set." -ForegroundColor Yellow
    Write-Host ""

    $ans = Read-UserInput "Run the elevated Repair-Microsoft.Graph fix now? (Y/n)"
    if ($ans -notmatch '^[Nn]') {
        if (Get-Command Invoke-MgGraphFullRepair -ErrorAction SilentlyContinue) {
            Invoke-MgGraphFullRepair
            Write-Host ""
            Write-Warn "Repair done. You MUST close this M365 Manager window and re-launch."
            Write-Host "  Press any key to return..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } else {
            Write-Warn "Repair helper not loaded; at the main menu type 98 <enter>."
        }
        return $true
    }

    Write-Host ""
    Write-Host "  Manual recovery from a NEW elevated PowerShell window:" -ForegroundColor Yellow
    Write-Host "    Get-Module Microsoft.Graph.* -ListAvailable | Uninstall-Module -AllVersions -Force -ErrorAction SilentlyContinue" -ForegroundColor Cyan
    Write-Host "    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" -ForegroundColor Cyan
    Write-Host "    Install-Module Microsoft.Graph.Authentication              -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck" -ForegroundColor Cyan
    Write-Host "    Install-Module Microsoft.Graph.Users                       -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck" -ForegroundColor Cyan
    Write-Host "    Install-Module Microsoft.Graph.Users.Actions               -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck" -ForegroundColor Cyan
    Write-Host "    Install-Module Microsoft.Graph.Groups                      -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck" -ForegroundColor Cyan
    Write-Host "    Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck" -ForegroundColor Cyan
    Write-Host ""
    return $true
}

function Invoke-MgGraphFullRepair {
    <#
        Operator-triggered full repair. Spawns an elevated PowerShell
        window that uninstalls every Microsoft.Graph.* module on
        disk and reinstalls a coherent set.

        CRITICAL: the elevated window MUST run with this M365 Manager
        process already exited. Windows holds file locks on every
        Microsoft.Graph.*.dll that the parent process imported, so
        even an elevated child can't delete them while we're alive:
        "WARNING: The version 'X' of module '...' is currently in use."

        So this function:
          1. Spawns the elevated child WITHOUT -Wait.
          2. Exits the M365 Manager process immediately, releasing the
             DLL locks so the child's uninstall can succeed.
          3. The child window remains open until the operator presses
             Enter, then prompts them to re-launch Launch.bat.
    #>
    Write-SectionHeader "Repair Microsoft.Graph modules"
    Write-Host "  This will:" -ForegroundColor $script:Colors.Info
    Write-Host "    1. Launch an elevated PowerShell window with the repair script." -ForegroundColor White
    Write-Host "    2. Immediately CLOSE this M365 Manager session (mandatory: the" -ForegroundColor White
    Write-Host "       Microsoft.Graph DLLs we have loaded would block the uninstall)." -ForegroundColor White
    Write-Host "    3. The elevated window uninstalls ALL Microsoft.Graph.* modules" -ForegroundColor White
    Write-Host "       and reinstalls Authentication, Users, Users.Actions, Groups," -ForegroundColor White
    Write-Host "       and Identity.DirectoryManagement PINNED TO ONE KNOWN-GOOD VERSION." -ForegroundColor White
    Write-Host "    4. When the elevated window says Done, double-click Launch.bat" -ForegroundColor White
    Write-Host "       to restart this tool." -ForegroundColor White
    Write-Host ""
    Write-Host "  Why a specific version? Microsoft.Graph 2.37.0 ships a broken" -ForegroundColor Yellow
    Write-Host "  Authentication.Core.dll (missing GetTokenAsync impl). 2.25.0 is" -ForegroundColor Yellow
    Write-Host "  the last widely-deployed stable build for Windows PowerShell 5.1." -ForegroundColor Yellow
    Write-Host ""

    $defaultPin = $script:MgGraphBaselineVersion
    $verIn = Read-UserInput ("Pin to Microsoft.Graph version (Enter for {0})" -f $defaultPin)
    $pinVersion = if ([string]::IsNullOrWhiteSpace($verIn)) { $defaultPin } else { $verIn.Trim() }

    if (-not (Confirm-Action ("Proceed (will install v{0} as AllUsers and CLOSE this session)?" -f $pinVersion))) { return }

    $cmds = @(
        "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12",
        "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force",
        "Write-Host 'Waiting 3s for parent M365 Manager to release DLL locks...' -ForegroundColor Yellow",
        "Start-Sleep -Seconds 3",
        "if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) { Install-PackageProvider -Name NuGet -Force -ForceBootstrap | Out-Null }",
        "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue",
        "Write-Host 'Uninstalling EVERY installed version of Microsoft.Graph.* ...' -ForegroundColor Yellow",
        "Get-Module Microsoft.Graph.* -ListAvailable | Uninstall-Module -AllVersions -Force -ErrorAction SilentlyContinue",
        "Write-Host ('Installing Microsoft.Graph.* v' + '$pinVersion' + ' AllUsers...') -ForegroundColor Yellow",
        "Install-Module Microsoft.Graph.Authentication              -RequiredVersion $pinVersion -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Continue",
        "Install-Module Microsoft.Graph.Users                       -RequiredVersion $pinVersion -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Continue",
        "Install-Module Microsoft.Graph.Users.Actions               -RequiredVersion $pinVersion -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Continue",
        "Install-Module Microsoft.Graph.Groups                      -RequiredVersion $pinVersion -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Continue",
        "Install-Module Microsoft.Graph.Identity.DirectoryManagement -RequiredVersion $pinVersion -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Continue",
        "Write-Host ''; Write-Host ('Done. All Microsoft.Graph.* pinned to v' + '$pinVersion' + '. Double-click Launch.bat to restart.') -ForegroundColor Green",
        "Read-Host 'Press Enter to close this window'"
    ) -join '; '

    $psExe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    try {
        # Detach (no -Wait) so this M365 Manager session can exit
        # immediately and free the DLL locks before the elevated
        # child reaches the Uninstall-Module step (it sleeps 3s).
        Start-Process -FilePath $psExe `
                      -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-Command', $cmds) `
                      -Verb RunAs -ErrorAction Stop
    } catch {
        Write-ErrorMsg "Could not start elevated PowerShell: $_"
        Pause-ForUser
        return
    }

    Write-Host ""
    Write-Success "Elevated repair window launched."
    Write-Warn "Closing this M365 Manager session NOW so the DLL locks release."
    Write-Host "  Watch the elevated window. When it says 'Done', re-launch Launch.bat." -ForegroundColor Yellow
    Write-Host ""
    Start-Sleep -Seconds 2

    # Best-effort: drop the in-process Microsoft.Graph.* modules so
    # any future Connect-* attempts here don't reuse the now-stale
    # binary. Then exit the process so file handles release.
    try {
        Get-Module Microsoft.Graph.* | Remove-Module -Force -ErrorAction SilentlyContinue
    } catch {}
    [Environment]::Exit(0)
}

function Assert-ModulesInstalled {
    Write-SectionHeader "Checking Dependencies"

    Initialize-PSGalleryBootstrap

    # ---- Define required modules with test commands ----
    # IMPORTANT: ExchangeOnlineManagement MUST be first.
    # It loads its MSAL assemblies first, preventing version conflicts
    # when Graph modules try to load a different MSAL version later.
    $requiredModules = @(
        @{ Name = "ExchangeOnlineManagement";                    TestCmd = "Connect-ExchangeOnline" },
        @{ Name = "Microsoft.Graph.Authentication";              TestCmd = "Get-MgContext" },
        @{ Name = "Microsoft.Graph.Users";                       TestCmd = "Get-MgUser" },
        @{ Name = "Microsoft.Graph.Users.Actions";               TestCmd = "Revoke-MgUserSignInSession" },
        @{ Name = "Microsoft.Graph.Groups";                      TestCmd = "Get-MgGroup" },
        @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; TestCmd = "Get-MgSubscribedSku" },
        @{ Name = "Microsoft.Online.SharePoint.PowerShell";      TestCmd = "Connect-SPOService"; Optional = $true }
    )

    # ---- Pre-flight: detect Microsoft.Graph sibling version drift. ----
    # If we don't fix this BEFORE the install/import loop, importing a
    # newer Users module against an older Authentication will throw and
    # the operator will see a confusing nested-import error.
    $coherence = Test-MgGraphVersionCoherence
    if ($coherence.Count -gt 0) {
        Write-Warn "Microsoft.Graph submodule version mismatch detected:"
        foreach ($l in $coherence) { Write-Host "    $l" -ForegroundColor Yellow }
        $ans = Read-UserInput "Reinstall all Microsoft.Graph.* siblings to a single version now? (Y/n)"
        if ($ans -notmatch '^[Nn]') {
            $graphNames = @($requiredModules | Where-Object { $_.Name -like 'Microsoft.Graph.*' } | ForEach-Object { $_.Name })
            Repair-MgGraphSiblings -Names $graphNames
        } else {
            Write-Warn "Continuing with mismatched versions; import may fail."
        }
    }

    # ---- Pass 1: figure out what's actually missing ----
    $missingRequired = @()
    $missingOptional = @()
    foreach ($mod in $requiredModules) {
        $inst = Get-Module -ListAvailable -Name $mod.Name | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $inst) {
            if ($mod.Optional) { $missingOptional += $mod.Name }
            else { $missingRequired += $mod.Name }
        }
    }

    # ---- Pass 2: if anything is missing, offer elevation up front ----
    $toInstall = @($missingRequired) + @($missingOptional | Where-Object { $true })
    if ($toInstall.Count -gt 0) {
        Write-Host ""
        Write-Warn ("Missing module(s): {0}" -f ($toInstall -join ', '))
        $isAdmin = Test-IsAdmin
        if (-not $isAdmin -and ($IsWindows -or ($PSVersionTable.PSVersion.Major -lt 6))) {
            $sel = Show-Menu -Title "How would you like to install?" -Options @(
                "Install for ALL users (request admin elevation)",
                "Install for CURRENT user only (no elevation)",
                "Skip install (continue without these modules)"
            ) -BackLabel "Quit"
            if ($sel -eq -1) { return $false }
            if ($sel -eq 0) {
                if (Request-AdminElevation -ModulesToInstall $toInstall) {
                    # User elevated and ran installs in another window. Refresh state.
                    $missingRequired = @($missingRequired | Where-Object {
                        -not (Get-Module -ListAvailable -Name $_ | Sort-Object Version -Descending | Select-Object -First 1)
                    })
                    $missingOptional = @($missingOptional | Where-Object {
                        -not (Get-Module -ListAvailable -Name $_ | Sort-Object Version -Descending | Select-Object -First 1)
                    })
                }
            }
            elseif ($sel -eq 2) {
                Write-Warn "Skipping install. Features depending on missing modules will fail."
                $missingRequired = @(); $missingOptional = @()  # treat as 'don't try to install'
            }
            # sel -eq 1: fall through to per-module CurrentUser install
        }
    }

    $allGood = $true
    $smokeResults = New-Object System.Collections.ArrayList

    foreach ($mod in $requiredModules) {
        $modName  = $mod.Name
        $optional = [bool]($mod.Optional)

        # ---- Step 1: Check if installed ----
        $installed = Get-Module -ListAvailable -Name $modName | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $installed) {
            if ($optional) {
                Write-InfoMsg "$modName not installed (optional -- SPO / OneDrive features will be unavailable until installed)."
                [void]$smokeResults.Add([PSCustomObject]@{ Module = $modName; Ok = $false; Reason = "not installed (optional)" })
                continue
            }
            Write-Warn "$modName is not installed. Installing for CurrentUser..."
            try {
                # -SkipPublisherCheck so PSGallery's signing cert change
                # in 2023 doesn't block 5.1 installs (the modules ARE
                # signed by Microsoft; the prior PSGallery publisher is
                # marked untrusted by default on older PowerShellGet).
                Install-Module $modName -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
                Write-Success "$modName installed."
                $installed = Get-Module -ListAvailable -Name $modName | Sort-Object Version -Descending | Select-Object -First 1
            } catch {
                $msg = "$_"
                Write-ErrorMsg "Failed to install $modName : $msg"
                if ($msg -match 'TLS|SSL|underlying connection') {
                    Write-InfoMsg "TLS handshake failed. Run this once, then retry:"
                    Write-Host "    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" -ForegroundColor Cyan
                }
                elseif ($msg -match 'NuGet|PackageProvider') {
                    Write-InfoMsg "Run this once, then retry:"
                    Write-Host "    Install-PackageProvider -Name NuGet -Force -ForceBootstrap" -ForegroundColor Cyan
                }
                elseif ($msg -match 'Administrator|elevation|Access.*denied|Unauthorized') {
                    Write-InfoMsg "This looks like a permission issue. Re-run elevated, or retry with -Scope CurrentUser."
                }
                elseif ($msg -match 'Untrusted|Untrusted repository') {
                    Write-InfoMsg "Run this once, then retry:"
                    Write-Host "    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted" -ForegroundColor Cyan
                }
                $allGood = $false
                [void]$smokeResults.Add([PSCustomObject]@{ Module = $modName; Ok = $false; Reason = "install failed" })
                continue
            }
        }

        # ---- Step 2: Import into session ----
        $loaded = Get-Module -Name $modName
        if (-not $loaded) {
            try {
                Import-Module $modName -ErrorAction Stop -Force
                Write-Success "$modName v$($installed.Version) loaded."
            } catch {
                Write-Warn "Could not import $modName : $_"
                # Try removing and reimporting
                try {
                    Remove-Module $modName -Force -ErrorAction SilentlyContinue
                    Import-Module $modName -ErrorAction Stop -Force
                    Write-Success "$modName v$($installed.Version) loaded (retry)."
                } catch {
                    $emsg = "$_"
                    Write-ErrorMsg "Failed to import $modName : $emsg"
                    if ($emsg -match "GetTokenAsync.*does not have an implementation") {
                        Write-Warn "  This is the known Microsoft.Graph 2.37.0 'GetTokenAsync' bug."
                        Write-Host "  Run option 98 from the main menu and accept the default pin (2.25.0)" -ForegroundColor Yellow
                        Write-Host "  to wipe every Microsoft.Graph.* and install a working version." -ForegroundColor Yellow
                    }
                    elseif ($emsg -match 'Could not load file or assembly.*Microsoft\.Graph') {
                        Write-Warn "  Mixed Microsoft.Graph versions on disk -- the new sibling is looking for"
                        Write-Host "  an old .Core that's not installed. Option 98 will reset to one version." -ForegroundColor Yellow
                    }
                    $allGood = $false
                    [void]$smokeResults.Add([PSCustomObject]@{ Module = $modName; Ok = $false; Reason = "import failed" })
                    continue
                }
            }
        } else {
            Write-InfoMsg "$modName v$($loaded.Version) already loaded."
        }

        # ---- Step 3: Verify test command exists (smoke test) ----
        $smoke = Test-ModuleSmoke -ModuleName $modName -TestCmd $mod.TestCmd
        if (-not $smoke.Ok) {
            Write-ErrorMsg "$modName smoke test failed: $($smoke.Reason)"
            if ($mod.TestCmd) { Write-InfoMsg "Try: Remove-Module $modName; Import-Module $modName" }
            if (-not $optional) { $allGood = $false }
        }
        [void]$smokeResults.Add([PSCustomObject]@{ Module = $modName; Ok = $smoke.Ok; Reason = $smoke.Reason })
    }

    # ---- Summary table so the operator can see at a glance what's OK ----
    Write-Host ""
    Write-Host "  Dependency status:" -ForegroundColor $script:Colors.Info
    foreach ($r in $smokeResults) {
        $mark = if ($r.Ok) { '[OK ]' } else { '[FAIL]' }
        $color = if ($r.Ok) { 'Green' } else { 'Red' }
        Write-Host ("    {0} {1,-50} {2}" -f $mark, $r.Module, $r.Reason) -ForegroundColor $color
    }
    Write-Host ""

    if (-not $allGood) {
        Write-ErrorMsg "Some required modules have issues. The tool may not work correctly."
        Write-InfoMsg "Try closing PowerShell, reopening, and running again."
        Write-Host ""
        $cont = Read-UserInput "Continue anyway? (y/n)"
        if ($cont -notmatch '^[Yy]') { return $false }
    } else {
        Write-Success "All dependencies verified."
    }

    return $true
}

# ============================================================
#  Tenant Mode Selection
# ============================================================

function Select-TenantMode {
    Write-SectionHeader "Tenant Selection"

    # Pre-merge fix: ensure the audit log filename reflects the
    # newly-chosen tenant. Switch-Tenant calls this already; the
    # legacy partner-center picker path didn't, so an operator who
    # switched tenants via this UI ended up writing entries to the
    # prior tenant's session log.
    if (Get-Command Reset-AuditLogPath -ErrorAction SilentlyContinue) { Reset-AuditLogPath }

    # When the operator has registered tenants (Direct OR Partner),
    # offer them at the top of the picker. Each profile carries its
    # own TenantId / primary domain so an MSP can flip between
    # several "Direct" orgs without manually re-typing tenant IDs.
    $tenants = @()
    if (Get-Command Get-Tenants -ErrorAction SilentlyContinue) {
        try { $tenants = @(Get-Tenants) } catch {}
    }

    $opts = @()
    foreach ($t in $tenants) {
        $kind = if ($t.credentialRef) { 'profile' } else { 'direct' }
        $hint = if ($t.primaryDomain) { $t.primaryDomain } else { $t.tenantId }
        $opts += ("{0}  [{1}]  {2}" -f $t.name, $kind, $hint)
    }
    $opts += "My own organization (direct admin, no profile)"
    $opts += "A customer tenant (GDAP partner access)"

    $mode = Show-Menu -Title "Which tenant are you managing?" -Options $opts -BackLabel "Quit"

    if ($mode -eq -1) { return $false }

    # Picked one of the registered tenants from the top of the menu.
    if ($mode -lt $tenants.Count) {
        $target = $tenants[$mode]
        if (Get-Command Switch-Tenant -ErrorAction SilentlyContinue) {
            Switch-Tenant -Name $target.name | Out-Null
            return $true
        }
        # No Switch-Tenant helper loaded: do the minimum so connections target the right place.
        $script:SessionState.TenantMode   = if ($target.credentialRef) { 'Profile' } else { 'Direct' }
        $script:SessionState.TenantId     = $target.tenantId
        $script:SessionState.TenantName   = $target.name
        $script:SessionState.TenantDomain = $target.primaryDomain
        Write-Success ("Tenant: {0}" -f $target.name)
        return $true
    }

    # Index past the registered tenants -> the two synthetic options.
    $synthetic = $mode - $tenants.Count

    if ($synthetic -eq 0) {
        $script:SessionState.TenantMode = "Direct"
        $script:SessionState.TenantId   = $null
        $script:SessionState.TenantName = "Own Tenant"
        Write-Success "Direct tenant mode selected."
        return $true
    }

    # ---- Partner / GDAP ----
    $script:SessionState.TenantMode = "Partner"

    Write-InfoMsg "Signing in to your PARTNER tenant to list customers..."
    Write-Host ""

    $partnerScopes = @("Directory.Read.All")
    $partnerConnected = $false

    # Attempt 1: Interactive browser
    try {
        Connect-MgGraph -Scopes $partnerScopes -NoWelcome -ErrorAction Stop
        $partnerConnected = $true
    } catch {
        Write-Warn "Browser login failed: $_"
        Write-InfoMsg "Trying device code flow instead..."
        Write-Host ""

        # Attempt 2: Device code (works when browser popup is blocked)
        try {
            Connect-MgGraph -Scopes $partnerScopes -NoWelcome -UseDeviceAuthentication -ErrorAction Stop
            $partnerConnected = $true
        } catch {
            Write-ErrorMsg "Device code flow also failed: $_"
        }
    }

    if (-not $partnerConnected) {
        Write-Host ""
        Write-ErrorMsg "Could not authenticate to partner tenant."
        Write-InfoMsg "You can still enter a customer tenant ID manually."
        $manual = Read-UserInput "Tenant ID or domain (or 'back' to cancel)"
        if ($manual -eq 'back' -or [string]::IsNullOrWhiteSpace($manual)) {
            $script:SessionState.TenantMode = "Direct"
            return $false
        }
        $script:SessionState.TenantId = $manual
        $script:SessionState.TenantName = $manual
        $script:SessionState.TenantDomain = $manual
        Write-Success "Will connect to tenant: $manual"
        return $true
    }

    $script:SessionState.PartnerConnected = $true
    $ctx = Get-MgContext
    Write-Success "Signed in as $($ctx.Account)"

    Write-InfoMsg "Fetching customer tenant list..."
    $customers = @()

    # ---- Method 1: GDAP delegatedAdminCustomers API ----
    Write-InfoMsg "Trying GDAP delegated admin customers API..."
    try {
        $gdapResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers" -ErrorAction Stop
        if ($gdapResponse.value -and $gdapResponse.value.Count -gt 0) {
            Write-Success "Found $($gdapResponse.value.Count) GDAP customer(s)."
            foreach ($c in $gdapResponse.value) {
                $custDomain = ""
                if ($c.tenantId) {
                    # Try to get the default domain
                    try {
                        $domainResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers/$($c.tenantId)/serviceManagementDetails" -ErrorAction SilentlyContinue
                    } catch {}
                    $custDomain = $c.tenantId
                }
                $customers += [PSCustomObject]@{
                    DisplayName   = $c.displayName
                    CustomerId    = $c.tenantId
                    DefaultDomain = if ($custDomain) { $custDomain } else { $c.tenantId }
                }
            }
        } else {
            Write-InfoMsg "GDAP API returned 0 customers."
        }
    } catch {
        Write-InfoMsg "GDAP API not available: $_"
    }

    # ---- Method 2: Try delegatedAdminRelationships (more detail) ----
    if ($customers.Count -eq 0) {
        Write-InfoMsg "Trying GDAP relationships API..."
        try {
            $relResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships?`$filter=status eq 'active'" -ErrorAction Stop
            if ($relResponse.value -and $relResponse.value.Count -gt 0) {
                Write-Success "Found $($relResponse.value.Count) active GDAP relationship(s)."
                $seen = @{}
                foreach ($r in $relResponse.value) {
                    $custId = $r.customer.tenantId
                    $custName = $r.customer.displayName
                    if (-not $seen.ContainsKey($custId)) {
                        $seen[$custId] = $true
                        $customers += [PSCustomObject]@{
                            DisplayName   = $custName
                            CustomerId    = $custId
                            DefaultDomain = $custId
                        }
                    }
                }
            } else {
                Write-InfoMsg "Relationships API returned 0 active relationships."
            }
        } catch {
            Write-InfoMsg "Relationships API not available: $_"
        }
    }

    # ---- Method 3: Legacy contracts (DAP) ----
    if ($customers.Count -eq 0) {
        Write-InfoMsg "Trying legacy contracts API (DAP)..."
        try {
            $contracts = @(Get-MgContract -All -ErrorAction Stop)
            if ($contracts.Count -gt 0) {
                Write-Success "Found $($contracts.Count) DAP contract(s)."
                foreach ($c in $contracts) {
                    $customers += [PSCustomObject]@{
                        DisplayName   = $c.DisplayName
                        CustomerId    = $c.CustomerId
                        DefaultDomain = $c.DefaultDomainName
                    }
                }
            }
        } catch {
            Write-InfoMsg "Contracts API not available: $_"
        }
    }

    # ---- Deduplicate ----
    if ($customers.Count -gt 0) {
        $customers = $customers | Sort-Object DisplayName -Unique
    }

    # ---- If no customers found by any method, manual entry ----
    if ($customers.Count -eq 0) {
        Write-Host ""
        Write-Warn "No customer tenants found via any API method."
        Write-InfoMsg "This could mean:"
        Write-InfoMsg "  - No active GDAP relationships"
        Write-InfoMsg "  - Insufficient partner admin roles"
        Write-InfoMsg "  - API permissions not granted"
        Write-Host ""
        Write-InfoMsg "You can enter a customer tenant ID or domain manually."
        $manual = Read-UserInput "Tenant ID or domain (or 'back' to cancel)"
        if ($manual -eq 'back' -or [string]::IsNullOrWhiteSpace($manual)) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            $script:SessionState.PartnerConnected = $false
            $script:SessionState.TenantMode = "Direct"
            return $false
        }
        $script:SessionState.TenantId = $manual
        $script:SessionState.TenantName = $manual
        $script:SessionState.TenantDomain = $manual
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $script:SessionState.MgGraph = $false; $script:SessionState.PartnerConnected = $false
        Write-Success "Will connect to tenant: $manual"
        return $true
    }

    # ---- Customer picker ----
    if ($customers.Count -gt 20) {
        $searchInput = Read-UserInput "Search customer by name (or 'all' to list all)"
        if ($searchInput -ne 'all') {
            $customers = @($customers | Where-Object { $_.DisplayName -like "*$searchInput*" })
            if ($customers.Count -eq 0) {
                Write-ErrorMsg "No customers matching '$searchInput'."
                Disconnect-MgGraph -ErrorAction SilentlyContinue
                $script:SessionState.TenantMode = "Direct"
                return $false
            }
        }
    }

    $custLabels = $customers | ForEach-Object { "$($_.DisplayName)  ($($_.DefaultDomain))" }
    $sel = Show-Menu -Title "Select Customer Tenant ($($customers.Count) found)" -Options $custLabels -BackLabel "Cancel"
    if ($sel -eq -1) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $script:SessionState.PartnerConnected = $false
        $script:SessionState.TenantMode = "Direct"
        return $false
    }

    $selected = $customers[$sel]
    $script:SessionState.TenantId = $selected.CustomerId
    $script:SessionState.TenantName = $selected.DisplayName
    $script:SessionState.TenantDomain = $selected.DefaultDomain

    Write-Host ""
    Write-Success "Selected: $($selected.DisplayName)"
    Write-StatusLine "Tenant ID" $selected.CustomerId "White"
    Write-StatusLine "Domain" $selected.DefaultDomain "White"
    Write-Host ""

    Disconnect-MgGraph -ErrorAction SilentlyContinue
    $script:SessionState.MgGraph = $false; $script:SessionState.PartnerConnected = $false
    return $true
}

# ============================================================
#  Startup Session Cleanup
# ============================================================

function Clear-StartupSession {
    <#
        Ensure the tool starts with a clean slate.
        Runs at every launch to defeat two scenarios:
          1. User launched inside a PowerShell host that already had an
             active Connect-MgGraph / Connect-ExchangeOnline / Connect-IPPSSession
             — those sessions would otherwise silently leak into this tool.
          2. A stale on-disk token cache from a previous run causes a
             surprise auto-login to the wrong tenant/account.
        Does NOT touch the shared MSAL cache under
        $env:LOCALAPPDATA\.IdentityService — that is used by other Microsoft
        apps on this machine and clearing it would sign the user out of them.
    #>
    Write-SectionHeader "Clearing Previous Session"

    # ---- In-process Graph ----
    if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
        try {
            $existing = $null
            if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
                $existing = Get-MgContext -ErrorAction SilentlyContinue
            }
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            if ($existing) { Write-InfoMsg "Stale Graph session ($($existing.Account)) cleared." }
        } catch {}
    }

    # ---- In-process Exchange Online / IPPSSession ----
    if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
        try {
            $exoConns = $null
            if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
                $exoConns = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
            }
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            if ($exoConns -and $exoConns.Count -gt 0) {
                Write-InfoMsg "$($exoConns.Count) stale EXO/SCC session(s) cleared."
            }
        } catch {}
    }

    # ---- On-disk Graph SDK token cache ----
    # The Microsoft.Graph PS SDK persists tokens here when ContextScope=CurrentUser.
    # Wiping it forces a fresh interactive login on the next Connect-Graph call.
    if ($env:USERPROFILE) {
        $cachePath = Join-Path $env:USERPROFILE ".graph"
        if (Test-Path $cachePath) {
            try {
                Remove-Item -Path $cachePath -Recurse -Force -ErrorAction Stop
                Write-InfoMsg "Graph SDK token cache cleared ($cachePath)."
            } catch {
                Write-Warn "Could not clear Graph token cache: $_"
            }
        }
    }

    # ---- Reset in-memory state (defensive — module-load already initialized it) ----
    $script:SessionState.MgGraph            = $false
    $script:SessionState.ExchangeOnline     = $false
    $script:SessionState.ComplianceCenter   = $false
    $script:SessionState.SharePointOnline   = $false
    $script:SessionState.SharePointAdminUrl = $null
    $script:SessionState.PartnerConnected   = $false
    $script:SessionState.TenantMode         = "Direct"
    $script:SessionState.TenantId           = $null
    $script:SessionState.TenantName         = $null
    $script:SessionState.TenantDomain       = $null

    Write-Success "Starting with a clean session."
}

# ============================================================
#  Full Session Cleanup (for tenant switch)
# ============================================================

function Reset-AllSessions {
    <# Disconnects every service and resets all state. Used when switching tenants. #>
    Write-SectionHeader "Cleaning Up All Sessions"

    if ($script:SessionState.MgGraph) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
        Write-Success "Microsoft Graph disconnected."
    }
    if ($script:SessionState.ExchangeOnline -or $script:SessionState.ComplianceCenter) {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
        Write-Success "Exchange Online / SCC disconnected."
    }
    if ($script:SessionState.SharePointOnline) {
        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
        Write-Success "SharePoint Online disconnected."
    }

    $script:SessionState.MgGraph            = $false
    $script:SessionState.ExchangeOnline     = $false
    $script:SessionState.ComplianceCenter   = $false
    $script:SessionState.SharePointOnline   = $false
    $script:SessionState.SharePointAdminUrl = $null
    $script:SessionState.PartnerConnected   = $false
    $script:SessionState.TenantId           = $null
    $script:SessionState.TenantName         = $null
    $script:SessionState.TenantDomain       = $null
    $script:SessionState.TenantMode         = "Direct"

    Write-Success "All sessions and tenant context cleared."
}

# ============================================================
#  Service Connections
# ============================================================

function Connect-Graph {
    if ($script:SessionState.MgGraph) { Write-InfoMsg "Microsoft Graph already connected."; return $true }

    $targetLabel = if ($script:SessionState.TenantMode -eq "Partner") { "$($script:SessionState.TenantName) (GDAP)" } else { "own tenant" }
    Write-InfoMsg "Connecting to Microsoft Graph ($targetLabel)..."

    $params = @{ Scopes = $script:MgScopes; NoWelcome = $true }
    if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantId) { $params["TenantId"] = $script:SessionState.TenantId }

    # Attempt 1: Interactive browser
    try {
        Connect-MgGraph @params -ErrorAction Stop
        $script:SessionState.MgGraph = $true
        $ctx = Get-MgContext
        Write-Success "Microsoft Graph connected as $($ctx.Account)"
        Verify-GraphScopes
        return $true
    } catch {
        Write-Warn "Browser login failed: $_"
    }

    # Attempt 2: Device code
    Write-InfoMsg "Trying device code flow..."
    try {
        $params["UseDeviceAuthentication"] = $true
        Connect-MgGraph @params -ErrorAction Stop
        $script:SessionState.MgGraph = $true
        $ctx = Get-MgContext
        Write-Success "Microsoft Graph connected via device code as $($ctx.Account)"
        Verify-GraphScopes
        return $true
    } catch {
        Write-ErrorMsg "All Graph connection methods failed: $_"
        return $false
    }
}

function Verify-GraphScopes {
    $ctx = Get-MgContext
    if ($ctx.Scopes) {
        $missing = @()
        foreach ($s in $script:MgScopes) {
            if ($ctx.Scopes -notcontains $s) { $missing += $s }
        }
        if ($missing.Count -gt 0) {
            Write-Warn "Missing scopes: $($missing -join ', ')"
            Write-InfoMsg "An admin may need to grant consent in Entra portal."
        }
    }
}

function Reconnect-GraphWithConsent {
    <# Disconnects Graph, clears cached token, and reconnects to trigger fresh consent. #>
    Write-InfoMsg "Disconnecting current Graph session..."
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $script:SessionState.MgGraph = $false

    # Clear cached Graph context to force fresh login
    try {
        $cachePath = Join-Path $env:USERPROFILE ".graph"
        if (Test-Path $cachePath) {
            Remove-Item -Path $cachePath -Recurse -Force -ErrorAction SilentlyContinue
            Write-InfoMsg "Graph token cache cleared."
        }
    } catch {}

    Write-InfoMsg "Reconnecting (a browser window will open for consent)..."
    Write-Warn "If you are an admin, check 'Consent on behalf of your organization' in the prompt."
    Write-Host ""

    try {
        $params = @{ Scopes = $script:MgScopes; NoWelcome = $true }
        if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantId) { $params["TenantId"] = $script:SessionState.TenantId }
        Connect-MgGraph @params -ErrorAction Stop
        $script:SessionState.MgGraph = $true
        $ctx = Get-MgContext
        Write-Success "Reconnected as $($ctx.Account)"
        Write-InfoMsg "Granted scopes: $($ctx.Scopes -join ', ')"
        return $true
    } catch {
        Write-ErrorMsg "Reconnect failed: $_"
        $script:SessionState.MgGraph = $false
        return $false
    }
}

function Connect-EXO {
    if ($script:SessionState.ExchangeOnline) { Write-InfoMsg "Exchange Online already connected."; return $true }

    $targetLabel = if ($script:SessionState.TenantMode -eq "Partner") { "$($script:SessionState.TenantName) (GDAP)" } else { "own tenant" }
    Write-InfoMsg "Connecting to Exchange Online ($targetLabel)..."

    # Disable MSAL WAM broker to prevent DLL version conflicts with MS Graph
    # This forces EXO to use standard browser auth instead
    [System.Environment]::SetEnvironmentVariable("MSAL_BROKER_ENABLED", "0", "Process")

    $exoParams = @{ ShowBanner = $false }
    if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantDomain) {
        $exoParams["DelegatedOrganization"] = $script:SessionState.TenantDomain
    }

    try {
        Connect-ExchangeOnline @exoParams -ErrorAction Stop
        $script:SessionState.ExchangeOnline = $true
        Write-Success "Exchange Online connected."
        return $true
    } catch {
        Write-ErrorMsg "Exchange Online connection failed: $_"
        return $false
    }
}

function Connect-SCC {
    if ($script:SessionState.ComplianceCenter) { Write-InfoMsg "SCC already connected."; return $true }

    $targetLabel = if ($script:SessionState.TenantMode -eq "Partner") { "$($script:SessionState.TenantName) (GDAP)" } else { "own tenant" }
    Write-InfoMsg "Connecting to Security & Compliance ($targetLabel)..."

    # Disable MSAL WAM broker (same conflict as EXO)
    [System.Environment]::SetEnvironmentVariable("MSAL_BROKER_ENABLED", "0", "Process")

    $sccParams = @{ ShowBanner = $false; EnableSearchOnlySession = $true }
    if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantDomain) {
        $sccParams["DelegatedOrganization"] = $script:SessionState.TenantDomain
    }

    try {
        Connect-IPPSSession @sccParams -ErrorAction Stop
        $script:SessionState.ComplianceCenter = $true
        Write-Success "SCC connected (search session)."
        return $true
    } catch {
        Write-Warn "SCC search session failed: $_"
    }

    # Fallback without search session
    try {
        $fallback = @{ ShowBanner = $false }
        if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantDomain) { $fallback["DelegatedOrganization"] = $script:SessionState.TenantDomain }
        Connect-IPPSSession @fallback -ErrorAction Stop
        $script:SessionState.ComplianceCenter = $true
        Write-Success "SCC connected (basic)."
        Write-Warn "Some search operations may require restart with search session."
        return $true
    } catch {
        Write-ErrorMsg "All SCC connection methods failed: $_"
        return $false
    }
}

# ============================================================
#  SharePoint Online connection (Phase 3)
#
#  Requires SharePoint Administrator role. Connect-SPOService
#  uses the tenant admin URL (e.g. https://contoso-admin.sharepoint.com)
#  -- we cache the last-used URL per-tenant in a small JSON state
#  file so the operator isn't prompted on every reconnect.
# ============================================================

function Get-StateDirectory {
    $base = $null
    if ($env:LOCALAPPDATA) { $base = Join-Path $env:LOCALAPPDATA 'M365Manager\state' }
    elseif ($env:HOME)     { $base = Join-Path $env:HOME '.m365manager/state' }
    else                   { $base = Join-Path (Get-Location).Path 'state' }
    if (-not (Test-Path -LiteralPath $base)) {
        try { New-Item -ItemType Directory -Path $base -Force | Out-Null } catch { return $null }
        if (-not $env:LOCALAPPDATA -and (Get-Command chmod -ErrorAction SilentlyContinue)) {
            try { & chmod 700 $base 2>$null | Out-Null } catch {}
        }
    }
    return $base
}

function Get-SPOAdminUrlCache {
    <#
        Returns a hashtable of tenantKey -> [string[]] of cached admin URLs.
        Reads spo-tenants.json. Backward-compatible with the legacy
        format where each tenantKey held a single string -- old strings
        are migrated to a 1-element array on read. The on-disk file is
        rewritten in the new format on the next Add/Remove.
    #>
    $dir = Get-StateDirectory
    if (-not $dir) { return @{} }
    $p = Join-Path $dir 'spo-tenants.json'
    if (-not (Test-Path -LiteralPath $p)) { return @{} }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json
        $h = @{}
        foreach ($prop in $raw.PSObject.Properties) {
            $val = $prop.Value
            $list = @()
            if ($null -eq $val)        { $list = @() }
            elseif ($val -is [string]) { $list = @([string]$val) }
            elseif ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) {
                foreach ($v in $val) { if ($v) { $list += [string]$v } }
            } else {
                $list = @([string]$val)
            }
            # Dedup, preserve order, trim
            $seen = @{}; $out = @()
            foreach ($u in $list) {
                $k = $u.Trim().TrimEnd('/').ToLowerInvariant()
                if ($k -and -not $seen.ContainsKey($k)) { $seen[$k] = $true; $out += $u.Trim().TrimEnd('/') }
            }
            $h[$prop.Name] = $out
        }
        return $h
    } catch { return @{} }
}

function Save-SPOAdminUrlCache {
    param([hashtable]$Cache)
    $dir = Get-StateDirectory
    if (-not $dir) { return }
    $p = Join-Path $dir 'spo-tenants.json'
    try { ($Cache | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $p -Encoding UTF8 -Force } catch {}
}

function Add-SPOAdminUrlCache {
    <#
        Append an admin URL to the cache for $TenantKey. Idempotent
        (dedup, case-insensitive). The most-recently-used URL is moved
        to the front of the list so the picker shows it first.
    #>
    param([string]$TenantKey, [string]$AdminUrl)
    if ([string]::IsNullOrWhiteSpace($AdminUrl)) { return }
    $url = $AdminUrl.Trim().TrimEnd('/')
    $cache = Get-SPOAdminUrlCache
    $existing = @()
    if ($cache.ContainsKey($TenantKey)) { $existing = @($cache[$TenantKey]) }
    # Remove any case-insensitive duplicate, then prepend the new one.
    $existing = @($existing | Where-Object { $_.ToLowerInvariant() -ne $url.ToLowerInvariant() })
    $merged = @($url) + $existing
    $cache[$TenantKey] = $merged
    Save-SPOAdminUrlCache -Cache $cache
}

function Remove-SPOAdminUrlCache {
    <#
        Remove a single URL from the tenant's cached list. If $AdminUrl
        is omitted, removes ALL URLs for the tenant.
    #>
    param([string]$TenantKey, [string]$AdminUrl)
    $cache = Get-SPOAdminUrlCache
    if (-not $cache.ContainsKey($TenantKey)) { return }
    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        $cache.Remove($TenantKey) | Out-Null
    } else {
        $url = $AdminUrl.Trim().TrimEnd('/')
        $cache[$TenantKey] = @($cache[$TenantKey] | Where-Object { $_.ToLowerInvariant() -ne $url.ToLowerInvariant() })
        if ($cache[$TenantKey].Count -eq 0) { $cache.Remove($TenantKey) | Out-Null }
    }
    Save-SPOAdminUrlCache -Cache $cache
}

# Compatibility shim for any external caller that still uses the old
# single-URL setter. Routes through Add-SPOAdminUrlCache so the new
# multi-URL list grows naturally.
function Set-SPOAdminUrlCache {
    param([string]$TenantKey, [string]$AdminUrl)
    Add-SPOAdminUrlCache -TenantKey $TenantKey -AdminUrl $AdminUrl
}

function Get-SPOAdminUrlSuggestions {
    <#
        Build a deduped, ordered list of plausible SharePoint admin
        URLs for the active tenant. Sources, in priority order:
          1. Cached URL for this tenant key (what we used last time).
          2. URL derived from TenantDomain (foo.onmicrosoft.com ->
             https://foo-admin.sharepoint.com).
          3. URLs derived from any verified domain on the tenant
             whose short name doesn't look like a custom vanity
             (Graph: organization.verifiedDomains, filter to
             *.onmicrosoft.com).
          4. URL derived from the current Mg-Graph account's domain.
        Returns @() of strings.
    #>
    $out = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    function Add-Url($u) {
        if (-not $u) { return }
        $u = ([string]$u).Trim().TrimEnd('/')
        if (-not $u) { return }
        if ($u -notmatch '^https?://') { return }
        if ($seen.ContainsKey($u.ToLowerInvariant())) { return }
        $seen[$u.ToLowerInvariant()] = $true
        $out.Add($u)
    }
    function Derive-FromDomain($dom) {
        if (-not $dom) { return $null }
        if ($dom -notlike '*.onmicrosoft.com') { return $null }
        $short = ($dom -split '\.')[0]
        if (-not $short) { return $null }
        return "https://$short-admin.sharepoint.com"
    }

    $tenantKey = if ($script:SessionState.TenantDomain) { $script:SessionState.TenantDomain }
                 elseif ($script:SessionState.TenantId) { $script:SessionState.TenantId }
                 else { 'default' }
    $cache = Get-SPOAdminUrlCache
    # Surface every cached URL for this tenant, most-recent first.
    if ($cache.ContainsKey($tenantKey)) {
        foreach ($cu in @($cache[$tenantKey])) { Add-Url $cu }
    }

    Add-Url (Derive-FromDomain $script:SessionState.TenantDomain)

    # Verified domains via Graph (only if connected)
    if ($script:SessionState.MgGraph -and (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        try {
            $org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
            foreach ($o in $org.value) {
                foreach ($vd in @($o.verifiedDomains)) {
                    Add-Url (Derive-FromDomain $vd.name)
                }
            }
        } catch {}
    }

    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $ctx = Get-MgContext
            if ($ctx -and $ctx.Account -and $ctx.Account -match '@(.+)$') {
                Add-Url (Derive-FromDomain $Matches[1])
            }
        } catch {}
    }

    # Return a plain string[] -- using `,@($out)` here wrapped the list
    # in an extra array on the caller side, so Connect-SPOService later
    # got an Object[] for -Url and threw "Cannot convert ... Specified
    # method is not supported."
    return [string[]]$out
}

function Connect-SPO {
    <#
        Connect-SPOService wrapper. Always prompts the operator to
        confirm which SharePoint admin URL to use (never silently
        reuses the cached value) but pre-fills the picker with the
        best guesses derived from the cache, the tenant domain, and
        Graph's verified-domain list. Pick "Enter manually" to type
        a different URL. The chosen URL is then re-cached.
    #>
    if ($script:SessionState.SharePointOnline) { Write-InfoMsg "SharePoint Online already connected."; return $true }

    if (-not (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
        Write-ErrorMsg "Microsoft.Online.SharePoint.PowerShell module not loaded."
        Write-InfoMsg "Install with: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser"
        return $false
    }
    Write-InfoMsg "SharePoint Online connection requires the SharePoint Administrator role."

    $tenantKey = if ($script:SessionState.TenantDomain) { $script:SessionState.TenantDomain }
                 elseif ($script:SessionState.TenantId) { $script:SessionState.TenantId }
                 else { 'default' }

    $suggestions = @(Get-SPOAdminUrlSuggestions)
    $adminUrl = $null

    while (-not $adminUrl) {
        $cache = Get-SPOAdminUrlCache
        $cachedForTenant = @()
        if ($cache.ContainsKey($tenantKey)) { $cachedForTenant = @($cache[$tenantKey]) }

        # Build menu: each suggestion is a row, then manual entry, then
        # (if anything cached) a per-URL remove option, then clear-all.
        $opts = @()
        for ($i = 0; $i -lt $suggestions.Count; $i++) {
            $tag = ''
            if ($cachedForTenant -contains $suggestions[$i]) { $tag = '  [saved]' }
            $opts += ("{0}{1}" -f $suggestions[$i], $tag)
        }
        $opts += "Enter manually (will be saved for next time)..."
        $removeIdx = -1; $clearAllIdx = -1
        if ($cachedForTenant.Count -gt 0) {
            $removeIdx = $opts.Count;   $opts += "Remove one saved URL..."
            $clearAllIdx = $opts.Count; $opts += ("Clear ALL saved URLs for this tenant ({0})" -f $cachedForTenant.Count)
        }

        $sel = Show-Menu -Title ("SharePoint admin URL ({0} option(s))" -f $suggestions.Count) -Options $opts -BackLabel "Cancel"
        if ($sel -eq -1) { Write-Warn "Cancelled."; return $false }

        if ($sel -lt $suggestions.Count) {
            $adminUrl = $suggestions[$sel]
        }
        elseif ($sel -eq $suggestions.Count) {
            $hint = ""
            if ($suggestions.Count -gt 0) { $hint = " (default: $($suggestions[0]))" }
            $entered = Read-UserInput "SharePoint admin URL$hint"
            if ([string]::IsNullOrWhiteSpace($entered)) {
                if ($suggestions.Count -gt 0) { $adminUrl = $suggestions[0] }
                else { Write-Warn "No URL entered."; continue }
            } else {
                $entered = $entered.Trim().TrimEnd('/')
                if ($entered -notmatch '^https?://') { $entered = "https://$entered" }
                $adminUrl = $entered
            }
        }
        elseif ($sel -eq $removeIdx) {
            # Pick a specific cached URL to drop.
            $rsel = Show-Menu -Title "Remove which saved URL?" -Options $cachedForTenant -BackLabel "Cancel"
            if ($rsel -ne -1) {
                Remove-SPOAdminUrlCache -TenantKey $tenantKey -AdminUrl $cachedForTenant[$rsel]
                Write-Success ("Removed: {0}" -f $cachedForTenant[$rsel])
                $suggestions = @(Get-SPOAdminUrlSuggestions)
            }
        }
        elseif ($sel -eq $clearAllIdx) {
            Remove-SPOAdminUrlCache -TenantKey $tenantKey
            Write-Success "All saved SharePoint URLs cleared for tenant '$tenantKey'."
            $suggestions = @(Get-SPOAdminUrlSuggestions)
        }
    }

    Write-InfoMsg "Connecting to SharePoint Online ($adminUrl)..."
    try {
        Connect-SPOService -Url ([string]$adminUrl) -ErrorAction Stop
        $script:SessionState.SharePointOnline   = $true
        $script:SessionState.SharePointAdminUrl = $adminUrl
        Add-SPOAdminUrlCache -TenantKey $tenantKey -AdminUrl $adminUrl
        Write-Success "SharePoint Online connected."
        return $true
    } catch {
        Write-ErrorMsg "SPO connection failed: $_"
        return $false
    }
}

# ============================================================
#  Per-task connection sets
# ============================================================

function Connect-ForTask {
    param(
        [ValidateSet(
            "Onboard","Offboard","License","Archive",
            "SecurityGroup","DistributionList","SharedMailbox","CalendarAccess",
            "UserProfile","Report","eDiscovery","GroupManager",
            "OneDrive","SharePoint","Teams","GuestUsers"
        )]
        [string]$Task
    )

    $map = @{
        Onboard          = @("EXO","Graph")
        Offboard         = @("EXO","Graph","SPO")
        License          = @("Graph")
        Archive          = @("EXO","Graph")
        SecurityGroup    = @("Graph")
        DistributionList = @("EXO","Graph")
        SharedMailbox    = @("EXO","Graph")
        CalendarAccess   = @("EXO","Graph")
        UserProfile      = @("Graph")
        Report           = @("EXO","Graph")
        eDiscovery       = @("SCC","Graph")
        GroupManager     = @("EXO","Graph")
        OneDrive         = @("SPO","Graph")
        SharePoint       = @("SPO","Graph")
        Teams            = @("Graph")
        GuestUsers       = @("Graph","SPO")
    }

    $services = $map[$Task]
    $needed = @()
    foreach ($svc in $services) {
        switch ($svc) {
            "Graph" { if (-not $script:SessionState.MgGraph)            { $needed += $svc } }
            "EXO"   { if (-not $script:SessionState.ExchangeOnline)     { $needed += $svc } }
            "SCC"   { if (-not $script:SessionState.ComplianceCenter)   { $needed += $svc } }
            "SPO"   { if (-not $script:SessionState.SharePointOnline)   { $needed += $svc } }
        }
    }

    if ($needed.Count -eq 0) { Write-InfoMsg "All required services connected."; return $true }

    Write-InfoMsg "Requires: $($services -join ', '). Connecting: $($needed -join ', ')"
    Write-Host ""

    foreach ($svc in $services) {
        switch ($svc) {
            "Graph" { if (-not (Connect-Graph)) { return $false } }
            "EXO"   { if (-not (Connect-EXO))   { return $false } }
            "SCC"   { if (-not (Connect-SCC))   { return $false } }
            "SPO"   {
                if (-not (Connect-SPO)) {
                    # SPO is optional for some tasks (Offboard / GuestUsers still
                    # complete without it; OneDrive / SharePoint can't). Treat
                    # the failure as soft and let the caller's per-step
                    # Get-Command guards handle missing functionality.
                    if ($Task -eq 'OneDrive' -or $Task -eq 'SharePoint') { return $false }
                    Write-Warn "SPO connect failed; continuing without SharePoint-dependent steps."
                }
            }
        }
    }
    return $true
}

# ============================================================
#  Disconnect (for quit)
# ============================================================

function Disconnect-AllSessions {
    Write-SectionHeader "Disconnecting Sessions"

    if ($script:SessionState.MgGraph) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue; Write-Success "Graph disconnected." } catch {}
    }
    if ($script:SessionState.ExchangeOnline -or $script:SessionState.ComplianceCenter) {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue; Write-Success "EXO/SCC disconnected." } catch {}
    }
    if ($script:SessionState.SharePointOnline) {
        try { Disconnect-SPOService -ErrorAction SilentlyContinue; Write-Success "SPO disconnected." } catch {}
    }

    $script:SessionState.MgGraph            = $false
    $script:SessionState.ExchangeOnline     = $false
    $script:SessionState.ComplianceCenter   = $false
    $script:SessionState.SharePointOnline   = $false
    $script:SessionState.SharePointAdminUrl = $null
    $script:SessionState.PartnerConnected   = $false
    Write-Success "All sessions cleared."
}

function Get-TenantDisplayString {
    if ($script:SessionState.TenantMode -eq "Partner") { return "GDAP: $($script:SessionState.TenantName)" }
    return "Direct (own tenant)"
}
