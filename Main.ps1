# ============================================================
#  Main.ps1 - M365 Administration Tool - Entry Point
# ============================================================

# ---- Find the folder containing all .ps1 files ----
# Try every known method; OneDrive/spaces/renamed folders break some of them
$ScriptRoot = $null

# Method 1: $PSScriptRoot (works when launched via -File)
if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot "UI.ps1"))) {
    $ScriptRoot = $PSScriptRoot
}

# Method 2: Environment variable set by Launch.bat
if (-not $ScriptRoot) {
    $envRoot = $env:M365ADMIN_ROOT
    if ($envRoot) {
        $envRoot = $envRoot.TrimEnd('\')
        if (Test-Path (Join-Path $envRoot "UI.ps1")) { $ScriptRoot = $envRoot }
    }
}

# Method 3: $MyInvocation path
if (-not $ScriptRoot) {
    $invPath = $MyInvocation.MyCommand.Path
    if ($invPath) {
        $dir = Split-Path -Parent $invPath
        if (Test-Path (Join-Path $dir "UI.ps1")) { $ScriptRoot = $dir }
    }
}

# Method 4: $MyInvocation.InvocationName (handles relative paths)
if (-not $ScriptRoot) {
    $invName = $MyInvocation.InvocationName
    if ($invName -and $invName -ne '&') {
        try {
            $resolved = (Resolve-Path $invName -ErrorAction Stop).Path
            $dir = Split-Path -Parent $resolved
            if (Test-Path (Join-Path $dir "UI.ps1")) { $ScriptRoot = $dir }
        } catch {}
    }
}

# Method 5: Current working directory
if (-not $ScriptRoot) {
    $cwd = (Get-Location).Path
    if (Test-Path (Join-Path $cwd "UI.ps1")) { $ScriptRoot = $cwd }
}

# ---- Fail if nothing worked ----
if (-not $ScriptRoot) {
    Write-Host "" -ForegroundColor Red
    Write-Host "  ERROR: Cannot locate tool files (UI.ps1 not found)." -ForegroundColor Red
    Write-Host "" -ForegroundColor Yellow
    Write-Host "  Searched in:" -ForegroundColor Yellow
    Write-Host "    PSScriptRoot : $PSScriptRoot" -ForegroundColor Gray
    Write-Host "    ENV          : $env:M365ADMIN_ROOT" -ForegroundColor Gray
    Write-Host "    InvocationPath: $($MyInvocation.MyCommand.Path)" -ForegroundColor Gray
    Write-Host "    WorkingDir   : $((Get-Location).Path)" -ForegroundColor Gray
    Write-Host "" -ForegroundColor Yellow
    Write-Host "  FIX: Use Launch.bat to start the tool." -ForegroundColor Yellow
    Write-Host "       Or cd into the tool folder first, then run .\Main.ps1" -ForegroundColor Yellow
    Write-Host "" -ForegroundColor Gray
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# ---- Strip Mark-of-the-Web on all .ps1 files in the folder.
# Files extracted from a downloaded ZIP get a Zone.Identifier NTFS
# stream that makes PowerShell refuse to dot-source them, even with
# -ExecutionPolicy Bypass. Unblock-File removes the tag silently so
# the dot-source loop below works on freshly-extracted copies.
try {
    Get-ChildItem -Path $ScriptRoot -Filter *.ps1 -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
} catch {}

# ---- Load all modules ----
$loadErrors = @()
$modules = @(
    "UI.ps1","Auth.ps1","Audit.ps1","Preview.ps1","Templates.ps1",
    "Notifications.ps1","TenantRegistry.ps1","TenantSwitch.ps1","TenantOverrides.ps1",
    "Onboard.ps1","BulkOnboard.ps1","TemplateGenerator.ps1","Offboard.ps1","BulkOffboard.ps1",
    "License.ps1","Archive.ps1","SecurityGroup.ps1","DistributionList.ps1",
    "SharedMailbox.ps1","CalendarAccess.ps1","UserProfile.ps1",
    "Reports.ps1","eDiscovery.ps1","GroupManager.ps1",
    "AuditViewer.ps1","Undo.ps1","SignInLookup.ps1","UnifiedAuditLog.ps1",
    "MFAManager.ps1","OneDriveManager.ps1","TeamsManager.ps1","SharePoint.ps1",
    "GuestUsers.ps1","LicenseOptimizer.ps1","Scheduler.ps1","BreakGlass.ps1",
    "MSPReports.ps1","MSPDashboard.ps1",
    "IncidentResponse.ps1","IncidentRegistry.ps1","IncidentBulk.ps1","IncidentTriggers.ps1",
    "AICostTracker.ps1","AISessionStore.ps1","AIUx.ps1","AIToolDispatch.ps1","AIPlanner.ps1","AIAssistant.ps1"
)

foreach ($mod in $modules) {
    $modPath = Join-Path $ScriptRoot $mod
    if (Test-Path $modPath) {
        # Use $Error capture in addition to try/catch: some dot-source
        # failures on PS 5.1 are non-terminating and bypass catch.
        $errBefore = $Error.Count
        try {
            . "$modPath"
        } catch {
            $loadErrors += "  $mod : $_"
        }
        $errAfter = $Error.Count
        if ($errAfter -gt $errBefore) {
            # Capture the new errors that appeared during the dot-source.
            $new = @($Error[0..($errAfter - $errBefore - 1)])
            foreach ($e in $new) {
                $loadErrors += "  $mod : $($e.Exception.Message)"
            }
        }
    }
    # AIAssistant.ps1 is optional, others are not
    elseif ($mod -ne "AIAssistant.ps1") {
        $loadErrors += "  $mod : FILE NOT FOUND at $modPath"
    }
}

if ($loadErrors.Count -gt 0) {
    Write-Host "" -ForegroundColor Red
    Write-Host "  WARNING: Some modules failed to load:" -ForegroundColor Yellow
    foreach ($e in $loadErrors) { Write-Host $e -ForegroundColor Red }
    Write-Host "" -ForegroundColor Gray
}

# ---- Verify the critical functions every flow expects are present.
# If a module dot-sourced silently (non-terminating error), the user
# would otherwise see "term X not recognized" deep inside a flow
# (e.g. Get-OnboardTemplates from Onboard.ps1). Print a single clear
# diagnostic upfront and continue with what we have.
$criticalFunctions = @(
    @{ Name='Initialize-UI';           From='UI.ps1' },
    @{ Name='Show-Menu';               From='UI.ps1' },
    @{ Name='Assert-ModulesInstalled'; From='Auth.ps1' },
    @{ Name='Connect-Graph';           From='Auth.ps1' },
    @{ Name='Get-OnboardTemplates';    From='Templates.ps1' },
    @{ Name='Start-TemplateGeneratorMenu'; From='TemplateGenerator.ps1' },
    @{ Name='Invoke-Action';           From='Audit.ps1' },
    @{ Name='Get-PreviewMode';         From='Preview.ps1' }
)
$missingCritical = @()
foreach ($c in $criticalFunctions) {
    if (-not (Get-Command $c.Name -ErrorAction SilentlyContinue)) {
        $missingCritical += "    [x] $($c.Name)  (expected from $($c.From))"
    }
}
if ($missingCritical.Count -gt 0) {
    Write-Host ""
    Write-Host "  WARNING: Critical helper functions are missing after load:" -ForegroundColor Yellow
    foreach ($m in $missingCritical) { Write-Host $m -ForegroundColor Red }
    Write-Host ""
    Write-Host "  This usually means one or more .ps1 files failed to dot-source." -ForegroundColor Yellow
    Write-Host "  Try the Mark-of-the-Web fix below (run from this folder):" -ForegroundColor Yellow
    Write-Host "    Get-ChildItem -Recurse | Unblock-File" -ForegroundColor Cyan
    Write-Host ""
}

# ---- Hard-stop if the UI module didn't load. Everything below depends
# on Initialize-UI / Write-Banner / Show-Menu: if those aren't defined
# we'd produce a wall of "term not recognized" errors with no fix path.
# The most common cause is Mark-of-the-Web blocking on a ZIP-extracted
# copy that Windows Defender hasn't been told to trust yet.
if (-not (Get-Command Initialize-UI -ErrorAction SilentlyContinue)) {
    Write-Host ""
    Write-Host "  ERROR: UI module did not load (Initialize-UI undefined)." -ForegroundColor Red
    Write-Host ""
    Write-Host "  Most likely: files extracted from a downloaded ZIP are still" -ForegroundColor Yellow
    Write-Host "  tagged with Mark-of-the-Web. Fix with ONE of:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    1.  Right-click M365-Manager.zip BEFORE extracting, choose" -ForegroundColor White
    Write-Host "        Properties, tick 'Unblock', then re-extract." -ForegroundColor White
    Write-Host ""
    Write-Host "    2.  Run this in an elevated PowerShell window from this folder:" -ForegroundColor White
    Write-Host "          Get-ChildItem -Recurse | Unblock-File" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    3.  Move the folder out of Downloads (e.g. C:\\Tools\\M365-Manager)" -ForegroundColor White
    Write-Host "        and run from there." -ForegroundColor White
    Write-Host ""
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# ---- Main Application ----
function Start-M365Admin {
    Initialize-UI
    Write-Banner

    if (-not (Assert-ModulesInstalled)) {
        Write-ErrorMsg "Dependency check failed. Exiting."
        Pause-ForUser
        return
    }
    Write-Host ""

    # Wipe any leftover in-memory connections and on-disk token cache from a
    # prior run before showing tenant selection — guarantees a clean start.
    Clear-StartupSession
    Write-Host ""

    # ---- Operating mode picker ----
    Write-SectionHeader "Operating Mode"
    $modeSel = Show-Menu -Title "Run this session in" -Options @(
        "LIVE     -- changes WILL be applied to the tenant",
        "PREVIEW  -- dry-run, no tenant changes"
    ) -BackLabel "Quit"
    if ($modeSel -eq -1) {
        Write-Host ""; Write-Host "  Goodbye!" -ForegroundColor $Global:M365Colors.Title; return
    }
    Set-PreviewMode -Enabled ($modeSel -eq 1)
    if (Get-PreviewMode) {
        Write-Warn "PREVIEW mode -- mutating cmdlets will be logged but not executed."
    } else {
        Write-InfoMsg "LIVE mode -- mutations will be applied."
    }
    Write-Host ""

    if (-not (Select-TenantMode)) {
        Write-Host ""
        Write-Host "  Goodbye!" -ForegroundColor $Global:M365Colors.Title
        return
    }

    # ---- Phase 6: first-run migration. Offer to register the current
    #      interactive tenant as a profile so future runs can /tenant
    #      switch without re-walking the partner-center picker.
    if (Get-Command Test-FirstRunMigration -ErrorAction SilentlyContinue) { Test-FirstRunMigration }

    Write-AuditBanner

    $running = $true
    while ($running) {
        Initialize-UI
        Write-Banner

        $b = $Global:M365Box

        # ---- Tenant + connection status bar ----
        # Phase 6: per-tenant color when a registered profile is active.
        $tenantHeaderColor = if (Get-Command Get-TenantBannerColor -ErrorAction SilentlyContinue) { Get-TenantBannerColor -Name $script:SessionState.TenantName } else { $Global:M365Colors.Accent }
        Write-Host ("  " + $b.TL + [string]::new($b.H, 1) + " Tenant " + [string]::new($b.H, 49) + $b.TR) -ForegroundColor $tenantHeaderColor
        Write-Host ("  " + $b.V + "  ") -ForegroundColor $tenantHeaderColor -NoNewline
        if ($script:SessionState.TenantMode -eq "Profile") {
            Write-Host "Profile " -NoNewline -ForegroundColor $tenantHeaderColor
            Write-Host $script:SessionState.TenantName -NoNewline -ForegroundColor White
            if ($script:SessionState.TenantDomain) { Write-Host " ($($script:SessionState.TenantDomain))" -NoNewline -ForegroundColor $Global:M365Colors.Info }
        } elseif ($script:SessionState.TenantMode -eq "Partner") {
            Write-Host "GDAP " -NoNewline -ForegroundColor $Global:M365Colors.Highlight
            Write-Host $script:SessionState.TenantName -NoNewline -ForegroundColor White
            if ($script:SessionState.TenantDomain) { Write-Host " ($($script:SessionState.TenantDomain))" -NoNewline -ForegroundColor $Global:M365Colors.Info }
        } else {
            Write-Host "Direct (own organization)" -NoNewline -ForegroundColor White
        }
        $cursorPos = $Host.UI.RawUI.CursorPosition.X; $remaining = 62 - $cursorPos
        if ($remaining -gt 0) { Write-Host (" " * $remaining) -NoNewline }
        Write-Host ($b.V) -ForegroundColor $Global:M365Colors.Accent

        $gs = if ($script:SessionState.MgGraph) { "OK" } else { "---" }
        $es = if ($script:SessionState.ExchangeOnline) { "OK" } else { "---" }
        $ss = if ($script:SessionState.ComplianceCenter) { "OK" } else { "---" }

        Write-Host ("  " + $b.V + "  Graph: ") -ForegroundColor $Global:M365Colors.Accent -NoNewline
        Write-Host ("{0,-6}" -f $gs) -ForegroundColor $(if ($script:SessionState.MgGraph) { "Green" } else { "Gray" }) -NoNewline
        Write-Host " EXO: " -ForegroundColor $Global:M365Colors.Accent -NoNewline
        Write-Host ("{0,-6}" -f $es) -ForegroundColor $(if ($script:SessionState.ExchangeOnline) { "Green" } else { "Gray" }) -NoNewline
        Write-Host " SCC: " -ForegroundColor $Global:M365Colors.Accent -NoNewline
        Write-Host ("{0,-6}" -f $ss) -ForegroundColor $(if ($script:SessionState.ComplianceCenter) { "Green" } else { "Gray" }) -NoNewline
        $cursorPos = $Host.UI.RawUI.CursorPosition.X; $remaining = 62 - $cursorPos
        if ($remaining -gt 0) { Write-Host (" " * $remaining) -NoNewline }
        Write-Host ($b.V) -ForegroundColor $Global:M365Colors.Accent

        Write-Host ("  " + $b.BL + [string]::new($b.H, 58) + $b.BR) -ForegroundColor $Global:M365Colors.Accent

        # ---- Mode banner (Preview / Live) ----
        $modeText  = if (Get-PreviewMode) { '  [ PREVIEW MODE -- dry-run, no tenant changes ]  ' } else { '  [ LIVE MODE -- changes apply to the tenant ]  ' }
        $modeColor = if (Get-PreviewMode) { 'Yellow' } else { 'Red' }
        Write-Host ""
        Write-Host $modeText -ForegroundColor $modeColor
        Write-Host ""

        $sel = Show-Menu -Title "Main Menu - Select a Task" -Options @(
            "Onboard New User",
            "Offboard User",
            "Add / Remove License",
            "Mailbox Archiving",
            "Security Group Management",
            "Distribution List Management",
            "Shared Mailbox Management",
            "Calendar Access Management",
            "User Profile Management",
            "Group Membership Manager",
            "Reporting",
            "eDiscovery",
            "Bulk Onboard from CSV...",
            "Generate Onboarding Templates from Tenant...",
            "Bulk Offboard from CSV...",
            "Audit & Reporting...",
            "MFA & Authentication...",
            "Teams Management...",
            "SharePoint...",
            "OneDrive Access...",
            "Guest Users...",
            "License & Cost...",
            "Scheduled Health Checks...",
            "Tenants...",
            "Incident Response...",
            "AI Assistant (Mark)..."
        ) -BackLabel "Quit and Disconnect" -HiddenOptions @(98, 99)

        switch ($sel) {
            0  { Start-Onboard }
            1  { Start-Offboard }
            2  { Start-LicenseManagement }
            3  { Start-ArchiveManagement }
            4  { Start-SecurityGroupManagement }
            5  { Start-DistributionListManagement }
            6  { Start-SharedMailboxManagement }
            7  { Start-CalendarAccessManagement }
            8  { Start-UserProfileManagement }
            9  { Start-GroupManagerMenu }
            10 { Start-ReportingMenu }
            11 { Start-eDiscoveryMenu }
            12 { Start-BulkOnboard }
            13 {
                if (Get-Command Start-TemplateGeneratorMenu -ErrorAction SilentlyContinue) {
                    Start-TemplateGeneratorMenu
                } else {
                    Write-Warn "Template Generator module not loaded."
                }
            }
            14 { Start-BulkOffboard }
            15 { Start-AuditReportingMenu }
            16 { Start-MFAMenu }
            17 { Start-TeamsMenu }
            18 { Start-SharePointMenu }
            19 {
                if (Get-Command Start-OneDriveManagerMenu -ErrorAction SilentlyContinue) {
                    Start-OneDriveManagerMenu
                } else {
                    Write-Warn "OneDrive manager not loaded."
                }
            }
            20 { Start-GuestUsersMenu }
            21 { Start-LicenseOptimizerMenu }
            22 { Start-SchedulerMenu }
            25 {
                if (Get-Command Start-AIAssistant -ErrorAction SilentlyContinue) {
                    Start-AIAssistant
                } else {
                    Write-Warn "AI Assistant module not loaded (AIAssistant.ps1 missing)."
                }
            }
            98 {
                if (Get-Command Invoke-MgGraphFullRepair -ErrorAction SilentlyContinue) {
                    Invoke-MgGraphFullRepair
                } else {
                    Write-Warn "Repair helper not loaded."
                }
            }
            99 { Start-AIAssistant }
            23 {
                # Phase 6: Tenants submenu drives switch / register / edit /
                # remove / dashboard. Legacy "Switch Tenant" path falls back
                # to Select-TenantMode when the registry is empty.
                if ((Get-Command Start-TenantMenu -ErrorAction SilentlyContinue) -and (Get-Tenants).Count -gt 0) {
                    Start-TenantMenu
                } else {
                    Write-Host ""
                    if (Confirm-Action "Disconnect ALL sessions and switch tenant?") {
                        Reset-AllSessions
                        if (-not (Select-TenantMode)) {
                            Write-Warn "No tenant selected. Defaulting to direct mode."
                            $script:SessionState.TenantMode = "Direct"
                            $script:SessionState.TenantName = "Own Tenant"
                        }
                    }
                }
            }
            24 {
                if (Get-Command Start-IncidentResponseMenu -ErrorAction SilentlyContinue) {
                    Start-IncidentResponseMenu
                } else {
                    Write-Warn "Incident response module not loaded."
                }
            }
            -1 {
                Write-Host ""
                if (Confirm-Action "Quit and disconnect all sessions?") {
                    Disconnect-AllSessions
                    $running = $false
                }
            }
        }
    }

    Write-Host ""
    Write-Host "  Goodbye!" -ForegroundColor $Global:M365Colors.Title
    Write-Host ""
}

# ---- Run ----
Start-M365Admin
