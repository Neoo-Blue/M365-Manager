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

# ---- Load all modules ----
$loadErrors = @()
$modules = @(
    "UI.ps1","Auth.ps1","Audit.ps1","Preview.ps1","Templates.ps1",
    "Onboard.ps1","BulkOnboard.ps1","Offboard.ps1","BulkOffboard.ps1",
    "License.ps1","Archive.ps1","SecurityGroup.ps1","DistributionList.ps1",
    "SharedMailbox.ps1","CalendarAccess.ps1","UserProfile.ps1",
    "Reports.ps1","eDiscovery.ps1","GroupManager.ps1",
    "AuditViewer.ps1","Undo.ps1","AIAssistant.ps1"
)

foreach ($mod in $modules) {
    $modPath = Join-Path $ScriptRoot $mod
    if (Test-Path $modPath) {
        try {
            . "$modPath"
        } catch {
            $loadErrors += "  $mod : $_"
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
        Write-Host ""; Write-Host "  Goodbye!" -ForegroundColor $script:Colors.Title; return
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
        Write-Host "  Goodbye!" -ForegroundColor $script:Colors.Title
        return
    }

    Write-AuditBanner

    $running = $true
    while ($running) {
        Initialize-UI
        Write-Banner

        $b = $script:Box

        # ---- Tenant + connection status bar ----
        Write-Host ("  " + $b.TL + [string]::new($b.H, 1) + " Tenant " + [string]::new($b.H, 49) + $b.TR) -ForegroundColor $script:Colors.Accent
        Write-Host ("  " + $b.V + "  ") -ForegroundColor $script:Colors.Accent -NoNewline
        if ($script:SessionState.TenantMode -eq "Partner") {
            Write-Host "GDAP " -NoNewline -ForegroundColor $script:Colors.Highlight
            Write-Host $script:SessionState.TenantName -NoNewline -ForegroundColor White
            if ($script:SessionState.TenantDomain) { Write-Host " ($($script:SessionState.TenantDomain))" -NoNewline -ForegroundColor $script:Colors.Info }
        } else {
            Write-Host "Direct (own organization)" -NoNewline -ForegroundColor White
        }
        $cursorPos = $Host.UI.RawUI.CursorPosition.X; $remaining = 62 - $cursorPos
        if ($remaining -gt 0) { Write-Host (" " * $remaining) -NoNewline }
        Write-Host ($b.V) -ForegroundColor $script:Colors.Accent

        $gs = if ($script:SessionState.MgGraph) { "OK" } else { "---" }
        $es = if ($script:SessionState.ExchangeOnline) { "OK" } else { "---" }
        $ss = if ($script:SessionState.ComplianceCenter) { "OK" } else { "---" }

        Write-Host ("  " + $b.V + "  Graph: ") -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host ("{0,-6}" -f $gs) -ForegroundColor $(if ($script:SessionState.MgGraph) { "Green" } else { "Gray" }) -NoNewline
        Write-Host " EXO: " -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host ("{0,-6}" -f $es) -ForegroundColor $(if ($script:SessionState.ExchangeOnline) { "Green" } else { "Gray" }) -NoNewline
        Write-Host " SCC: " -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host ("{0,-6}" -f $ss) -ForegroundColor $(if ($script:SessionState.ComplianceCenter) { "Green" } else { "Gray" }) -NoNewline
        $cursorPos = $Host.UI.RawUI.CursorPosition.X; $remaining = 62 - $cursorPos
        if ($remaining -gt 0) { Write-Host (" " * $remaining) -NoNewline }
        Write-Host ($b.V) -ForegroundColor $script:Colors.Accent

        Write-Host ("  " + $b.BL + [string]::new($b.H, 58) + $b.BR) -ForegroundColor $script:Colors.Accent

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
            "Bulk Offboard from CSV...",
            "Audit & Reporting...",
            "Switch Tenant"
        ) -BackLabel "Quit and Disconnect" -HiddenOptions @(99)

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
            13 { Start-BulkOffboard }
            14 { Start-AuditReportingMenu }
            99 { Start-AIAssistant }
            15 {
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
    Write-Host "  Goodbye!" -ForegroundColor $script:Colors.Title
    Write-Host ""
}

# ---- Run ----
Start-M365Admin
