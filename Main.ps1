# ============================================================
#  Main.ps1 - M365 Administration Tool - Entry Point
# ============================================================
#  Usage:  powershell -ExecutionPolicy Bypass -File Main.ps1
# ============================================================

$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

# ---- Dot-source all modules ----
. "$ScriptRoot\UI.ps1"
. "$ScriptRoot\Auth.ps1"
. "$ScriptRoot\Onboard.ps1"
. "$ScriptRoot\Offboard.ps1"
. "$ScriptRoot\License.ps1"
. "$ScriptRoot\Archive.ps1"
. "$ScriptRoot\SecurityGroup.ps1"
. "$ScriptRoot\DistributionList.ps1"
. "$ScriptRoot\SharedMailbox.ps1"
. "$ScriptRoot\CalendarAccess.ps1"
. "$ScriptRoot\UserProfile.ps1"
. "$ScriptRoot\Reports.ps1"

# ---- Bootstrap ----
function Start-M365Admin {
    Initialize-UI
    Write-Banner

    # Pre-flight: check modules
    if (-not (Assert-ModulesInstalled)) {
        Write-ErrorMsg "Required modules are missing. Exiting."
        Pause-ForUser
        return
    }

    Write-Success "All required PowerShell modules detected."
    Write-Host ""

    # ---- Tenant selection (own org vs GDAP customer) ----
    if (-not (Select-TenantMode)) {
        Write-Host ""
        Write-Host "  Goodbye!" -ForegroundColor $script:Colors.Title
        return
    }

    # ---- Main loop ----
    $running = $true
    while ($running) {
        Initialize-UI
        Write-Banner

        $b = $script:Box

        # ---- Tenant context bar ----
        $tenantDisplay = Get-TenantDisplayString
        Write-Host ("  " + $b.TL + [string]::new($b.H, 1) + " Tenant " + [string]::new($b.H, 49) + $b.TR) -ForegroundColor $script:Colors.Accent
        Write-Host ("  " + $b.V + "  ") -ForegroundColor $script:Colors.Accent -NoNewline

        if ($script:SessionState.TenantMode -eq "Partner") {
            Write-Host "GDAP " -NoNewline -ForegroundColor $script:Colors.Highlight
            Write-Host $script:SessionState.TenantName -NoNewline -ForegroundColor White
            if ($script:SessionState.TenantDomain) {
                Write-Host " ($($script:SessionState.TenantDomain))" -NoNewline -ForegroundColor $script:Colors.Info
            }
        } else {
            Write-Host "Direct (own organization)" -NoNewline -ForegroundColor White
        }

        # Pad to fill the box
        $cursorPos = $Host.UI.RawUI.CursorPosition.X
        $remaining = 62 - $cursorPos
        if ($remaining -gt 0) { Write-Host (" " * $remaining) -NoNewline }
        Write-Host ($b.V) -ForegroundColor $script:Colors.Accent

        # ---- Connection status bar ----
        $graphStatus = if ($script:SessionState.MgGraph)         { "Connected" } else { "---" }
        $exoStatus   = if ($script:SessionState.ExchangeOnline)  { "Connected" } else { "---" }

        Write-Host ("  " + $b.V + "  Graph: ") -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host ("{0,-13}" -f $graphStatus) -ForegroundColor $(if ($script:SessionState.MgGraph) { "Green" } else { "Gray" }) -NoNewline
        Write-Host " EXO: " -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host ("{0,-13}" -f $exoStatus) -ForegroundColor $(if ($script:SessionState.ExchangeOnline) { "Green" } else { "Gray" }) -NoNewline

        # Pad to fill
        $cursorPos = $Host.UI.RawUI.CursorPosition.X
        $remaining = 62 - $cursorPos
        if ($remaining -gt 0) { Write-Host (" " * $remaining) -NoNewline }
        Write-Host ($b.V) -ForegroundColor $script:Colors.Accent

        Write-Host ("  " + $b.BL + [string]::new($b.H, 58) + $b.BR) -ForegroundColor $script:Colors.Accent

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
            "Reporting",
            "Switch Tenant"
        ) -BackLabel "Quit and Disconnect"

        switch ($sel) {
            0 { Start-Onboard }
            1 { Start-Offboard }
            2 { Start-LicenseManagement }
            3 { Start-ArchiveManagement }
            4 { Start-SecurityGroupManagement }
            5 { Start-DistributionListManagement }
            6 { Start-SharedMailboxManagement }
            7 { Start-CalendarAccessManagement }
            8 { Start-UserProfileManagement }
            9 { Start-ReportingMenu }
            10 {
                # ---- Switch Tenant ----
                Write-Host ""
                if (Confirm-Action "Disconnect current sessions and switch tenant?") {
                    Disconnect-AllSessions
                    if (-not (Select-TenantMode)) {
                        Write-Warn "Tenant selection cancelled. Returning to menu."
                        # Restore to direct mode so the tool still works
                        $script:SessionState.TenantMode = "Direct"
                        $script:SessionState.TenantId   = $null
                        $script:SessionState.TenantName  = "Own Tenant"
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
