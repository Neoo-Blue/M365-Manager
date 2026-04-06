# ============================================================
#  Main.ps1 - M365 Administration Tool - Entry Point
# ============================================================

$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }

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
. "$ScriptRoot\eDiscovery.ps1"
. "$ScriptRoot\GroupManager.ps1"

function Start-M365Admin {
    Initialize-UI
    Write-Banner

    if (-not (Assert-ModulesInstalled)) {
        Write-ErrorMsg "Dependency check failed. Exiting."
        Pause-ForUser
        return
    }
    Write-Host ""

    if (-not (Select-TenantMode)) {
        Write-Host ""
        Write-Host "  Goodbye!" -ForegroundColor $script:Colors.Title
        return
    }

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
            "Switch Tenant"
        ) -BackLabel "Quit and Disconnect"

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
            12 {
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

Start-M365Admin
