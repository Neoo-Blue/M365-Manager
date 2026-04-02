# ============================================================
#  Main.ps1 - M365 Administration Tool - Entry Point
# ============================================================
#  Usage:  powershell -ExecutionPolicy Bypass -File Main.ps1
# ============================================================

# ---- Determine script root ----
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

    # ---- Main loop ----
    $running = $true
    while ($running) {
        Initialize-UI
        Write-Banner

        $b = $script:Box

        # Show connection status bar
        Write-Host ("  " + $b.TL + [string]::new($b.H, 1) + " Session Status " + [string]::new($b.H, 40) + $b.TR) -ForegroundColor $script:Colors.Accent

        $aadStatus  = if ($script:SessionState.AzureAD)        { "Connected" } else { "---" }
        $exoStatus  = if ($script:SessionState.ExchangeOnline)  { "Connected" } else { "---" }
        $msolStatus = if ($script:SessionState.MSOnline)        { "Connected" } else { "---" }

        Write-Host ("  " + $b.V + "  AzureAD: ") -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host ("{0,-12}" -f $aadStatus) -ForegroundColor $(if ($script:SessionState.AzureAD) { "Green" } else { "Gray" }) -NoNewline
        Write-Host " EXO: " -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host ("{0,-12}" -f $exoStatus) -ForegroundColor $(if ($script:SessionState.ExchangeOnline) { "Green" } else { "Gray" }) -NoNewline
        Write-Host " MSOL: " -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host ("{0,-8}" -f $msolStatus) -ForegroundColor $(if ($script:SessionState.MSOnline) { "Green" } else { "Gray" }) -NoNewline
        Write-Host (" " + $b.V) -ForegroundColor $script:Colors.Accent

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
            "User Profile Management"
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
