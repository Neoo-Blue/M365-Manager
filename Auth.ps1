# ============================================================
#  Auth.ps1 - Authentication & Session Management
# ============================================================
# Uses browser-based interactive (OAuth) login for all services.
# Each service is connected at most once per session.
# On exit, Disconnect-AllSessions tears everything down.
# ============================================================

$script:SessionState = @{
    AzureAD         = $false
    ExchangeOnline  = $false
    MSOnline        = $false
}

# ---- Module pre-check ----
function Assert-ModulesInstalled {
    $required = @(
        @{ Name = "AzureAD";                 Install = "Install-Module AzureAD -Scope CurrentUser -Force" },
        @{ Name = "ExchangeOnlineManagement"; Install = "Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force" },
        @{ Name = "MSOnline";                 Install = "Install-Module MSOnline -Scope CurrentUser -Force" }
    )

    $missing = @()
    foreach ($mod in $required) {
        if (-not (Get-Module -ListAvailable -Name $mod.Name)) {
            $missing += $mod
        }
    }

    if ($missing.Count -gt 0) {
        Write-SectionHeader "Missing PowerShell Modules"
        foreach ($m in $missing) {
            Write-Warn "$($m.Name) is not installed."
            Write-InfoMsg "  Run: $($m.Install)"
        }
        Write-Host ""
        if (Confirm-Action "Attempt to install missing modules now?") {
            foreach ($m in $missing) {
                Write-InfoMsg "Installing $($m.Name)..."
                try {
                    Invoke-Expression $m.Install
                    Write-Success "$($m.Name) installed."
                } catch {
                    Write-ErrorMsg "Failed to install $($m.Name): $_"
                }
            }
        } else {
            Write-ErrorMsg "Cannot continue without required modules."
            return $false
        }
    }
    return $true
}

# ---- Service connections (browser-based interactive login) ----

function Connect-AAD {
    if ($script:SessionState.AzureAD) {
        Write-InfoMsg "Azure AD already connected."
        return $true
    }

    Write-InfoMsg "Connecting to Azure AD (browser login)..."
    try {
        Connect-AzureAD -ErrorAction Stop | Out-Null
        $script:SessionState.AzureAD = $true
        Write-Success "Azure AD connected."
        return $true
    } catch {
        Write-ErrorMsg "Azure AD connection failed: $_"
        return $false
    }
}

function Connect-EXO {
    if ($script:SessionState.ExchangeOnline) {
        Write-InfoMsg "Exchange Online already connected."
        return $true
    }

    Write-InfoMsg "Connecting to Exchange Online (browser login)..."
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $script:SessionState.ExchangeOnline = $true
        Write-Success "Exchange Online connected."
        return $true
    } catch {
        Write-ErrorMsg "Exchange Online connection failed: $_"
        return $false
    }
}

function Connect-MSOL {
    if ($script:SessionState.MSOnline) {
        Write-InfoMsg "MSOnline already connected."
        return $true
    }

    Write-InfoMsg "Connecting to MSOnline (browser login)..."
    try {
        Connect-MsolService -ErrorAction Stop
        $script:SessionState.MSOnline = $true
        Write-Success "MSOnline connected."
        return $true
    } catch {
        Write-ErrorMsg "MSOnline connection failed: $_"
        return $false
    }
}

# ---- Per-task connection sets ----

function Connect-ForTask {
    <#
    .SYNOPSIS
        Connects to the services required for a given task.
        Only prompts the browser login for services not yet connected.
        Returns $true if all required services are connected.
    #>
    param(
        [ValidateSet(
            "Onboard","Offboard","License","Archive",
            "SecurityGroup","DistributionList","SharedMailbox","CalendarAccess",
            "UserProfile"
        )]
        [string]$Task
    )

    $map = @{
        Onboard          = @("AAD","MSOL","EXO")
        Offboard         = @("AAD","MSOL","EXO")
        License          = @("AAD","MSOL")
        Archive          = @("AAD","EXO")
        SecurityGroup    = @("AAD")
        DistributionList = @("AAD","EXO")
        SharedMailbox    = @("AAD","EXO")
        CalendarAccess   = @("AAD","EXO")
        UserProfile      = @("AAD")
    }

    $services = $map[$Task]

    # Figure out which ones still need login
    $needed = @()
    foreach ($svc in $services) {
        switch ($svc) {
            "AAD"  { if (-not $script:SessionState.AzureAD)        { $needed += $svc } }
            "EXO"  { if (-not $script:SessionState.ExchangeOnline)  { $needed += $svc } }
            "MSOL" { if (-not $script:SessionState.MSOnline)        { $needed += $svc } }
        }
    }

    if ($needed.Count -eq 0) {
        Write-InfoMsg "All required services already connected."
        return $true
    }

    Write-InfoMsg "This task requires: $($services -join ', ')"
    Write-InfoMsg "Need to connect: $($needed -join ', ')"
    Write-InfoMsg "A browser window will open for sign-in (once per service)."
    Write-Host ""

    foreach ($svc in $services) {
        switch ($svc) {
            "AAD"  { if (-not (Connect-AAD))  { return $false } }
            "EXO"  { if (-not (Connect-EXO))  { return $false } }
            "MSOL" { if (-not (Connect-MSOL)) { return $false } }
        }
    }
    return $true
}

# ---- Disconnect all ----
function Disconnect-AllSessions {
    Write-SectionHeader "Disconnecting Sessions"

    if ($script:SessionState.AzureAD) {
        try {
            Disconnect-AzureAD -ErrorAction SilentlyContinue
            Write-Success "Azure AD disconnected."
        } catch { Write-Warn "Azure AD disconnect issue: $_" }
    }
    if ($script:SessionState.ExchangeOnline) {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Write-Success "Exchange Online disconnected."
        } catch { Write-Warn "EXO disconnect issue: $_" }
    }
    if ($script:SessionState.MSOnline) {
        Write-InfoMsg "MSOnline session will expire automatically."
    }

    $script:SessionState.AzureAD        = $false
    $script:SessionState.ExchangeOnline  = $false
    $script:SessionState.MSOnline        = $false

    Write-Success "All sessions cleared."
}
