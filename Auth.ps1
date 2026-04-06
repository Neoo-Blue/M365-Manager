# ============================================================
#  Auth.ps1 - Authentication & Session Management
# ============================================================
# Supports:
#  - Direct tenant management (your own org)
#  - GDAP partner access (manage customer tenants)
# ============================================================

$script:SessionState = @{
    MgGraph         = $false
    ExchangeOnline  = $false
    TenantMode      = "Direct"     # "Direct" or "Partner"
    TenantId        = $null        # Customer tenant ID when in Partner mode
    TenantName      = $null        # Customer display name
    TenantDomain    = $null        # Customer default domain
    PartnerConnected = $false      # Whether we already auth'd to partner tenant
}

# Scopes for direct or delegated operations
$script:MgScopes = @(
    "User.ReadWrite.All",
    "Group.ReadWrite.All",
    "Directory.ReadWrite.All",
    "Organization.Read.All",
    "UserAuthenticationMethod.ReadWrite.All"
)

# Extra scopes needed to list customer tenants from partner
$script:MgPartnerScopes = @(
    "Directory.Read.All",
    "Contract.Read.All"
)

# ---- Module pre-check ----
function Assert-ModulesInstalled {
    $required = @(
        @{ Name = "Microsoft.Graph.Authentication";            Install = "Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force" },
        @{ Name = "Microsoft.Graph.Users";                     Install = "Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force" },
        @{ Name = "Microsoft.Graph.Users.Actions";             Install = "Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser -Force" },
        @{ Name = "Microsoft.Graph.Groups";                    Install = "Install-Module Microsoft.Graph.Groups -Scope CurrentUser -Force" },
        @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; Install = "Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force" },
        @{ Name = "ExchangeOnlineManagement";                  Install = "Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force" }
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
                try { Invoke-Expression $m.Install; Write-Success "$($m.Name) installed." }
                catch { Write-ErrorMsg "Failed to install $($m.Name): $_" }
            }
        } else {
            Write-ErrorMsg "Cannot continue without required modules."
            return $false
        }
    }

    $exoMod = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if ($exoMod) { Write-InfoMsg "ExchangeOnlineManagement: v$($exoMod.Version)" }
    $mgAuth = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
    if ($mgAuth) { Write-InfoMsg "Microsoft.Graph.Authentication: v$($mgAuth.Version)" }

    return $true
}

# ============================================================
#  Tenant Mode Selection
# ============================================================
function Select-TenantMode {
    <#
    .SYNOPSIS
        Called once at startup. Asks if managing own tenant or a customer tenant.
        If partner mode, connects to partner tenant, lists customers via GDAP,
        and stores the selected customer tenant ID for all subsequent connections.
    #>

    Write-SectionHeader "Tenant Selection"

    $mode = Show-Menu -Title "Which tenant are you managing?" -Options @(
        "My own organization (direct admin)",
        "A customer tenant (GDAP partner access)"
    ) -BackLabel "Quit"

    if ($mode -eq -1) { return $false }

    if ($mode -eq 0) {
        # ---- Direct mode ----
        $script:SessionState.TenantMode = "Direct"
        $script:SessionState.TenantId   = $null
        $script:SessionState.TenantName = "Own Tenant"
        Write-Success "Direct tenant mode selected."
        return $true
    }

    # ---- Partner / GDAP mode ----
    $script:SessionState.TenantMode = "Partner"

    Write-InfoMsg "First, signing in to your PARTNER tenant to list customer tenants..."
    Write-InfoMsg "A browser window will open for sign-in."
    Write-Host ""

    # Connect to partner tenant (no TenantId = own tenant)
    try {
        Connect-MgGraph -Scopes ($script:MgPartnerScopes + $script:MgScopes) -NoWelcome -ErrorAction Stop
        $script:SessionState.PartnerConnected = $true
        $ctx = Get-MgContext
        Write-Success "Signed in as $($ctx.Account) (Partner Tenant: $($ctx.TenantId))"
    } catch {
        Write-ErrorMsg "Failed to connect to partner tenant: $_"
        return $false
    }

    # ---- List customer tenants via contracts ----
    Write-InfoMsg "Fetching customer tenant list..."
    Write-Host ""

    $customers = @()
    try {
        # Get-MgContract lists all delegated admin relationships
        $contracts = @(Get-MgContract -All -ErrorAction Stop)

        if ($contracts.Count -eq 0) {
            Write-Warn "No customer contracts found."
            Write-InfoMsg "This could mean:"
            Write-InfoMsg "  - No GDAP relationships are established"
            Write-InfoMsg "  - Your account lacks the required admin role"
            Write-Host ""

            # Offer manual entry as fallback
            $manual = Read-UserInput "Enter a customer tenant ID or domain manually (or 'quit')"
            if ($manual -eq 'quit' -or [string]::IsNullOrWhiteSpace($manual)) { return $false }

            $script:SessionState.TenantId     = $manual
            $script:SessionState.TenantName   = $manual
            $script:SessionState.TenantDomain  = $manual

            # Disconnect from partner, will reconnect to customer
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            $script:SessionState.MgGraph = $false
            $script:SessionState.PartnerConnected = $false

            Write-Success "Will connect to tenant: $manual"
            return $true
        }

        # Build customer list from contracts
        foreach ($c in $contracts) {
            $customers += [PSCustomObject]@{
                DisplayName   = $c.DisplayName
                CustomerId    = $c.CustomerId
                DefaultDomain = $c.DefaultDomainName
            }
        }

        # Deduplicate by CustomerId (a customer may have multiple contracts)
        $customers = $customers | Sort-Object DisplayName -Unique

    } catch {
        Write-ErrorMsg "Failed to list customers: $_"
        Write-InfoMsg "Falling back to manual tenant entry."
        Write-Host ""

        $manual = Read-UserInput "Enter a customer tenant ID or domain (or 'quit')"
        if ($manual -eq 'quit' -or [string]::IsNullOrWhiteSpace($manual)) { return $false }

        $script:SessionState.TenantId     = $manual
        $script:SessionState.TenantName   = $manual
        $script:SessionState.TenantDomain  = $manual

        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $script:SessionState.MgGraph = $false
        $script:SessionState.PartnerConnected = $false

        Write-Success "Will connect to tenant: $manual"
        return $true
    }

    # ---- Show customer picker ----
    Write-SectionHeader "Customer Tenants ($($customers.Count) found)"

    # If too many to show in a menu, offer search
    if ($customers.Count -gt 20) {
        $searchInput = Read-UserInput "Search customer by name (or 'all' to list all)"
        if ($searchInput -ne 'all') {
            $customers = @($customers | Where-Object { $_.DisplayName -like "*$searchInput*" })
            if ($customers.Count -eq 0) {
                Write-ErrorMsg "No customers matching '$searchInput'."
                return $false
            }
        }
    }

    $custLabels = $customers | ForEach-Object {
        "$($_.DisplayName)  ($($_.DefaultDomain))"
    }

    $sel = Show-Menu -Title "Select Customer Tenant" -Options $custLabels -BackLabel "Cancel"
    if ($sel -eq -1) { return $false }

    $selected = $customers[$sel]
    $script:SessionState.TenantId     = $selected.CustomerId
    $script:SessionState.TenantName   = $selected.DisplayName
    $script:SessionState.TenantDomain  = $selected.DefaultDomain

    Write-Host ""
    Write-Success "Selected: $($selected.DisplayName)"
    Write-StatusLine "Tenant ID" $selected.CustomerId "White"
    Write-StatusLine "Domain" $selected.DefaultDomain "White"
    Write-Host ""

    # Disconnect partner graph session (will reconnect to customer)
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    $script:SessionState.MgGraph = $false
    $script:SessionState.PartnerConnected = $false

    return $true
}

# ============================================================
#  Service Connections (tenant-aware)
# ============================================================

function Connect-Graph {
    if ($script:SessionState.MgGraph) {
        Write-InfoMsg "Microsoft Graph already connected."
        return $true
    }

    $targetLabel = if ($script:SessionState.TenantMode -eq "Partner") {
        "$($script:SessionState.TenantName) (GDAP)"
    } else { "own tenant" }

    Write-InfoMsg "Connecting to Microsoft Graph ($targetLabel)..."

    try {
        $params = @{
            Scopes    = $script:MgScopes
            NoWelcome = $true
        }

        if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantId) {
            $params["TenantId"] = $script:SessionState.TenantId
        }

        Connect-MgGraph @params -ErrorAction Stop
        $script:SessionState.MgGraph = $true

        $ctx = Get-MgContext
        Write-Success "Microsoft Graph connected as $($ctx.Account)"
        if ($script:SessionState.TenantMode -eq "Partner") {
            Write-Success "  Target tenant: $($script:SessionState.TenantName) ($($ctx.TenantId))"
        }
        return $true
    } catch {
        Write-ErrorMsg "Microsoft Graph connection failed: $_"
        if ($script:SessionState.TenantMode -eq "Partner") {
            Write-Warn "Ensure your GDAP relationship grants the required admin roles."
            Write-InfoMsg "Required roles: User Administrator, Groups Administrator,"
            Write-InfoMsg "  License Administrator, Directory Readers."
        }
        return $false
    }
}

function Connect-EXO {
    if ($script:SessionState.ExchangeOnline) {
        Write-InfoMsg "Exchange Online already connected."
        return $true
    }

    $targetLabel = if ($script:SessionState.TenantMode -eq "Partner") {
        "$($script:SessionState.TenantName) (GDAP)"
    } else { "own tenant" }

    Write-InfoMsg "Connecting to Exchange Online ($targetLabel)..."

    # Build EXO connection params
    $exoParams = @{ ShowBanner = $false }
    if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantDomain) {
        $exoParams["DelegatedOrganization"] = $script:SessionState.TenantDomain
    }

    # Attempt 1: Standard browser login
    try {
        Connect-ExchangeOnline @exoParams -ErrorAction Stop
        $script:SessionState.ExchangeOnline = $true
        Write-Success "Exchange Online connected."
        return $true
    } catch {
        $errMsg = $_.Exception.Message
        Write-Warn "Browser login failed: $errMsg"
    }

    # ---- Detect MSAL broker DLL conflict ----
    if ($errMsg -match "BrokerExtension|WithBroker|Method not found") {
        Write-Host ""
        Write-Warn "MSAL broker DLL conflict detected."
        Write-InfoMsg "Microsoft Graph and ExchangeOnlineManagement load conflicting MSAL broker versions."
        Write-InfoMsg "Fix: disable the broker DLL so EXO falls back to standard browser auth."
        Write-Host ""

        if (Confirm-Action "Auto-fix: disable the conflicting MSAL broker DLL and retry?") {

            # Find the broker DLL(s) inside the EXO module directory
            $exoMod = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
                Sort-Object Version -Descending | Select-Object -First 1

            if ($exoMod) {
                $brokerDlls = @(Get-ChildItem -Path $exoMod.ModuleBase -Recurse `
                    -Filter "Microsoft.Identity.Client.Broker.dll" -ErrorAction SilentlyContinue)

                if ($brokerDlls.Count -gt 0) {
                    foreach ($dll in $brokerDlls) {
                        $bakPath = "$($dll.FullName).bak"
                        try {
                            # Rename to .bak (disables it; the DLL isn't loaded yet in THIS session
                            # because the first Connect attempt loaded the assembly from the Graph
                            # module path, not from the EXO path -- the conflict IS the mismatch)
                            if (Test-Path $dll.FullName) {
                                Rename-Item -Path $dll.FullName -NewName "$($dll.Name).bak" -Force -ErrorAction Stop
                                Write-Success "Disabled: $($dll.FullName)"
                            }
                        } catch {
                            Write-Warn "Could not rename $($dll.FullName): $_"
                            Write-InfoMsg "Try running the tool as Administrator if permission denied."
                        }
                    }
                } else {
                    Write-Warn "Broker DLL not found in EXO module directory."
                }
            }

            # Retry connection -- without the broker DLL, EXO falls back to browser-only auth
            Write-Host ""
            Write-InfoMsg "Retrying Exchange Online connection without broker..."
            try {
                Connect-ExchangeOnline @exoParams -ErrorAction Stop
                $script:SessionState.ExchangeOnline = $true
                Write-Success "Exchange Online connected (broker disabled)."
                return $true
            } catch {
                $retryErr = $_.Exception.Message
                Write-ErrorMsg "Still failed after disabling broker: $retryErr"
                Write-Host ""

                # The broker assembly may already be loaded in memory from the first attempt.
                # Need a full restart for the rename to take effect.
                Write-Warn "The broken DLL was already loaded in memory from the first attempt."
                Write-InfoMsg "The broker has been disabled on disk. Restarting the tool to pick up the fix..."
                Start-Sleep -Seconds 2

                # Relaunch
                $mainScript = Join-Path $PSScriptRoot "Main.ps1"
                if (-not $PSScriptRoot) {
                    $mainScript = Join-Path (Split-Path -Parent $MyInvocation.ScriptName) "Main.ps1"
                }

                Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$mainScript`""
                Write-Success "New instance launched. This window will close."
                Start-Sleep -Seconds 2
                exit
            }
        }

        Write-Host ""
        Write-Warn "Manual fix:"
        Write-InfoMsg "1. Find the ExchangeOnlineManagement module folder:"
        Write-InfoMsg "   (Get-Module -ListAvailable ExchangeOnlineManagement).ModuleBase"
        Write-InfoMsg "2. Find and rename Microsoft.Identity.Client.Broker.dll to .bak"
        Write-InfoMsg "3. Restart PowerShell and run the tool again."
        Write-Host ""
        return $false
    }

    # ---- Generic failure (not MSAL broker) ----
    Write-Host ""
    Write-ErrorMsg "Exchange Online connection failed."
    Write-InfoMsg "Check your credentials, network, and module version."
    return $false
}

# ============================================================
#  Per-task connection sets
# ============================================================

function Connect-ForTask {
    param(
        [ValidateSet(
            "Onboard","Offboard","License","Archive",
            "SecurityGroup","DistributionList","SharedMailbox","CalendarAccess",
            "UserProfile","Report"
        )]
        [string]$Task
    )

    # EXO listed first so it loads its MSAL before Graph loads a conflicting version
    $map = @{
        Onboard          = @("EXO","Graph")
        Offboard         = @("EXO","Graph")
        License          = @("Graph")
        Archive          = @("EXO","Graph")
        SecurityGroup    = @("Graph")
        DistributionList = @("EXO","Graph")
        SharedMailbox    = @("EXO","Graph")
        CalendarAccess   = @("EXO","Graph")
        UserProfile      = @("Graph")
        Report           = @("EXO","Graph")
    }

    $services = $map[$Task]
    $needed = @()
    foreach ($svc in $services) {
        switch ($svc) {
            "Graph" { if (-not $script:SessionState.MgGraph)        { $needed += $svc } }
            "EXO"   { if (-not $script:SessionState.ExchangeOnline) { $needed += $svc } }
        }
    }

    if ($needed.Count -eq 0) {
        Write-InfoMsg "All required services already connected."
        return $true
    }

    Write-InfoMsg "This task requires: $($services -join ', ')"
    Write-InfoMsg "Need to connect: $($needed -join ', ')"
    Write-Host ""

    foreach ($svc in $services) {
        switch ($svc) {
            "Graph" { if (-not (Connect-Graph)) { return $false } }
            "EXO"   { if (-not (Connect-EXO))   { return $false } }
        }
    }
    return $true
}

# ============================================================
#  Disconnect all
# ============================================================
function Disconnect-AllSessions {
    Write-SectionHeader "Disconnecting Sessions"

    if ($script:SessionState.MgGraph) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue; Write-Success "Microsoft Graph disconnected." }
        catch { Write-Warn "Graph disconnect issue: $_" }
    }
    if ($script:SessionState.ExchangeOnline) {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue; Write-Success "Exchange Online disconnected." }
        catch { Write-Warn "EXO disconnect issue: $_" }
    }

    $script:SessionState.MgGraph          = $false
    $script:SessionState.ExchangeOnline    = $false
    $script:SessionState.PartnerConnected  = $false

    Write-Success "All sessions cleared."
}

# ============================================================
#  Helper: Get tenant display string for status bar
# ============================================================
function Get-TenantDisplayString {
    if ($script:SessionState.TenantMode -eq "Partner") {
        return "GDAP: $($script:SessionState.TenantName)"
    }
    return "Direct (own tenant)"
}
