# ============================================================
#  Auth.ps1 - Authentication, Dependencies & Session Management
# ============================================================

$script:SessionState = @{
    MgGraph          = $false
    ExchangeOnline   = $false
    ComplianceCenter = $false
    TenantMode       = "Direct"
    TenantId         = $null
    TenantName       = $null
    TenantDomain     = $null
    PartnerConnected = $false
}

$script:MgScopes = @(
    "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All",
    "Organization.Read.All","UserAuthenticationMethod.ReadWrite.All"
)
$script:MgPartnerScopes = @("Directory.Read.All","Contract.Read.All")

# ============================================================
#  Dependency Management - install, import, verify
# ============================================================

function Assert-ModulesInstalled {
    Write-SectionHeader "Checking Dependencies"

    # ---- Define required modules with test commands ----
    $requiredModules = @(
        @{ Name = "Microsoft.Graph.Authentication";              TestCmd = "Get-MgContext" },
        @{ Name = "Microsoft.Graph.Users";                       TestCmd = "Get-MgUser" },
        @{ Name = "Microsoft.Graph.Users.Actions";               TestCmd = $null },
        @{ Name = "Microsoft.Graph.Groups";                      TestCmd = "Get-MgGroup" },
        @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; TestCmd = "Get-MgSubscribedSku" },
        @{ Name = "ExchangeOnlineManagement";                    TestCmd = "Connect-ExchangeOnline" }
    )

    $allGood = $true

    foreach ($mod in $requiredModules) {
        $modName = $mod.Name

        # ---- Step 1: Check if installed ----
        $installed = Get-Module -ListAvailable -Name $modName | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $installed) {
            Write-Warn "$modName is not installed. Installing..."
            try {
                Install-Module $modName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Success "$modName installed."
                $installed = Get-Module -ListAvailable -Name $modName | Sort-Object Version -Descending | Select-Object -First 1
            } catch {
                Write-ErrorMsg "Failed to install $modName : $_"
                $allGood = $false
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
                    Write-ErrorMsg "Failed to import $modName : $_"
                    $allGood = $false
                    continue
                }
            }
        } else {
            Write-InfoMsg "$modName v$($loaded.Version) already loaded."
        }

        # ---- Step 3: Verify test command exists ----
        if ($mod.TestCmd) {
            $cmdExists = Get-Command $mod.TestCmd -ErrorAction SilentlyContinue
            if (-not $cmdExists) {
                Write-ErrorMsg "$modName loaded but command '$($mod.TestCmd)' not found."
                Write-InfoMsg "Try: Remove-Module $modName; Import-Module $modName"
                $allGood = $false
            }
        }
    }

    if (-not $allGood) {
        Write-Host ""
        Write-ErrorMsg "Some modules have issues. The tool may not work correctly."
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

    $mode = Show-Menu -Title "Which tenant are you managing?" -Options @(
        "My own organization (direct admin)",
        "A customer tenant (GDAP partner access)"
    ) -BackLabel "Quit"

    if ($mode -eq -1) { return $false }

    if ($mode -eq 0) {
        $script:SessionState.TenantMode = "Direct"
        $script:SessionState.TenantId   = $null
        $script:SessionState.TenantName = "Own Tenant"
        Write-Success "Direct tenant mode selected."
        return $true
    }

    # ---- Partner / GDAP ----
    $script:SessionState.TenantMode = "Partner"

    Write-InfoMsg "Signing in to your PARTNER tenant to list customers..."
    try {
        Connect-MgGraph -Scopes ($script:MgPartnerScopes + $script:MgScopes) -NoWelcome -ErrorAction Stop
        $script:SessionState.PartnerConnected = $true
        $ctx = Get-MgContext
        Write-Success "Signed in as $($ctx.Account)"
    } catch {
        Write-ErrorMsg "Partner tenant login failed: $_"
        return $false
    }

    Write-InfoMsg "Fetching customer tenant list..."
    $customers = @()
    try {
        $contracts = @(Get-MgContract -All -ErrorAction Stop)
        if ($contracts.Count -eq 0) {
            Write-Warn "No customer contracts found."
            $manual = Read-UserInput "Enter tenant ID or domain manually (or 'back')"
            if ($manual -eq 'back' -or [string]::IsNullOrWhiteSpace($manual)) { Disconnect-MgGraph -ErrorAction SilentlyContinue; $script:SessionState.PartnerConnected = $false; return $false }
            $script:SessionState.TenantId = $manual; $script:SessionState.TenantName = $manual; $script:SessionState.TenantDomain = $manual
            Disconnect-MgGraph -ErrorAction SilentlyContinue; $script:SessionState.MgGraph = $false; $script:SessionState.PartnerConnected = $false
            return $true
        }

        foreach ($c in $contracts) {
            $customers += [PSCustomObject]@{ DisplayName = $c.DisplayName; CustomerId = $c.CustomerId; DefaultDomain = $c.DefaultDomainName }
        }
        $customers = $customers | Sort-Object DisplayName -Unique
    } catch {
        Write-ErrorMsg "Failed to list customers: $_"
        $manual = Read-UserInput "Enter tenant ID or domain (or 'back')"
        if ($manual -eq 'back' -or [string]::IsNullOrWhiteSpace($manual)) { Disconnect-MgGraph -ErrorAction SilentlyContinue; $script:SessionState.PartnerConnected = $false; return $false }
        $script:SessionState.TenantId = $manual; $script:SessionState.TenantName = $manual; $script:SessionState.TenantDomain = $manual
        Disconnect-MgGraph -ErrorAction SilentlyContinue; $script:SessionState.MgGraph = $false; $script:SessionState.PartnerConnected = $false
        return $true
    }

    if ($customers.Count -gt 20) {
        $searchInput = Read-UserInput "Search customer by name (or 'all')"
        if ($searchInput -ne 'all') {
            $customers = @($customers | Where-Object { $_.DisplayName -like "*$searchInput*" })
            if ($customers.Count -eq 0) { Write-ErrorMsg "None found."; return $false }
        }
    }

    $custLabels = $customers | ForEach-Object { "$($_.DisplayName)  ($($_.DefaultDomain))" }
    $sel = Show-Menu -Title "Select Customer Tenant" -Options $custLabels -BackLabel "Cancel"
    if ($sel -eq -1) { Disconnect-MgGraph -ErrorAction SilentlyContinue; $script:SessionState.PartnerConnected = $false; return $false }

    $selected = $customers[$sel]
    $script:SessionState.TenantId = $selected.CustomerId
    $script:SessionState.TenantName = $selected.DisplayName
    $script:SessionState.TenantDomain = $selected.DefaultDomain

    Write-Success "Selected: $($selected.DisplayName) ($($selected.DefaultDomain))"
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    $script:SessionState.MgGraph = $false; $script:SessionState.PartnerConnected = $false
    return $true
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

    $script:SessionState.MgGraph          = $false
    $script:SessionState.ExchangeOnline    = $false
    $script:SessionState.ComplianceCenter  = $false
    $script:SessionState.PartnerConnected  = $false
    $script:SessionState.TenantId          = $null
    $script:SessionState.TenantName        = $null
    $script:SessionState.TenantDomain      = $null
    $script:SessionState.TenantMode        = "Direct"

    Write-Success "All sessions and tenant context cleared."
}

# ============================================================
#  Service Connections
# ============================================================

function Connect-Graph {
    if ($script:SessionState.MgGraph) { Write-InfoMsg "Microsoft Graph already connected."; return $true }

    $targetLabel = if ($script:SessionState.TenantMode -eq "Partner") { "$($script:SessionState.TenantName) (GDAP)" } else { "own tenant" }
    Write-InfoMsg "Connecting to Microsoft Graph ($targetLabel)..."

    try {
        $params = @{ Scopes = $script:MgScopes; NoWelcome = $true }
        if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantId) { $params["TenantId"] = $script:SessionState.TenantId }
        Connect-MgGraph @params -ErrorAction Stop
        $script:SessionState.MgGraph = $true
        $ctx = Get-MgContext
        Write-Success "Microsoft Graph connected as $($ctx.Account)"
        return $true
    } catch {
        Write-ErrorMsg "Microsoft Graph connection failed: $_"
        return $false
    }
}

function Connect-EXO {
    if ($script:SessionState.ExchangeOnline) { Write-InfoMsg "Exchange Online already connected."; return $true }

    $targetLabel = if ($script:SessionState.TenantMode -eq "Partner") { "$($script:SessionState.TenantName) (GDAP)" } else { "own tenant" }
    Write-InfoMsg "Connecting to Exchange Online ($targetLabel)..."

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
        $errMsg = $_.Exception.Message
        Write-Warn "Browser login failed: $errMsg"
    }

    if ($errMsg -match "BrokerExtension|WithBroker|Method not found") {
        Write-Warn "MSAL broker DLL conflict detected."
        if (Confirm-Action "Auto-fix: disable the conflicting broker DLL and retry?") {
            $exoMod = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
            if ($exoMod) {
                $brokerDlls = @(Get-ChildItem -Path $exoMod.ModuleBase -Recurse -Filter "Microsoft.Identity.Client.Broker.dll" -ErrorAction SilentlyContinue)
                foreach ($dll in $brokerDlls) {
                    try { if (Test-Path $dll.FullName) { Rename-Item -Path $dll.FullName -NewName "$($dll.Name).bak" -Force -ErrorAction Stop; Write-Success "Disabled: $($dll.FullName)" } }
                    catch { Write-Warn "Could not rename: $_" }
                }
            }
            try {
                Connect-ExchangeOnline @exoParams -ErrorAction Stop
                $script:SessionState.ExchangeOnline = $true
                Write-Success "Exchange Online connected (broker disabled)."
                return $true
            } catch {
                Write-Warn "DLL already in memory. Restarting..."
                $mainScript = Join-Path $PSScriptRoot "Main.ps1"
                if (-not $PSScriptRoot) { $mainScript = Join-Path (Split-Path -Parent $MyInvocation.ScriptName) "Main.ps1" }
                Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$mainScript`""
                Start-Sleep -Seconds 2; exit
            }
        }
    }

    Write-ErrorMsg "Exchange Online connection failed."
    return $false
}

function Connect-SCC {
    if ($script:SessionState.ComplianceCenter) { Write-InfoMsg "SCC already connected."; return $true }

    $targetLabel = if ($script:SessionState.TenantMode -eq "Partner") { "$($script:SessionState.TenantName) (GDAP)" } else { "own tenant" }
    Write-InfoMsg "Connecting to Security & Compliance ($targetLabel)..."

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
        $errMsg = $_.Exception.Message
        Write-Warn "SCC connect failed: $errMsg"
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
#  Per-task connection sets
# ============================================================

function Connect-ForTask {
    param(
        [ValidateSet(
            "Onboard","Offboard","License","Archive",
            "SecurityGroup","DistributionList","SharedMailbox","CalendarAccess",
            "UserProfile","Report","eDiscovery","GroupManager"
        )]
        [string]$Task
    )

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
        eDiscovery       = @("SCC","Graph")
        GroupManager     = @("EXO","Graph")
    }

    $services = $map[$Task]
    $needed = @()
    foreach ($svc in $services) {
        switch ($svc) {
            "Graph" { if (-not $script:SessionState.MgGraph)          { $needed += $svc } }
            "EXO"   { if (-not $script:SessionState.ExchangeOnline)   { $needed += $svc } }
            "SCC"   { if (-not $script:SessionState.ComplianceCenter) { $needed += $svc } }
        }
    }

    if ($needed.Count -eq 0) { Write-InfoMsg "All required services connected."; return $true }

    Write-InfoMsg "Requires: $($services -join ', '). Connecting: $($needed -join ', ')"
    Write-Host ""

    foreach ($svc in $services) {
        switch ($svc) {
            "Graph" { if (-not (Connect-Graph)) { return $false } }
            "EXO"   { if (-not (Connect-EXO))   { return $false } }
            "SCC"   { if (-not (Connect-SCC))    { return $false } }
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

    $script:SessionState.MgGraph = $false
    $script:SessionState.ExchangeOnline = $false
    $script:SessionState.ComplianceCenter = $false
    $script:SessionState.PartnerConnected = $false
    Write-Success "All sessions cleared."
}

function Get-TenantDisplayString {
    if ($script:SessionState.TenantMode -eq "Partner") { return "GDAP: $($script:SessionState.TenantName)" }
    return "Direct (own tenant)"
}
