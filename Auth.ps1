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

function Assert-ModulesInstalled {
    Write-SectionHeader "Checking Dependencies"

    # ---- Define required modules with test commands ----
    # IMPORTANT: ExchangeOnlineManagement MUST be first.
    # It loads its MSAL assemblies first, preventing version conflicts
    # when Graph modules try to load a different MSAL version later.
    $requiredModules = @(
        @{ Name = "ExchangeOnlineManagement";                    TestCmd = "Connect-ExchangeOnline" },
        @{ Name = "Microsoft.Graph.Authentication";              TestCmd = "Get-MgContext" },
        @{ Name = "Microsoft.Graph.Users";                       TestCmd = "Get-MgUser" },
        @{ Name = "Microsoft.Graph.Users.Actions";               TestCmd = $null },
        @{ Name = "Microsoft.Graph.Groups";                      TestCmd = "Get-MgGroup" },
        @{ Name = "Microsoft.Graph.Identity.DirectoryManagement"; TestCmd = "Get-MgSubscribedSku" },
        @{ Name = "Microsoft.Online.SharePoint.PowerShell";      TestCmd = "Connect-SPOService"; Optional = $true }
    )

    $allGood = $true

    foreach ($mod in $requiredModules) {
        $modName  = $mod.Name
        $optional = [bool]($mod.Optional)

        # ---- Step 1: Check if installed ----
        $installed = Get-Module -ListAvailable -Name $modName | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $installed) {
            if ($optional) {
                Write-InfoMsg "$modName not installed (optional -- SPO / OneDrive features will be unavailable until installed)."
                continue
            }
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

    # Pre-merge fix: ensure the audit log filename reflects the
    # newly-chosen tenant. Switch-Tenant calls this already; the
    # legacy partner-center picker path didn't, so an operator who
    # switched tenants via this UI ended up writing entries to the
    # prior tenant's session log.
    if (Get-Command Reset-AuditLogPath -ErrorAction SilentlyContinue) { Reset-AuditLogPath }

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
    $dir = Get-StateDirectory
    if (-not $dir) { return @{} }
    $p = Join-Path $dir 'spo-tenants.json'
    if (-not (Test-Path -LiteralPath $p)) { return @{} }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json
        $h = @{}
        foreach ($prop in $raw.PSObject.Properties) { $h[$prop.Name] = [string]$prop.Value }
        return $h
    } catch { return @{} }
}

function Set-SPOAdminUrlCache {
    param([string]$TenantKey, [string]$AdminUrl)
    $dir = Get-StateDirectory
    if (-not $dir) { return }
    $p = Join-Path $dir 'spo-tenants.json'
    $h = Get-SPOAdminUrlCache
    $h[$TenantKey] = $AdminUrl
    try { ($h | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $p -Encoding UTF8 -Force } catch {}
}

function Connect-SPO {
    <#
        Connect-SPOService wrapper. Prompts for the tenant admin URL
        on first use, caches by tenant key (TenantDomain or TenantId)
        so subsequent runs reuse the same URL silently.
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

    $cache = Get-SPOAdminUrlCache
    $adminUrl = $cache[$tenantKey]
    if (-not $adminUrl) {
        $hint = ""
        if ($script:SessionState.TenantDomain -and $script:SessionState.TenantDomain -like '*.onmicrosoft.com') {
            $shortName = ($script:SessionState.TenantDomain -split '\.')[0]
            $hint = " (e.g. https://$shortName-admin.sharepoint.com)"
        }
        $adminUrl = Read-UserInput "SharePoint admin URL$hint"
        if ([string]::IsNullOrWhiteSpace($adminUrl)) { Write-Warn "Cancelled."; return $false }
        $adminUrl = $adminUrl.Trim().TrimEnd('/')
    }

    Write-InfoMsg "Connecting to SharePoint Online ($adminUrl)..."
    try {
        Connect-SPOService -Url $adminUrl -ErrorAction Stop
        $script:SessionState.SharePointOnline   = $true
        $script:SessionState.SharePointAdminUrl = $adminUrl
        Set-SPOAdminUrlCache -TenantKey $tenantKey -AdminUrl $adminUrl
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
