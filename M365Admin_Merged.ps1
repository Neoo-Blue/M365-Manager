# ============================================================
#  M365 Administration Tool v2.0 (Compiled)
#  Built: 2026-04-03 16:35
# ============================================================

# ============ UI.ps1 ============
# ============================================================
#  UI.ps1 - Shared TUI helpers (colors, menus, confirmations)
#  User lookup uses Microsoft Graph PowerShell SDK
# ============================================================

$script:Colors = @{
    BG          = "DarkBlue"
    Title       = "Cyan"
    Menu        = "White"
    Highlight   = "Yellow"
    Success     = "Green"
    Error       = "Red"
    Warning     = "DarkYellow"
    Info        = "Gray"
    Prompt      = "Magenta"
    Accent      = "DarkCyan"
}

$script:Box = @{
    TL  = [char]0x250C
    TR  = [char]0x2510
    BL  = [char]0x2514
    BR  = [char]0x2518
    H   = [char]0x2500
    V   = [char]0x2502
    DTL = [char]0x2554
    DTR = [char]0x2557
    DBL = [char]0x255A
    DBR = [char]0x255D
    DH  = [char]0x2550
    DV  = [char]0x2551
}

function Initialize-UI {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $Host.UI.RawUI.BackgroundColor = $script:Colors.BG
    Clear-Host
}

function Write-Banner {
    $b = $script:Box
    $w = 56
    $title = "            M365 Administration Tool  v2.0"
    $pad   = $w - $title.Length
    Write-Host ""
    Write-Host ("  " + $b.DTL + [string]::new($b.DH, $w) + $b.DTR) -ForegroundColor $script:Colors.Title
    Write-Host ("  " + $b.DV + $title + (" " * $pad) + $b.DV) -ForegroundColor $script:Colors.Title
    Write-Host ("  " + $b.DBL + [string]::new($b.DH, $w) + $b.DBR) -ForegroundColor $script:Colors.Title
}

function Write-SectionHeader {
    param([string]$Title)
    $b = $script:Box
    $prefix = "  " + $b.TL + [string]::new($b.H, 3) + " " + $Title + " "
    $pad = 62 - $prefix.Length
    if ($pad -lt 1) { $pad = 1 }
    $line = $prefix + [string]::new($b.H, $pad) + $b.TR
    Write-Host ""
    Write-Host $line -ForegroundColor $script:Colors.Title
    Write-Host ""
}

function Write-StatusLine {
    param([string]$Label, [string]$Value, [string]$Color = "White")
    Write-Host "    $Label : " -ForegroundColor $script:Colors.Info -NoNewline
    Write-Host $Value -ForegroundColor $Color
}

function Write-Success  { param([string]$Msg) Write-Host "  [+] $Msg" -ForegroundColor $script:Colors.Success }
function Write-ErrorMsg { param([string]$Msg) Write-Host "  [x] $Msg" -ForegroundColor $script:Colors.Error }
function Write-Warn     { param([string]$Msg) Write-Host "  [!] $Msg" -ForegroundColor $script:Colors.Warning }
function Write-InfoMsg  { param([string]$Msg) Write-Host "  [i] $Msg" -ForegroundColor $script:Colors.Info }

function Read-UserInput {
    param([string]$Prompt)
    Write-Host ""
    Write-Host "  $Prompt" -ForegroundColor $script:Colors.Prompt -NoNewline
    Write-Host ": " -ForegroundColor $script:Colors.Prompt -NoNewline
    return (Read-Host)
}

function Confirm-Action {
    param([string]$Message, [string]$Details = "")
    $b = $script:Box
    $w = 58
    Write-Host ""
    $headerText = " CONFIRMATION REQUIRED "
    $hpad = $w - 2 - $headerText.Length
    if ($hpad -lt 1) { $hpad = 1 }
    Write-Host ("  " + $b.DTL + [string]::new($b.DH, 2) + $headerText + [string]::new($b.DH, $hpad) + $b.DTR) -ForegroundColor $script:Colors.Warning
    Write-Host ("  " + $b.DV + " " + $Message) -ForegroundColor White
    if ($Details) {
        foreach ($line in ($Details -split "`n")) {
            Write-Host ("  " + $b.DV + "   " + $line) -ForegroundColor $script:Colors.Info
        }
    }
    Write-Host ("  " + $b.DBL + [string]::new($b.DH, $w) + $b.DBR) -ForegroundColor $script:Colors.Warning
    Write-Host ""
    Write-Host "  Proceed? [Y/N]" -ForegroundColor $script:Colors.Highlight -NoNewline
    Write-Host ": " -NoNewline
    $answer = Read-Host
    return ($answer -match '^[Yy](es)?$')
}

function Show-Menu {
    param([string]$Title, [string[]]$Options, [string]$BackLabel = "Back to Main Menu")
    Write-SectionHeader $Title
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "    [" -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host ($i + 1) -NoNewline -ForegroundColor $script:Colors.Highlight
        Write-Host "] " -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host $Options[$i] -ForegroundColor $script:Colors.Menu
    }
    Write-Host ""
    Write-Host "    [" -NoNewline -ForegroundColor $script:Colors.Accent
    Write-Host "0" -NoNewline -ForegroundColor $script:Colors.Highlight
    Write-Host "] " -NoNewline -ForegroundColor $script:Colors.Accent
    Write-Host $BackLabel -ForegroundColor $script:Colors.Error
    Write-Host ""
    while ($true) {
        Write-Host "  Select option" -ForegroundColor $script:Colors.Prompt -NoNewline
        Write-Host ": " -NoNewline
        $sel = Read-Host
        if ($sel -match '^\d+$') {
            $num = [int]$sel
            if ($num -eq 0) { return -1 }
            if ($num -ge 1 -and $num -le $Options.Count) { return ($num - 1) }
        }
        Write-ErrorMsg "Invalid selection. Please try again."
    }
}

function Show-MultiSelect {
    param([string]$Title, [string[]]$Options, [string]$Prompt = "Enter selection(s) (e.g. 1,3,5)")
    Write-SectionHeader $Title
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "    [" -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host ($i + 1) -NoNewline -ForegroundColor $script:Colors.Highlight
        Write-Host "] " -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host $Options[$i] -ForegroundColor $script:Colors.Menu
    }
    Write-Host ""
    while ($true) {
        $raw = Read-UserInput $Prompt
        $nums = $raw -split ',' | ForEach-Object { $_.Trim() }
        $valid = $true
        $indices = @()
        foreach ($n in $nums) {
            if ($n -match '^\d+$') {
                $idx = [int]$n
                if ($idx -ge 1 -and $idx -le $Options.Count) { $indices += ($idx - 1) }
                else { $valid = $false; break }
            } else { $valid = $false; break }
        }
        if ($valid -and $indices.Count -gt 0) { return $indices }
        Write-ErrorMsg "Invalid input. Use numbers separated by commas (e.g. 1,3,5)."
    }
}

function Show-UserDataTable {
    param([hashtable]$Data, [string[]]$FieldOrder)
    $b = $script:Box
    Write-Host ""
    Write-Host ("  " + $b.TL + [string]::new($b.H, 54) + $b.TR) -ForegroundColor $script:Colors.Accent
    $idx = 1
    foreach ($field in $FieldOrder) {
        $val = if ($Data.ContainsKey($field)) { $Data[$field] } else { "(empty)" }
        $label = ("{0,3}. {1,-20}" -f $idx, $field)
        Write-Host ("  " + $b.V + " ") -ForegroundColor $script:Colors.Accent -NoNewline
        Write-Host $label -ForegroundColor $script:Colors.Info -NoNewline
        Write-Host ": " -NoNewline
        Write-Host ("{0,-28}" -f $val) -ForegroundColor White -NoNewline
        Write-Host (" " + $b.V) -ForegroundColor $script:Colors.Accent
        $idx++
    }
    Write-Host ("  " + $b.BL + [string]::new($b.H, 54) + $b.BR) -ForegroundColor $script:Colors.Accent
    Write-Host ""
}

function Edit-UserDataTable {
    param([hashtable]$Data, [string[]]$FieldOrder)
    while ($true) {
        Show-UserDataTable -Data $Data -FieldOrder $FieldOrder
        $choice = Read-UserInput "Enter field # to edit, or 'ok' to confirm"
        if ($choice -match '^ok$') { return $Data }
        if ($choice -match '^\d+$') {
            $fi = [int]$choice
            if ($fi -ge 1 -and $fi -le $FieldOrder.Count) {
                $fieldName = $FieldOrder[$fi - 1]
                $newVal = Read-UserInput "New value for '$fieldName'"
                $Data[$fieldName] = $newVal
            } else { Write-ErrorMsg "Invalid field number." }
        } else { Write-ErrorMsg "Type a field number or 'ok'." }
    }
}

function Resolve-UserIdentity {
    <#
    .SYNOPSIS
        Asks for name or email, searches Microsoft Graph, confirms the right user.
        Returns a user object or $null.
    #>
    param([string]$PromptText = "Enter user name or email")

    $searchInput = Read-UserInput $PromptText
    if ([string]::IsNullOrWhiteSpace($searchInput)) { return $null }

    Write-InfoMsg "Searching for user..."

    try {
        if ($searchInput -match '@') {
            $user = Get-MgUser -UserId $searchInput -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department,GivenName,Surname,CompanyName,OfficeLocation,StreetAddress,City,State,PostalCode,Country,UsageLocation,BusinessPhones,MobilePhone,Mail,MailNickname,AccountEnabled,UserType" -ErrorAction Stop
        } else {
            # Use Search with ConsistencyLevel for partial matching
            $users = @(Get-MgUser -Search "displayName:$searchInput" -ConsistencyLevel eventual -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department" -ErrorAction Stop)
            if ($users.Count -eq 0) {
                # Fallback: try filter with startsWith
                $users = @(Get-MgUser -Filter "startsWith(displayName,'$searchInput')" -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department" -ErrorAction Stop)
            }
            if ($users.Count -eq 0) {
                Write-ErrorMsg "No users found matching '$searchInput'."
                return $null
            }
            if ($users.Count -eq 1) {
                # Fetch full properties
                $user = Get-MgUser -UserId $users[0].Id -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department,GivenName,Surname,CompanyName,OfficeLocation,StreetAddress,City,State,PostalCode,Country,UsageLocation,BusinessPhones,MobilePhone,Mail,MailNickname,AccountEnabled,UserType" -ErrorAction Stop
            } else {
                Write-Warn "Multiple users found:"
                $names = $users | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
                $sel = Show-Menu -Title "Select User" -Options $names -BackLabel "Cancel"
                if ($sel -eq -1) { return $null }
                $user = Get-MgUser -UserId $users[$sel].Id -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department,GivenName,Surname,CompanyName,OfficeLocation,StreetAddress,City,State,PostalCode,Country,UsageLocation,BusinessPhones,MobilePhone,Mail,MailNickname,AccountEnabled,UserType" -ErrorAction Stop
            }
        }

        Write-Host ""
        Write-StatusLine "Display Name" $user.DisplayName "White"
        Write-StatusLine "UPN" $user.UserPrincipalName "White"
        Write-StatusLine "Job Title" $user.JobTitle "White"
        Write-StatusLine "Department" $user.Department "White"
        Write-Host ""

        if (-not (Confirm-Action "Is this the correct user?")) { return $null }
        return $user
    }
    catch {
        Write-ErrorMsg "Could not find user: $_"
        return $null
    }
}

function Pause-ForUser {
    Write-Host ""
    Write-Host "  Press any key to continue..." -ForegroundColor $script:Colors.Info
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ---- License SKU friendly name mapping ----
$script:SkuFriendlyNames = @{
    "SPE_E3"                    = "Microsoft 365 E3"
    "SPE_E5"                    = "Microsoft 365 E5"
    "SPE_F1"                    = "Microsoft 365 F3"
    "M365_F1"                   = "Microsoft 365 F1"
    "ENTERPRISEPACK"            = "Office 365 E3"
    "ENTERPRISEPREMIUM"         = "Office 365 E5"
    "STANDARDPACK"              = "Office 365 E1"
    "DESKLESSPACK"              = "Office 365 F3"
    "O365_BUSINESS_PREMIUM"     = "Microsoft 365 Business Standard"
    "O365_BUSINESS_ESSENTIALS"  = "Microsoft 365 Business Basic"
    "SMB_BUSINESS"              = "Microsoft 365 Apps for Business"
    "SMB_BUSINESS_PREMIUM"      = "Microsoft 365 Business Premium"
    "OFFICESUBSCRIPTION"        = "Microsoft 365 Apps for Enterprise"
    "EXCHANGESTANDARD"          = "Exchange Online (Plan 1)"
    "EXCHANGEENTERPRISE"        = "Exchange Online (Plan 2)"
    "EXCHANGEARCHIVE_ADDON"     = "Exchange Online Archiving"
    "EXCHANGEARCHIVE"           = "Exchange Online Archiving"
    "EMS_E3"                    = "Enterprise Mobility + Security E3"
    "EMS_E5"                    = "Enterprise Mobility + Security E5"
    "EMSPREMIUM"                = "Enterprise Mobility + Security E5"
    "AAD_PREMIUM"               = "Azure AD Premium P1"
    "AAD_PREMIUM_P2"            = "Azure AD Premium P2"
    "INTUNE_A"                  = "Microsoft Intune Plan 1"
    "INTUNE_P1"                 = "Microsoft Intune Plan 1"
    "ATP_ENTERPRISE"            = "Microsoft Defender for Office 365 P1"
    "THREAT_INTELLIGENCE"       = "Microsoft Defender for Office 365 P2"
    "WIN_DEF_ATP"               = "Microsoft Defender for Endpoint P2"
    "MDATP_XPLAT"               = "Microsoft Defender for Endpoint P2"
    "IDENTITY_THREAT_PROTECTION"= "Microsoft 365 E5 Security"
    "INFORMATION_PROTECTION_COMPLIANCE" = "Microsoft 365 E5 Compliance"
    "POWER_BI_STANDARD"         = "Power BI (Free)"
    "POWER_BI_PRO"              = "Power BI Pro"
    "POWER_BI_PREMIUM_PER_USER" = "Power BI Premium Per User"
    "FLOW_FREE"                 = "Power Automate (Free)"
    "POWERAPPS_VIRAL"           = "Power Apps (Free)"
    "POWERAPPS_PER_USER"        = "Power Apps Per User"
    "TEAMS_EXPLORATORY"         = "Microsoft Teams Exploratory"
    "TEAMS_FREE"                = "Microsoft Teams (Free)"
    "PROJECTPROFESSIONAL"       = "Project Plan 3"
    "PROJECTPREMIUM"            = "Project Plan 5"
    "PROJECT_P1"                = "Project Plan 1"
    "VISIOONLINE_PLAN1"         = "Visio Plan 1"
    "VISIOCLIENT"               = "Visio Plan 2"
    "MCOEV"                     = "Teams Phone Standard"
    "MCOPSTN1"                  = "Domestic Calling Plan"
    "MCOPSTN2"                  = "International Calling Plan"
    "MCOCAP"                    = "Common Area Phone"
    "PHONESYSTEM_VIRTUALUSER"   = "Teams Phone Resource Account"
    "MEETING_ROOM"              = "Teams Rooms Standard"
    "RIGHTSMANAGEMENT"          = "Azure Information Protection P1"
    "STREAM"                    = "Microsoft Stream"
    "DEVELOPERPACK"             = "Office 365 E3 Developer"
    "WINDOWS_STORE"             = "Windows Store for Business"
    "FORMS_PRO"                 = "Dynamics 365 Customer Voice"
    "D365_SALES_ENT"            = "Dynamics 365 Sales Enterprise"
    "CRMSTANDARD"               = "Dynamics 365 Professional"
    "SHAREPOINTSTANDARD"        = "SharePoint Online (Plan 1)"
    "SHAREPOINTENTERPRISE"      = "SharePoint Online (Plan 2)"
    "PLANNERSTANDALONE"         = "Planner Plan 1"
    "MICROSOFT_BUSINESS_CENTER" = "Microsoft Business Center"
    "DYN365_ENTERPRISE_PLAN1"   = "Dynamics 365 Plan"
    "SPB"                       = "Microsoft 365 Business Premium"
    "ENTERPRISEWITHSCAL"        = "Office 365 E4"
    "MIDSIZEPACK"               = "Office 365 Midsize Business"
    "LITEPACK"                  = "Office 365 Small Business"
    "WACONEDRIVESTANDARD"       = "OneDrive for Business (Plan 1)"
    "WACONEDRIVEENTERPRISE"     = "OneDrive for Business (Plan 2)"
    "CCIBOTS_PRIVPREV_VIRAL"    = "Power Virtual Agents Viral Trial"
    "CDS_DB_CAPACITY"           = "Common Data Service DB Capacity"
    "PBI_PREMIUM_EM1_ADDON"     = "Power BI Premium EM1"
    "NONPROFIT_PORTAL"          = "Nonprofit Portal"
}

function Get-SkuFriendlyName {
    param([string]$SkuPartNumber)
    if ($script:SkuFriendlyNames.ContainsKey($SkuPartNumber)) {
        return $script:SkuFriendlyNames[$SkuPartNumber]
    }
    # Try to make it more readable by replacing underscores and title-casing
    return $SkuPartNumber
}

function Format-LicenseLabel {
    <#
    .SYNOPSIS
        Returns "Friendly Name (SKU_PART_NUMBER)" for display
    #>
    param([string]$SkuPartNumber)
    $friendly = Get-SkuFriendlyName $SkuPartNumber
    if ($friendly -eq $SkuPartNumber) {
        return $SkuPartNumber
    }
    return "$friendly ($SkuPartNumber)"
}


# ============ Auth.ps1 ============
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
            "UserProfile"
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


# ============ Onboard.ps1 ============
# ============================================================
#  Onboard.ps1 - New User Onboarding (Microsoft Graph)
# ============================================================

function Start-Onboard {
    Write-SectionHeader "User Onboarding"

    $fields = @(
        "FirstName","LastName","DisplayName","UserPrincipalName",
        "JobTitle","Department","CompanyName","OfficeLocation",
        "StreetAddress","City","State","PostalCode","Country",
        "UsageLocation","BusinessPhone","MobilePhone","Manager"
    )

    $replicateSource   = $null
    $replicateLicenses = @()
    $replicateGroups   = @()

    $choice = Show-Menu -Title "How to provide user data?" -Options @(
        "Parse from a text file",
        "Enter manually",
        "Replicate from an existing user"
    ) -BackLabel "Cancel"
    if ($choice -eq -1) { return }

    $userData = @{}

    if ($choice -eq 0) {
        $filePath = Read-UserInput "Enter full path to the text file"
        if (-not (Test-Path $filePath)) { Write-ErrorMsg "File not found: $filePath"; Pause-ForUser; return }
        Write-InfoMsg "Parsing file..."
        $lines = Get-Content $filePath
        foreach ($line in $lines) {
            if ($line -match '^\s*([^:=]+)\s*[:=]\s*(.+)$') {
                $key = $Matches[1].Trim(); $value = $Matches[2].Trim()
                switch -Regex ($key) {
                    'first\s*name'              { $userData["FirstName"]         = $value }
                    'last\s*name'               { $userData["LastName"]          = $value }
                    'display\s*name'            { $userData["DisplayName"]       = $value }
                    'upn|email|user\s*principal' { $userData["UserPrincipalName"] = $value }
                    'title|job'                 { $userData["JobTitle"]          = $value }
                    'depart'                    { $userData["Department"]        = $value }
                    'company'                   { $userData["CompanyName"]       = $value }
                    'office'                    { $userData["OfficeLocation"]    = $value }
                    'street'                    { $userData["StreetAddress"]     = $value }
                    'city'                      { $userData["City"]              = $value }
                    'state'                     { $userData["State"]             = $value }
                    'postal|zip'                { $userData["PostalCode"]        = $value }
                    'country'                   { $userData["Country"]           = $value }
                    'usage\s*location'          { $userData["UsageLocation"]     = $value }
                    'business\s*phone'          { $userData["BusinessPhone"]     = $value }
                    'mobile'                    { $userData["MobilePhone"]       = $value }
                    'manager'                   { $userData["Manager"]           = $value }
                    default                     { $userData[$key]                = $value }
                }
            }
        }
        foreach ($f in $fields) { if (-not $userData.ContainsKey($f)) { $userData[$f] = "" } }
        if ([string]::IsNullOrWhiteSpace($userData["DisplayName"]) -and $userData["FirstName"] -and $userData["LastName"]) {
            $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
        }
        Write-Success "Parsed data:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }
    elseif ($choice -eq 1) {
        foreach ($f in $fields) { $userData[$f] = Read-UserInput "Enter $f" }
        if ([string]::IsNullOrWhiteSpace($userData["DisplayName"]) -and $userData["FirstName"] -and $userData["LastName"]) {
            $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
        }
        Write-InfoMsg "Review the information:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }
    elseif ($choice -eq 2) {
        if (-not (Connect-ForTask "Onboard")) { Write-ErrorMsg "Could not connect."; Pause-ForUser; return }

        Write-SectionHeader "Select User to Replicate From"
        $replicateSource = Resolve-UserIdentity -PromptText "Enter the existing user's name or email"
        if ($null -eq $replicateSource) { Write-ErrorMsg "No source user."; Pause-ForUser; return }

        $src = $replicateSource
        $userData["FirstName"] = ""; $userData["LastName"] = ""; $userData["DisplayName"] = ""; $userData["UserPrincipalName"] = ""
        $userData["JobTitle"]       = if ($src.JobTitle)       { $src.JobTitle } else { "" }
        $userData["Department"]     = if ($src.Department)     { $src.Department } else { "" }
        $userData["CompanyName"]    = if ($src.CompanyName)    { $src.CompanyName } else { "" }
        $userData["OfficeLocation"] = if ($src.OfficeLocation) { $src.OfficeLocation } else { "" }
        $userData["StreetAddress"]  = if ($src.StreetAddress)  { $src.StreetAddress } else { "" }
        $userData["City"]           = if ($src.City)           { $src.City } else { "" }
        $userData["State"]          = if ($src.State)          { $src.State } else { "" }
        $userData["PostalCode"]     = if ($src.PostalCode)     { $src.PostalCode } else { "" }
        $userData["Country"]        = if ($src.Country)        { $src.Country } else { "" }
        $userData["UsageLocation"]  = if ($src.UsageLocation)  { $src.UsageLocation } else { "" }
        $userData["BusinessPhone"]  = if ($src.BusinessPhones.Count -gt 0) { $src.BusinessPhones[0] } else { "" }
        $userData["MobilePhone"]    = if ($src.MobilePhone)    { $src.MobilePhone } else { "" }

        try {
            $srcMgr = Get-MgUserManager -UserId $src.Id -ErrorAction SilentlyContinue
            $userData["Manager"] = if ($srcMgr) { $srcMgr.AdditionalProperties["userPrincipalName"] } else { "" }
        } catch { $userData["Manager"] = "" }

        Write-InfoMsg "Reading licenses..."
        try { $replicateLicenses = @(Get-MgUserLicenseDetail -UserId $src.Id -ErrorAction Stop) } catch { $replicateLicenses = @() }

        Write-InfoMsg "Reading security groups..."
        try {
            $allMemberships = @(Get-MgUserMemberOf -UserId $src.Id -All -ErrorAction Stop)
            $replicateGroups = @($allMemberships | Where-Object {
                $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group" -and
                $_.AdditionalProperties["securityEnabled"] -eq $true
            })
        } catch { $replicateGroups = @() }

        Write-SectionHeader "Replicated from $($src.DisplayName)"
        if ($replicateLicenses.Count -gt 0) {
            Write-InfoMsg "Licenses:"; $replicateLicenses | ForEach-Object { Write-Host "    - $(Format-LicenseLabel $_.SkuPartNumber)" -ForegroundColor White }
        }
        if ($replicateGroups.Count -gt 0) {
            Write-InfoMsg "Security groups:"; $replicateGroups | ForEach-Object { Write-Host "    - $($_.AdditionalProperties['displayName'])" -ForegroundColor White }
        }
        Write-Host ""
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }

    # Validate
    foreach ($r in @("FirstName","LastName","UserPrincipalName","UsageLocation")) {
        if ([string]::IsNullOrWhiteSpace($userData[$r])) { Write-ErrorMsg "'$r' is required."; Pause-ForUser; return }
    }
    if ([string]::IsNullOrWhiteSpace($userData["DisplayName"])) {
        $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
    }

    if (-not (Connect-ForTask "Onboard")) { Pause-ForUser; return }

    # ---- Create Account ----
    Write-SectionHeader "Step 1 - Create Account"
    $password = -join ((48..57) + (65..90) + (97..122) + (33,35,36,37,38) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
    $mailNick = ($userData["UserPrincipalName"] -split '@')[0]

    $details = "UPN: $($userData['UserPrincipalName'])`nName: $($userData['DisplayName'])`nDept: $($userData['Department'])"
    if (-not (Confirm-Action "Create this user account?" $details)) { Pause-ForUser; return }

    try {
        $body = @{
            AccountEnabled    = $true
            DisplayName       = $userData["DisplayName"]
            GivenName         = $userData["FirstName"]
            Surname           = $userData["LastName"]
            UserPrincipalName = $userData["UserPrincipalName"]
            MailNickname      = $mailNick
            UsageLocation     = $userData["UsageLocation"]
            PasswordProfile   = @{ ForceChangePasswordNextSignIn = $true; Password = $password }
        }
        if ($userData["JobTitle"])       { $body["JobTitle"]       = $userData["JobTitle"] }
        if ($userData["Department"])     { $body["Department"]     = $userData["Department"] }
        if ($userData["CompanyName"])    { $body["CompanyName"]    = $userData["CompanyName"] }
        if ($userData["OfficeLocation"]) { $body["OfficeLocation"] = $userData["OfficeLocation"] }
        if ($userData["StreetAddress"])  { $body["StreetAddress"]  = $userData["StreetAddress"] }
        if ($userData["City"])           { $body["City"]           = $userData["City"] }
        if ($userData["State"])          { $body["State"]          = $userData["State"] }
        if ($userData["PostalCode"])     { $body["PostalCode"]     = $userData["PostalCode"] }
        if ($userData["Country"])        { $body["Country"]        = $userData["Country"] }
        if ($userData["MobilePhone"])    { $body["MobilePhone"]    = $userData["MobilePhone"] }
        if ($userData["BusinessPhone"])  { $body["BusinessPhones"] = @($userData["BusinessPhone"]) }

        $newUser = New-MgUser -BodyParameter $body -ErrorAction Stop
        Write-Success "Account created: $($userData['UserPrincipalName'])"
    } catch { Write-ErrorMsg "Failed to create account: $_"; Pause-ForUser; return }

    # ---- Licenses ----
    Write-SectionHeader "Step 2 - Assign Licenses"
    Start-Sleep -Seconds 5
    if ($replicateLicenses.Count -gt 0) {
        foreach ($lic in $replicateLicenses) {
            if (Confirm-Action "Assign license '$(Format-LicenseLabel $lic.SkuPartNumber)'?") {
                try { Set-MgUserLicense -UserId $newUser.Id -AddLicenses @(@{SkuId = $lic.SkuId}) -RemoveLicenses @() -ErrorAction Stop; Write-Success "'$(Get-SkuFriendlyName $lic.SkuPartNumber)' assigned." }
                catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        $more = Read-UserInput "Add additional licenses? (y/n)"
        if ($more -match '^[Yy]') { Select-AndAssignLicenses -UserId $newUser.Id }
    } else { Select-AndAssignLicenses -UserId $newUser.Id }

    # ---- Security Groups ----
    Write-SectionHeader "Step 3 - Assign Security Groups"
    if ($replicateGroups.Count -gt 0) {
        foreach ($grp in $replicateGroups) {
            $gName = $grp.AdditionalProperties["displayName"]
            if (Confirm-Action "Add to group '$gName'?") {
                try { New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $newUser.Id -ErrorAction Stop; Write-Success "Added to '$gName'." }
                catch { if ($_.Exception.Message -match "already exist") { Write-Warn "Already a member." } else { Write-ErrorMsg "Failed: $_" } }
            }
        }
        $more = Read-UserInput "Add additional groups? (y/n)"
        if ($more -match '^[Yy]') { Search-AndAssignGroups -UserId $newUser.Id }
    } else { Search-AndAssignGroups -UserId $newUser.Id }

    # ---- Manager ----
    if ($userData["Manager"]) {
        Write-SectionHeader "Step 4 - Set Manager"
        try {
            $mgr = Get-MgUser -UserId $userData["Manager"] -ErrorAction Stop
            if (Confirm-Action "Set manager to $($mgr.DisplayName)?") {
                Set-MgUserManagerByRef -UserId $newUser.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)" } -ErrorAction Stop
                Write-Success "Manager set."
            }
        } catch { Write-Warn "Could not set manager: $_" }
    }

    # ---- Password ----
    Write-SectionHeader "Step 5 - Temporary Password"
    $b = $script:Box
    Write-Host ""; Write-Host ("  " + $b.TL + [string]::new($b.H, 48) + $b.TR) -ForegroundColor $script:Colors.Highlight
    Write-Host ("  " + $b.V + "  User : $($userData['UserPrincipalName'])") -ForegroundColor White
    Write-Host ("  " + $b.V + "  Pass : $password") -ForegroundColor $script:Colors.Success
    Write-Host ("  " + $b.V + "  (User must change on first login)") -ForegroundColor $script:Colors.Info
    Write-Host ("  " + $b.BL + [string]::new($b.H, 48) + $b.BR) -ForegroundColor $script:Colors.Highlight
    Write-Warn "Securely deliver these credentials to the user."
    Write-Success "Onboarding complete!"
    Pause-ForUser
}

function Select-AndAssignLicenses {
    param([string]$UserId)
    try {
        $skus = Get-MgSubscribedSku -ErrorAction Stop
        $available = $skus | ForEach-Object {
            $total = $_.PrepaidUnits.Enabled; $used = $_.ConsumedUnits
            [PSCustomObject]@{ SkuPartNumber = $_.SkuPartNumber; SkuId = $_.SkuId; FriendlyName = Format-LicenseLabel $_.SkuPartNumber; Total = $total; Used = $used; Free = $total - $used }
        }
        $labels = $available | ForEach-Object { "$($_.FriendlyName)  [Total: $($_.Total) | Used: $($_.Used) | Free: $($_.Free)]" }
        $selected = Show-MultiSelect -Title "Select license(s)" -Options $labels
        foreach ($idx in $selected) {
            $sku = $available[$idx]
            if ($sku.Free -le 0) { Write-Warn "$(Get-SkuFriendlyName $sku.SkuPartNumber) has no seats."; continue }
            if (Confirm-Action "Assign '$(Get-SkuFriendlyName $sku.SkuPartNumber)'?") {
                Set-MgUserLicense -UserId $UserId -AddLicenses @(@{SkuId = $sku.SkuId}) -RemoveLicenses @() -ErrorAction Stop
                Write-Success "'$(Get-SkuFriendlyName $sku.SkuPartNumber)' assigned."
            }
        }
    } catch { Write-ErrorMsg "License error: $_" }
}

function Search-AndAssignGroups {
    param([string]$UserId)
    $searchInput = Read-UserInput "Search for a security group (or 'skip')"
    if ($searchInput -eq 'skip') { return }
    try {
        $groups = @(Get-MgGroup -Search "displayName:$searchInput" -ConsistencyLevel eventual -ErrorAction Stop | Where-Object { $_.SecurityEnabled })
        if ($groups.Count -eq 0) { Write-Warn "No groups found."; return }
        $gLabels = $groups | ForEach-Object { $_.DisplayName }
        $gSel = Show-MultiSelect -Title "Select group(s)" -Options $gLabels
        foreach ($gi in $gSel) {
            $grp = $groups[$gi]
            if (Confirm-Action "Add to '$($grp.DisplayName)'?") {
                try { New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $UserId -ErrorAction Stop; Write-Success "Added." }
                catch { if ($_.Exception.Message -match "already exist") { Write-Warn "Already a member." } else { Write-ErrorMsg "Failed: $_" } }
            }
        }
    } catch { Write-ErrorMsg "Group error: $_" }
}


# ============ Offboard.ps1 ============
# ============================================================
#  Offboard.ps1 - User Offboarding (Microsoft Graph)
# ============================================================

function Start-Offboard {
    Write-SectionHeader "User Offboarding"

    $fields = @("UserPrincipalName","ForwardingEmail","OOOMessage")
    $choice = Show-Menu -Title "How to provide offboarding data?" -Options @("Parse from a text file","Enter manually") -BackLabel "Cancel"
    if ($choice -eq -1) { return }

    $userData = @{}
    if ($choice -eq 0) {
        $filePath = Read-UserInput "Enter full path to the text file"
        if (-not (Test-Path $filePath)) { Write-ErrorMsg "File not found."; Pause-ForUser; return }
        foreach ($line in (Get-Content $filePath)) {
            if ($line -match '^\s*([^:=]+)\s*[:=]\s*(.+)$') {
                $key = $Matches[1].Trim(); $value = $Matches[2].Trim()
                switch -Regex ($key) {
                    'upn|email|user'               { $userData["UserPrincipalName"] = $value }
                    'forward'                      { $userData["ForwardingEmail"]   = $value }
                    'ooo|out.of.office|auto.reply'  { $userData["OOOMessage"]        = $value }
                }
            }
        }
        foreach ($f in $fields) { if (-not $userData.ContainsKey($f)) { $userData[$f] = "" } }
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    } else {
        $userData["UserPrincipalName"] = Read-UserInput "Enter User Principal Name (email)"
        $userData["ForwardingEmail"] = ""; $userData["OOOMessage"] = ""
    }

    $upn = $userData["UserPrincipalName"]
    if ([string]::IsNullOrWhiteSpace($upn)) { Write-ErrorMsg "UPN is required."; Pause-ForUser; return }

    if (-not (Connect-ForTask "Offboard")) { Pause-ForUser; return }

    try {
        $user = Get-MgUser -UserId $upn -Property "Id,DisplayName,UserPrincipalName" -ErrorAction Stop
        Write-Success "User found: $($user.DisplayName) ($upn)"
    } catch { Write-ErrorMsg "User not found: $_"; Pause-ForUser; return }

    if (-not (Confirm-Action "Begin offboarding for $($user.DisplayName)?")) { Pause-ForUser; return }

    # Step 1: Revoke sessions
    Write-SectionHeader "Step 1 - Revoke All Sessions"
    if (Confirm-Action "Revoke all sessions for $upn?") {
        try { Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop; Write-Success "Sessions revoked." }
        catch { Write-ErrorMsg "Failed: $_" }
    }

    # Step 2: Block sign-in
    Write-SectionHeader "Step 2 - Block Sign-In"
    if (Confirm-Action "Block sign-in for $upn?") {
        try { Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop; Write-Success "Sign-in blocked." }
        catch { Write-ErrorMsg "Failed: $_" }
    }

    # Step 3: OOO
    Write-SectionHeader "Step 3 - Out-of-Office"
    if ([string]::IsNullOrWhiteSpace($userData["OOOMessage"])) { $userData["OOOMessage"] = Read-UserInput "Enter OOO message (or 'skip')" }
    if ($userData["OOOMessage"] -ne 'skip' -and $userData["OOOMessage"]) {
        if (Confirm-Action "Set auto-reply?" "Message: $($userData['OOOMessage'])") {
            try {
                Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $userData["OOOMessage"] -ExternalMessage $userData["OOOMessage"] -ErrorAction Stop
                Write-Success "OOO set."
            } catch { Write-ErrorMsg "Failed: $_" }
        }
    }

    # Step 4: Forwarding
    Write-SectionHeader "Step 4 - Email Forwarding"
    if ([string]::IsNullOrWhiteSpace($userData["ForwardingEmail"])) { $userData["ForwardingEmail"] = Read-UserInput "Forwarding email (or 'skip')" }
    if ($userData["ForwardingEmail"] -ne 'skip' -and $userData["ForwardingEmail"]) {
        $fwdEmail = $userData["ForwardingEmail"]
        try { $fwdUser = Get-MgUser -UserId $fwdEmail -ErrorAction Stop; Write-Success "Target found: $($fwdUser.DisplayName)" }
        catch { Write-Warn "'$fwdEmail' not found in tenant."; if (-not (Confirm-Action "Use anyway?")) { $fwdEmail = $null } }

        if ($fwdEmail) {
            $dc = Show-Menu -Title "Delivery option" -Options @("Forward only","Forward and keep copy") -BackLabel "Skip"
            if ($dc -ne -1) {
                $keepCopy = ($dc -eq 1)
                if (Confirm-Action "Set forwarding to $fwdEmail (keep copy: $keepCopy)?") {
                    try { Set-Mailbox -Identity $upn -ForwardingSmtpAddress "smtp:$fwdEmail" -DeliverToMailboxAndForward $keepCopy -ErrorAction Stop; Write-Success "Forwarding set." }
                    catch { Write-ErrorMsg "Failed: $_" }
                }
            }
        }
    }

    # Step 5: Shared mailbox
    Write-SectionHeader "Step 5 - Convert to Shared Mailbox"
    if (Confirm-Action "Convert $upn to Shared Mailbox?") {
        try { Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop; Write-Success "Converted." }
        catch { Write-ErrorMsg "Failed: $_" }
    }

    # Step 6: Grant access
    Write-SectionHeader "Step 6 - Grant Mailbox Access"
    $ga = Show-Menu -Title "Anyone need access?" -Options @("Yes","No") -BackLabel "Skip"
    if ($ga -eq 0) {
        $adding = $true
        while ($adding) {
            $ai = Read-UserInput "User name or email to grant access"
            if ([string]::IsNullOrWhiteSpace($ai)) { break }
            try {
                $au = if ($ai -match '@') { Get-MgUser -UserId $ai -ErrorAction Stop } else {
                    $found = @(Get-MgUser -Search "displayName:$ai" -ConsistencyLevel eventual -ErrorAction Stop)
                    if ($found.Count -eq 0) { Write-ErrorMsg "Not found."; continue }
                    if ($found.Count -eq 1) { $found[0] } else {
                        $sel = Show-Menu -Title "Select" -Options ($found | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                        if ($sel -eq -1) { continue }; $found[$sel]
                    }
                }
                if (Confirm-Action "Grant Full Access to $($au.DisplayName)?") {
                    try { Add-MailboxPermission -Identity $upn -User $au.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop; Write-Success "Full Access granted." }
                    catch { Write-ErrorMsg "Failed: $_" }
                }
                $pc = Show-Menu -Title "Send permissions for $($au.DisplayName)?" -Options @("Send As","Send on Behalf","Both","None") -BackLabel "Skip"
                if ($pc -ne -1 -and $pc -ne 3) {
                    if ($pc -eq 0 -or $pc -eq 2) { if (Confirm-Action "Grant Send As?") { try { Add-RecipientPermission -Identity $upn -Trustee $au.UserPrincipalName -AccessRights SendAs -Confirm:$false -ErrorAction Stop; Write-Success "Send As granted." } catch { Write-ErrorMsg "$_" } } }
                    if ($pc -eq 1 -or $pc -eq 2) { if (Confirm-Action "Grant Send on Behalf?") { try { Set-Mailbox -Identity $upn -GrantSendOnBehalfTo @{Add = $au.UserPrincipalName} -ErrorAction Stop; Write-Success "Send on Behalf granted." } catch { Write-ErrorMsg "$_" } } }
                }
            } catch { Write-ErrorMsg "Error: $_" }
            $more = Read-UserInput "Grant access to another user? (y/n)"
            if ($more -notmatch '^[Yy]') { $adding = $false }
        }
    }

    # Step 7: Remove licenses
    Write-SectionHeader "Step 7 - Remove Licenses"
    try {
        $lics = @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop)
        if ($lics.Count -eq 0) { Write-InfoMsg "No licenses." }
        else {
            # Check assignment states
            $fullUser = $null
            try { $fullUser = Get-MgUser -UserId $user.Id -Property "LicenseAssignmentStates" -ErrorAction Stop } catch {}
            $assignInfo = @{}
            if ($fullUser -and $fullUser.LicenseAssignmentStates) {
                foreach ($s in $fullUser.LicenseAssignmentStates) {
                    $sid = "$($s.SkuId)"; $ai = @{ Direct = $false; Groups = @() }
                    if ($assignInfo.ContainsKey($sid)) { $ai = $assignInfo[$sid] }
                    if ($null -eq $s.AssignedByGroup -or $s.AssignedByGroup -eq "") { $ai.Direct = $true } else { $ai.Groups += $s.AssignedByGroup }
                    $assignInfo[$sid] = $ai
                }
            }

            Write-InfoMsg "Current licenses:"
            foreach ($lic in $lics) {
                $tag = ""
                $sid = "$($lic.SkuId)"
                if ($assignInfo.ContainsKey($sid) -and $assignInfo[$sid].Groups.Count -gt 0 -and -not $assignInfo[$sid].Direct) { $tag = " [GROUP]" }
                Write-Host "    - $(Format-LicenseLabel $lic.SkuPartNumber)$tag" -ForegroundColor White
            }

            if (Confirm-Action "Remove all directly-assigned licenses?") {
                foreach ($lic in $lics) {
                    $sid = "$($lic.SkuId)"
                    $friendly = Get-SkuFriendlyName $lic.SkuPartNumber
                    if ($assignInfo.ContainsKey($sid) -and $assignInfo[$sid].Groups.Count -gt 0 -and -not $assignInfo[$sid].Direct) {
                        Write-Warn "$friendly is group-assigned. Skipping (remove user from the group instead)."
                        continue
                    }
                    try { Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($lic.SkuId) -ErrorAction Stop; Write-Success "Removed: $friendly" }
                    catch { Write-ErrorMsg "Failed to remove $friendly : $_" }
                }
            }
        }
    } catch { Write-ErrorMsg "License error: $_" }

    Write-Success "Offboarding complete for $($user.DisplayName)!"
    Pause-ForUser
}


# ============ License.ps1 ============
# ============================================================
#  License.ps1 - License Management (Microsoft Graph)
#  Shows friendly names, detects group-assigned licenses
# ============================================================

function Start-LicenseManagement {
    Write-SectionHeader "License Management"
    if (-not (Connect-ForTask "License")) { Pause-ForUser; return }

    $user = Resolve-UserIdentity
    if ($null -eq $user) { Pause-ForUser; return }

    # ---- Get license assignment states (direct vs group) ----
    $fullUser = $null
    try {
        $fullUser = Get-MgUser -UserId $user.Id -Property "Id,DisplayName,UserPrincipalName,LicenseAssignmentStates" -ErrorAction Stop
    } catch {
        Write-Warn "Could not read assignment states: $_"
    }

    # Build a lookup: SkuId -> assignment info
    $assignmentInfo = @{}
    if ($fullUser -and $fullUser.LicenseAssignmentStates) {
        foreach ($state in $fullUser.LicenseAssignmentStates) {
            $skuId = "$($state.SkuId)"
            $info = @{ Direct = $false; Groups = @() }
            if ($assignmentInfo.ContainsKey($skuId)) { $info = $assignmentInfo[$skuId] }

            if ($null -eq $state.AssignedByGroup -or $state.AssignedByGroup -eq "") {
                $info.Direct = $true
            } else {
                $info.Groups += $state.AssignedByGroup
            }
            $assignmentInfo[$skuId] = $info
        }
    }

    # ---- Show current licenses ----
    Write-SectionHeader "Current Licenses for $($user.DisplayName)"

    try { $currentLics = @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop) } catch { $currentLics = @() }

    if ($currentLics.Count -eq 0) {
        Write-InfoMsg "No licenses assigned."
    } else {
        for ($i = 0; $i -lt $currentLics.Count; $i++) {
            $lic = $currentLics[$i]
            $label = Format-LicenseLabel $lic.SkuPartNumber

            # Check assignment type
            $skuId = "$($lic.SkuId)"
            $assignTag = ""
            if ($assignmentInfo.ContainsKey($skuId)) {
                $ai = $assignmentInfo[$skuId]
                if ($ai.Direct -and $ai.Groups.Count -gt 0) {
                    $assignTag = " [Direct + Group]"
                } elseif ($ai.Groups.Count -gt 0) {
                    $assignTag = " [Group-assigned]"
                } else {
                    $assignTag = " [Direct]"
                }
            }

            Write-Host "    [" -NoNewline -ForegroundColor $script:Colors.Accent
            Write-Host ($i + 1) -NoNewline -ForegroundColor $script:Colors.Highlight
            Write-Host "] " -NoNewline -ForegroundColor $script:Colors.Accent
            Write-Host $label -NoNewline -ForegroundColor White
            if ($assignTag -match "Group") {
                Write-Host $assignTag -ForegroundColor $script:Colors.Warning
            } else {
                Write-Host $assignTag -ForegroundColor $script:Colors.Info
            }
        }
    }
    Write-Host ""

    # ---- Add or Remove ----
    $action = Show-Menu -Title "Action" -Options @("Add license(s)", "Remove license(s)") -BackLabel "Cancel"
    if ($action -eq -1) { return }

    if ($action -eq 0) {
        # ---- ADD ----
        Write-SectionHeader "Available Tenant Licenses"
        try {
            $skus = Get-MgSubscribedSku -ErrorAction Stop
            $available = $skus | ForEach-Object {
                $t = $_.PrepaidUnits.Enabled; $u = $_.ConsumedUnits
                [PSCustomObject]@{
                    SkuPartNumber = $_.SkuPartNumber
                    SkuId         = $_.SkuId
                    FriendlyName  = Format-LicenseLabel $_.SkuPartNumber
                    Total         = $t
                    Used          = $u
                    Free          = $t - $u
                }
            }

            $labels = $available | ForEach-Object {
                "$($_.FriendlyName)  [Total: $($_.Total) | Used: $($_.Used) | Free: $($_.Free)]"
            }

            $selected = Show-MultiSelect -Title "Select license(s) to add" -Options $labels
            foreach ($idx in $selected) {
                $sku = $available[$idx]
                if ($sku.Free -le 0) {
                    Write-Warn "$(Get-SkuFriendlyName $sku.SkuPartNumber) has no available seats. Skipping."
                    continue
                }
                if (Confirm-Action "Assign '$(Get-SkuFriendlyName $sku.SkuPartNumber)' to $($user.UserPrincipalName)?") {
                    try {
                        Set-MgUserLicense -UserId $user.Id -AddLicenses @(@{SkuId = $sku.SkuId}) -RemoveLicenses @() -ErrorAction Stop
                        Write-Success "$(Get-SkuFriendlyName $sku.SkuPartNumber) assigned."
                    } catch {
                        Write-ErrorMsg "Failed to assign: $_"
                    }
                }
            }
        } catch { Write-ErrorMsg "License error: $_" }
    }
    else {
        # ---- REMOVE ----
        if ($currentLics.Count -eq 0) {
            Write-Warn "No licenses to remove."
            Pause-ForUser; return
        }

        $labels = @()
        for ($i = 0; $i -lt $currentLics.Count; $i++) {
            $lic = $currentLics[$i]
            $label = Format-LicenseLabel $lic.SkuPartNumber
            $skuId = "$($lic.SkuId)"
            if ($assignmentInfo.ContainsKey($skuId) -and $assignmentInfo[$skuId].Groups.Count -gt 0 -and -not $assignmentInfo[$skuId].Direct) {
                $label += "  !! GROUP-ASSIGNED"
            }
            $labels += $label
        }

        $selected = Show-MultiSelect -Title "Select license(s) to remove" -Options $labels

        foreach ($idx in $selected) {
            $lic = $currentLics[$idx]
            $friendlyName = Get-SkuFriendlyName $lic.SkuPartNumber
            $skuId = "$($lic.SkuId)"

            # Check if group-assigned only
            if ($assignmentInfo.ContainsKey($skuId)) {
                $ai = $assignmentInfo[$skuId]
                if ($ai.Groups.Count -gt 0 -and -not $ai.Direct) {
                    Write-Host ""
                    Write-ErrorMsg "'$friendlyName' is assigned via a group, not directly."
                    Write-Warn "You cannot remove group-assigned licenses from individual users."
                    Write-InfoMsg "To remove this license:"
                    Write-InfoMsg "  1. Remove the user from the licensing group, OR"
                    Write-InfoMsg "  2. Remove this license from the group itself."

                    # Try to resolve group name
                    foreach ($gId in $ai.Groups) {
                        try {
                            $grp = Get-MgGroup -GroupId $gId -Property "DisplayName" -ErrorAction Stop
                            Write-InfoMsg "  Assigned by group: $($grp.DisplayName) ($gId)"
                        } catch {
                            Write-InfoMsg "  Assigned by group ID: $gId"
                        }
                    }
                    Write-Host ""
                    continue
                }
            }

            if (Confirm-Action "Remove '$friendlyName' from $($user.UserPrincipalName)?") {
                try {
                    Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($lic.SkuId) -ErrorAction Stop
                    Write-Success "$friendlyName removed."
                } catch {
                    $errMsg = $_.Exception.Message
                    if ($errMsg -match "group-based|inherited|cannot remove") {
                        Write-ErrorMsg "Cannot remove: this license is inherited from a group."
                        Write-InfoMsg "Remove the user from the licensing group instead."
                    } else {
                        Write-ErrorMsg "Failed to remove: $errMsg"
                    }
                }
            }
        }
    }

    Write-Success "License management complete."
    Pause-ForUser
}


# ============ Archive.ps1 ============
# ============================================================
#  Archive.ps1 - Mailbox Archiving Management
# ============================================================

function Start-ArchiveManagement {
    Write-SectionHeader "Mailbox Archiving"

    if (-not (Connect-ForTask "Archive")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    # ---- Identify user ----
    $user = Resolve-UserIdentity -PromptText "Enter user name or email"
    if ($null -eq $user) { Pause-ForUser; return }

    $upn = $user.UserPrincipalName

    # ---- Check current archive status ----
    Write-SectionHeader "Archive Status for $($user.DisplayName)"

    try {
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop

        $archiveEnabled = $mailbox.ArchiveStatus -eq "Active"
        $archiveName    = $mailbox.ArchiveName
        $archiveGuid    = $mailbox.ArchiveGuid

        Write-StatusLine "Archive Enabled" $(if ($archiveEnabled) { "Yes" } else { "No" }) `
            $(if ($archiveEnabled) { "Green" } else { "Red" })

        if ($archiveEnabled) {
            Write-StatusLine "Archive Name" $archiveName "White"
            Write-StatusLine "Archive GUID" $archiveGuid "White"
        }

        # Check retention policy
        $retentionPolicy = $mailbox.RetentionPolicy
        if ($retentionPolicy) {
            Write-StatusLine "Retention Policy" $retentionPolicy "Cyan"
        } else {
            Write-StatusLine "Retention Policy" "(none)" "Gray"
        }
    }
    catch {
        Write-ErrorMsg "Could not retrieve mailbox info: $_"
        Pause-ForUser; return
    }

    Write-Host ""

    if ($archiveEnabled) {
        Write-InfoMsg "Archive is already enabled for this user."

        $changePolicy = Show-Menu -Title "Options" -Options @(
            "Change retention policy",
            "View archive details only"
        ) -BackLabel "Done"

        if ($changePolicy -eq 0) {
            Set-ArchiveRetentionPolicy -UPN $upn
        }
    }
    else {
        # ---- Enable archive ----
        if (Confirm-Action "Enable archive mailbox for $upn and start archiving immediately?") {
            try {
                Enable-Mailbox -Identity $upn -Archive -ErrorAction Stop
                Write-Success "Archive mailbox enabled for $upn."

                # Start managed folder assistant to begin archiving immediately
                Write-InfoMsg "Starting Managed Folder Assistant to initiate archiving..."
                Start-ManagedFolderAssistant -Identity $upn -ErrorAction Stop
                Write-Success "Managed Folder Assistant started. Archiving will begin processing."

            } catch {
                Write-ErrorMsg "Failed to enable archive: $_"
                Pause-ForUser; return
            }

            # Offer to set retention policy
            $setPolicy = Show-Menu -Title "Set a retention policy?" -Options @(
                "Yes, choose a retention policy",
                "No, use default"
            ) -BackLabel "Skip"

            if ($setPolicy -eq 0) {
                Set-ArchiveRetentionPolicy -UPN $upn
            }
        }
    }

    Write-Success "Archive management complete."
    Pause-ForUser
}

function Set-ArchiveRetentionPolicy {
    param([string]$UPN)

    Write-SectionHeader "Available Retention Policies"

    try {
        $policies = Get-RetentionPolicy -ErrorAction Stop
        if ($policies.Count -eq 0) {
            Write-Warn "No retention policies found in the tenant."
            return
        }

        $policyLabels = $policies | ForEach-Object {
            $tags = ($_.RetentionPolicyTagLinks | ForEach-Object { $_.Name }) -join ", "
            if ([string]::IsNullOrWhiteSpace($tags)) { $tags = "(no tags)" }
            "$($_.Name)  [$tags]"
        }

        $sel = Show-Menu -Title "Select a retention policy" -Options $policyLabels -BackLabel "Cancel"
        if ($sel -eq -1) { return }

        $chosenPolicy = $policies[$sel]

        if (Confirm-Action "Apply retention policy '$($chosenPolicy.Name)' to $UPN?") {
            Set-Mailbox -Identity $UPN -RetentionPolicy $chosenPolicy.Name -ErrorAction Stop
            Write-Success "Retention policy '$($chosenPolicy.Name)' applied."

            Write-InfoMsg "Starting Managed Folder Assistant to process immediately..."
            Start-ManagedFolderAssistant -Identity $UPN -ErrorAction Stop
            Write-Success "Managed Folder Assistant started."
        }
    }
    catch {
        Write-ErrorMsg "Retention policy error: $_"
    }
}


# ============ SecurityGroup.ps1 ============
# ============================================================
#  SecurityGroup.ps1 - Security Group Management (MS Graph)
# ============================================================

function Start-SecurityGroupManagement {
    Write-SectionHeader "Security Group Management"
    if (-not (Connect-ForTask "SecurityGroup")) { Pause-ForUser; return }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new security group","Add / remove members",
        "View / edit group properties","Delete a security group"
    ) -BackLabel "Back to Main Menu"

    switch ($action) {
        0 { New-SecurityGroup }
        1 { Edit-SecurityGroupMembers }
        2 { Edit-SecurityGroupProperties }
        3 { Remove-SecurityGroupFlow }
    }
}

function New-SecurityGroup {
    Write-SectionHeader "Create New Security Group"
    $name = Read-UserInput "Group display name"
    if ([string]::IsNullOrWhiteSpace($name)) { Pause-ForUser; return }
    $desc = Read-UserInput "Description (or Enter to skip)"
    $mail = Read-UserInput "Mail nickname (no spaces)"
    if ([string]::IsNullOrWhiteSpace($mail)) { $mail = ($name -replace '[^a-zA-Z0-9]','').ToLower() }
    $me = Show-Menu -Title "Mail-enabled?" -Options @("No (standard)","Yes (mail-enabled)") -BackLabel "Cancel"
    if ($me -eq -1) { return }

    if (Confirm-Action "Create security group '$name'?") {
        try {
            $body = @{ DisplayName = $name; MailEnabled = ($me -eq 1); MailNickname = $mail; SecurityEnabled = $true }
            if ($desc) { $body["Description"] = $desc }
            $g = New-MgGroup -BodyParameter $body -ErrorAction Stop
            Write-Success "Created. Id: $($g.Id)"
            $add = Read-UserInput "Add members now? (y/n)"
            if ($add -match '^[Yy]') { Add-MembersLoop -GroupId $g.Id -GroupName $name }
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Edit-SecurityGroupMembers {
    $group = Find-SecurityGroup; if ($null -eq $group) { Pause-ForUser; return }
    Show-GroupMembers -GroupId $group.Id -GroupName $group.DisplayName
    $action = Show-Menu -Title "Action" -Options @("Add member(s)","Remove member(s)") -BackLabel "Done"
    if ($action -eq 0) { Add-MembersLoop -GroupId $group.Id -GroupName $group.DisplayName }
    elseif ($action -eq 1) {
        try {
            $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop)
            if ($members.Count -eq 0) { Write-InfoMsg "No members."; Pause-ForUser; return }
            $labels = $members | ForEach-Object { "$($_.AdditionalProperties['displayName']) ($($_.AdditionalProperties['userPrincipalName']))" }
            $selected = Show-MultiSelect -Title "Select member(s) to remove" -Options $labels
            foreach ($idx in $selected) {
                $m = $members[$idx]
                if (Confirm-Action "Remove '$($m.AdditionalProperties['displayName'])'?") {
                    try { Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $m.Id -ErrorAction Stop; Write-Success "Removed." }
                    catch { Write-ErrorMsg "Failed: $_" }
                }
            }
        } catch { Write-ErrorMsg "Error: $_" }
    }
    Pause-ForUser
}

function Edit-SecurityGroupProperties {
    $group = Find-SecurityGroup; if ($null -eq $group) { Pause-ForUser; return }
    Write-StatusLine "Display Name" $group.DisplayName "White"
    Write-StatusLine "Description" $(if ($group.Description) { $group.Description } else { "(none)" }) "White"
    Write-StatusLine "Mail Nickname" $group.MailNickname "White"
    Write-StatusLine "Mail Enabled" "$($group.MailEnabled)" "White"
    Write-StatusLine "Object ID" $group.Id "Gray"
    try { $mc = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop).Count; Write-StatusLine "Members" "$mc" "Cyan" } catch {}

    $ec = Show-Menu -Title "Edit" -Options @("Change name","Change description","Change mail nickname") -BackLabel "Done"
    switch ($ec) {
        0 { $v = Read-UserInput "New name"; if ($v -and (Confirm-Action "Rename to '$v'?")) { try { Update-MgGroup -GroupId $group.Id -DisplayName $v; Write-Success "Updated." } catch { Write-ErrorMsg "$_" } } }
        1 { $v = Read-UserInput "New description (or 'clear')"; $sv = if ($v -eq 'clear') { "" } else { $v }; if (Confirm-Action "Update description?") { try { Update-MgGroup -GroupId $group.Id -Description $sv; Write-Success "Updated." } catch { Write-ErrorMsg "$_" } } }
        2 { $v = Read-UserInput "New mail nickname"; if ($v -and (Confirm-Action "Change to '$v'?")) { try { Update-MgGroup -GroupId $group.Id -MailNickname $v; Write-Success "Updated." } catch { Write-ErrorMsg "$_" } } }
    }
    Pause-ForUser
}

function Remove-SecurityGroupFlow {
    $group = Find-SecurityGroup; if ($null -eq $group) { Pause-ForUser; return }
    Write-StatusLine "Group" $group.DisplayName "White"
    Write-Warn "This is irreversible!"
    if (Confirm-Action "DELETE '$($group.DisplayName)'?") {
        $check = Read-UserInput "Type the group name to confirm"
        if ($check -eq $group.DisplayName) {
            try { Remove-MgGroup -GroupId $group.Id -ErrorAction Stop; Write-Success "Deleted." }
            catch { Write-ErrorMsg "Failed: $_" }
        } else { Write-Warn "Name mismatch. Cancelled." }
    }
    Pause-ForUser
}

function Find-SecurityGroup {
    $s = Read-UserInput "Search security group by name"
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try {
        $groups = @(Get-MgGroup -Search "displayName:$s" -ConsistencyLevel eventual -ErrorAction Stop | Where-Object { $_.SecurityEnabled })
        if ($groups.Count -eq 0) { Write-ErrorMsg "None found."; return $null }
        if ($groups.Count -eq 1) { Write-Success "Found: $($groups[0].DisplayName)"; return $groups[0] }
        $labels = $groups | ForEach-Object { $_.DisplayName }
        $sel = Show-Menu -Title "Select" -Options $labels -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }; return $groups[$sel]
    } catch { Write-ErrorMsg "Search error: $_"; return $null }
}

function Show-GroupMembers {
    param([string]$GroupId, [string]$GroupName)
    try {
        $members = @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop)
        Write-InfoMsg "Members of '$GroupName' ($($members.Count)):"
        if ($members.Count -eq 0) { Write-InfoMsg "  (none)" }
        else { $members | ForEach-Object { Write-Host "    - $($_.AdditionalProperties['displayName']) ($($_.AdditionalProperties['userPrincipalName']))" -ForegroundColor White } }
    } catch { Write-Warn "Could not read members: $_" }
}

function Add-MembersLoop {
    param([string]$GroupId, [string]$GroupName)
    while ($true) {
        $ui = Read-UserInput "User name or email to add (or 'done')"
        if ($ui -match '^done$') { break }
        try {
            $tu = if ($ui -match '@') { Get-MgUser -UserId $ui -ErrorAction Stop } else {
                $f = @(Get-MgUser -Search "displayName:$ui" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; continue }
                if ($f.Count -eq 1) { $f[0] } else {
                    $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                    if ($sel -eq -1) { continue }; $f[$sel]
                }
            }
            if (Confirm-Action "Add '$($tu.DisplayName)' to '$GroupName'?") {
                New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $tu.Id -ErrorAction Stop; Write-Success "Added."
            }
        } catch { if ($_.Exception.Message -match "already exist") { Write-Warn "Already a member." } else { Write-ErrorMsg "Failed: $_" } }
    }
}


# ============ DistributionList.ps1 ============
# ============================================================
#  DistributionList.ps1 - DL Management (MS Graph + EXO)
# ============================================================

function Start-DistributionListManagement {
    Write-SectionHeader "Distribution List Management"
    if (-not (Connect-ForTask "DistributionList")) { Pause-ForUser; return }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new distribution list","Add / remove members",
        "View / edit DL properties","Delete a distribution list"
    ) -BackLabel "Back to Main Menu"

    switch ($action) {
        0 { New-DistributionListFlow }
        1 { Edit-DistributionListMembers }
        2 { Edit-DistributionListProperties }
        3 { Remove-DistributionListFlow }
    }
}

function New-DistributionListFlow {
    Write-SectionHeader "Create New Distribution List"
    $name = Read-UserInput "Display name"; if ([string]::IsNullOrWhiteSpace($name)) { Pause-ForUser; return }
    $alias = Read-UserInput "Email alias (e.g. sales-team)"; if ([string]::IsNullOrWhiteSpace($alias)) { $alias = ($name -replace '[^a-zA-Z0-9-]','').ToLower() }
    $smtp = Read-UserInput "Full primary email (e.g. sales@contoso.com)"; if ([string]::IsNullOrWhiteSpace($smtp)) { Pause-ForUser; return }
    $owner = Read-UserInput "Owner email (or Enter to skip)"
    $ra = Show-Menu -Title "Who can send?" -Options @("Anyone","Internal only") -BackLabel "Cancel"; if ($ra -eq -1) { return }

    if (Confirm-Action "Create DL '$name' ($smtp)?") {
        try {
            $p = @{ Name = $name; DisplayName = $name; Alias = $alias; PrimarySmtpAddress = $smtp; Type = "Distribution"; RequireSenderAuthenticationEnabled = ($ra -eq 1) }
            if ($owner) { $p["ManagedBy"] = $owner }
            New-DistributionGroup @p -ErrorAction Stop | Out-Null; Write-Success "Created."
            $add = Read-UserInput "Add members now? (y/n)"; if ($add -match '^[Yy]') { Add-DLMembersLoop -Id $smtp -Name $name }
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Edit-DistributionListMembers {
    $user = Resolve-UserIdentity; if ($null -eq $user) { Pause-ForUser; return }
    $upn = $user.UserPrincipalName
    $action = Show-Menu -Title "Action" -Options @("Add to DL(s)","Remove from DL(s)") -BackLabel "Cancel"; if ($action -eq -1) { return }

    if ($action -eq 0) {
        $dl = Find-DistributionList; if ($null -eq $dl) { Pause-ForUser; return }
        if (Confirm-Action "Add to '$($dl.DisplayName)'?") {
            try { Add-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -Member $upn -ErrorAction Stop; Write-Success "Added." }
            catch { if ($_.Exception.Message -match "already") { Write-Warn "Already a member." } else { Write-ErrorMsg "$_" } }
        }
        $pc = Show-Menu -Title "Send permissions?" -Options @("Send As","Send on Behalf","Both","None") -BackLabel "Skip"
        if ($pc -ne -1 -and $pc -ne 3) {
            if ($pc -eq 0 -or $pc -eq 2) { if (Confirm-Action "Grant Send As?") { try { Add-RecipientPermission -Identity $dl.PrimarySmtpAddress -Trustee $upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop; Write-Success "Granted." } catch { Write-ErrorMsg "$_" } } }
            if ($pc -eq 1 -or $pc -eq 2) { if (Confirm-Action "Grant Send on Behalf?") { try { Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -GrantSendOnBehalfTo @{Add=$upn} -ErrorAction Stop; Write-Success "Granted." } catch { Write-ErrorMsg "$_" } } }
        }
    } else {
        Write-InfoMsg "Finding DL memberships..."
        try {
            $allDLs = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
            $memberOf = @(); foreach ($dl in $allDLs) {
                $ms = Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ErrorAction SilentlyContinue
                if ($ms | Where-Object { $_.PrimarySmtpAddress -eq $upn }) { $memberOf += $dl }
            }
            if ($memberOf.Count -eq 0) { Write-InfoMsg "Not in any DLs."; Pause-ForUser; return }
            $labels = $memberOf | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }
            $sel = Show-MultiSelect -Title "Remove from" -Options $labels
            foreach ($idx in $sel) { $dl = $memberOf[$idx]; if (Confirm-Action "Remove from '$($dl.DisplayName)'?") { try { Remove-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -Member $upn -Confirm:$false -ErrorAction Stop; Write-Success "Removed." } catch { Write-ErrorMsg "$_" } } }
        } catch { Write-ErrorMsg "$_" }
    }
    Pause-ForUser
}

function Edit-DistributionListProperties {
    $dl = Find-DistributionList; if ($null -eq $dl) { Pause-ForUser; return }
    try { $dl = Get-DistributionGroup -Identity $dl.PrimarySmtpAddress -ErrorAction Stop } catch {}
    Write-StatusLine "Name" $dl.DisplayName "White"; Write-StatusLine "Email" $dl.PrimarySmtpAddress "White"
    Write-StatusLine "Description" $(if ($dl.Description) { $dl.Description } else { "(none)" }) "White"
    Write-StatusLine "Managed By" ($dl.ManagedBy -join "; ") "White"
    Write-StatusLine "Auth Required" "$($dl.RequireSenderAuthenticationEnabled)" "White"
    Write-StatusLine "Hidden" "$($dl.HiddenFromAddressListsEnabled)" "White"

    $ec = Show-Menu -Title "Edit" -Options @("Change name","Change description","Change owner","Toggle sender auth","Toggle hidden") -BackLabel "Done"
    switch ($ec) {
        0 { $v = Read-UserInput "New name"; if ($v -and (Confirm-Action "Rename?")) { try { Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -DisplayName $v; Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
        1 { $v = Read-UserInput "New description (or 'clear')"; if (Confirm-Action "Update?") { try { Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -Description $(if ($v -eq 'clear') {""} else {$v}); Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
        2 { $v = Read-UserInput "New owner email"; if ($v -and (Confirm-Action "Set owner?")) { try { Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -ManagedBy $v; Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
        3 { $nv = -not $dl.RequireSenderAuthenticationEnabled; if (Confirm-Action "Set to $nv?") { try { Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -RequireSenderAuthenticationEnabled $nv; Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
        4 { $nv = -not $dl.HiddenFromAddressListsEnabled; if (Confirm-Action "Set hidden to $nv?") { try { Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -HiddenFromAddressListsEnabled $nv; Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
    }
    Pause-ForUser
}

function Remove-DistributionListFlow {
    $dl = Find-DistributionList; if ($null -eq $dl) { Pause-ForUser; return }
    Write-Warn "Irreversible!"
    if (Confirm-Action "DELETE '$($dl.DisplayName)'?") {
        $check = Read-UserInput "Type the DL email to confirm"
        if ($check -eq $dl.PrimarySmtpAddress) { try { Remove-DistributionGroup -Identity $dl.PrimarySmtpAddress -Confirm:$false; Write-Success "Deleted." } catch { Write-ErrorMsg "$_" } }
        else { Write-Warn "Mismatch. Cancelled." }
    }
    Pause-ForUser
}

function Find-DistributionList {
    $sm = Show-Menu -Title "Find DL by" -Options @("Name","Email") -BackLabel "Cancel"; if ($sm -eq -1) { return $null }
    $si = Read-UserInput $(if ($sm -eq 0) { "DL name" } else { "DL email" }); if ([string]::IsNullOrWhiteSpace($si)) { return $null }
    try {
        $dls = @(if ($sm -eq 0) { Get-DistributionGroup -Filter "DisplayName -like '*$si*'" -ResultSize 50 } else { Get-DistributionGroup -Filter "PrimarySmtpAddress -like '*$si*'" -ResultSize 50 })
        if ($dls.Count -eq 0) { Write-ErrorMsg "None found."; return $null }
        if ($dls.Count -eq 1) { return $dls[0] }
        $sel = Show-Menu -Title "Select" -Options ($dls | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }) -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }; return $dls[$sel]
    } catch { Write-ErrorMsg "$_"; return $null }
}

function Add-DLMembersLoop {
    param([string]$Id, [string]$Name)
    while ($true) {
        $ui = Read-UserInput "User to add (or 'done')"; if ($ui -match '^done$') { break }
        try {
            $tu = if ($ui -match '@') { Get-MgUser -UserId $ui -ErrorAction Stop } else {
                $f = @(Get-MgUser -Search "displayName:$ui" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; continue }; if ($f.Count -eq 1) { $f[0] } else {
                    $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                    if ($sel -eq -1) { continue }; $f[$sel]
                }
            }
            if (Confirm-Action "Add '$($tu.DisplayName)'?") { Add-DistributionGroupMember -Identity $Id -Member $tu.UserPrincipalName -ErrorAction Stop; Write-Success "Added." }
        } catch { if ($_.Exception.Message -match "already") { Write-Warn "Already a member." } else { Write-ErrorMsg "$_" } }
    }
}


# ============ SharedMailbox.ps1 ============
# ============================================================
#  SharedMailbox.ps1 - Shared Mailbox Management (MS Graph + EXO)
# ============================================================

function Start-SharedMailboxManagement {
    Write-SectionHeader "Shared Mailbox Management"
    if (-not (Connect-ForTask "SharedMailbox")) { Pause-ForUser; return }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new shared mailbox","Add / remove user access",
        "View / edit mailbox properties","Delete a shared mailbox"
    ) -BackLabel "Back to Main Menu"

    switch ($action) { 0 { New-SharedMailboxFlow } 1 { Edit-SharedMailboxAccess } 2 { Edit-SharedMailboxProperties } 3 { Remove-SharedMailboxFlow } }
}

function New-SharedMailboxFlow {
    Write-SectionHeader "Create New Shared Mailbox"
    $name = Read-UserInput "Display name"; if ([string]::IsNullOrWhiteSpace($name)) { Pause-ForUser; return }
    $email = Read-UserInput "Email address"; if ([string]::IsNullOrWhiteSpace($email)) { Pause-ForUser; return }
    $alias = Read-UserInput "Alias (or Enter to auto)"; if ([string]::IsNullOrWhiteSpace($alias)) { $alias = ($email -split '@')[0] }
    if (Confirm-Action "Create shared mailbox '$name' ($email)?") {
        try {
            New-Mailbox -Name $name -DisplayName $name -Alias $alias -PrimarySmtpAddress $email -Shared -ErrorAction Stop | Out-Null
            Write-Success "Created."
            $add = Read-UserInput "Grant access now? (y/n)"; if ($add -match '^[Yy]') { Add-SharedMailboxAccessLoop -Id $email -Name $name }
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Edit-SharedMailboxAccess {
    $box = Find-SharedMailbox; if ($null -eq $box) { Pause-ForUser; return }
    Show-MailboxPermissions -Id $box.PrimarySmtpAddress -Name $box.DisplayName
    $action = Show-Menu -Title "Action" -Options @("Grant access","Remove access") -BackLabel "Done"
    if ($action -eq 0) { Add-SharedMailboxAccessLoop -Id $box.PrimarySmtpAddress -Name $box.DisplayName }
    elseif ($action -eq 1) {
        try {
            $perms = @(Get-MailboxPermission -Identity $box.PrimarySmtpAddress -ErrorAction Stop | Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-*" -and -not $_.IsInherited })
            if ($perms.Count -eq 0) { Write-InfoMsg "No custom permissions."; Pause-ForUser; return }
            $labels = $perms | ForEach-Object { "$($_.User)  ($($_.AccessRights -join ', '))" }
            $sel = Show-MultiSelect -Title "Remove" -Options $labels
            foreach ($idx in $sel) {
                $p = $perms[$idx]
                if (Confirm-Action "Remove all permissions for '$($p.User)'?") {
                    try { Remove-MailboxPermission -Identity $box.PrimarySmtpAddress -User $p.User -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop; Write-Success "Full Access removed." } catch { Write-ErrorMsg "$_" }
                    try { Remove-RecipientPermission -Identity $box.PrimarySmtpAddress -Trustee $p.User -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue; Write-Success "Send As removed." } catch {}
                    try { Set-Mailbox -Identity $box.PrimarySmtpAddress -GrantSendOnBehalfTo @{Remove=$p.User} -ErrorAction SilentlyContinue; Write-Success "Send on Behalf removed." } catch {}
                }
            }
        } catch { Write-ErrorMsg "$_" }
    }
    Pause-ForUser
}

function Edit-SharedMailboxProperties {
    $box = Find-SharedMailbox; if ($null -eq $box) { Pause-ForUser; return }
    try { $box = Get-Mailbox -Identity $box.PrimarySmtpAddress -ErrorAction Stop } catch {}
    Write-StatusLine "Name" $box.DisplayName "White"; Write-StatusLine "Email" $box.PrimarySmtpAddress "White"
    Write-StatusLine "Alias" $box.Alias "White"; Write-StatusLine "Hidden" "$($box.HiddenFromAddressListsEnabled)" "White"
    Write-StatusLine "Forwarding" $(if ($box.ForwardingSmtpAddress) { $box.ForwardingSmtpAddress } else { "(none)" }) "White"

    $ec = Show-Menu -Title "Edit" -Options @("Change name","Add email alias","Remove email alias","Set forwarding","Remove forwarding","Toggle hidden","Set auto-reply") -BackLabel "Done"
    switch ($ec) {
        0 { $v = Read-UserInput "New name"; if ($v -and (Confirm-Action "Rename?")) { try { Set-Mailbox -Identity $box.PrimarySmtpAddress -DisplayName $v; Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
        1 { $v = Read-UserInput "New alias email"; if ($v -and (Confirm-Action "Add '$v'?")) { try { Set-Mailbox -Identity $box.PrimarySmtpAddress -EmailAddresses @{Add="smtp:$v"}; Write-Success "Added." } catch { Write-ErrorMsg "$_" } } }
        2 { $a = @($box.EmailAddresses | Where-Object { $_ -like "smtp:*" }); if ($a.Count -eq 0) { Write-InfoMsg "No aliases." } else { $al = $a | ForEach-Object { $_ -replace '^smtp:','' }; $s = Show-MultiSelect -Title "Remove" -Options $al; foreach ($i in $s) { if (Confirm-Action "Remove '$($al[$i])'?") { try { Set-Mailbox -Identity $box.PrimarySmtpAddress -EmailAddresses @{Remove=$a[$i]}; Write-Success "Removed." } catch { Write-ErrorMsg "$_" } } } } }
        3 { $v = Read-UserInput "Forward to"; $kc = Show-Menu -Title "Keep copy?" -Options @("Yes","No") -BackLabel "Cancel"; if ($kc -ne -1 -and $v) { if (Confirm-Action "Set forwarding?") { try { Set-Mailbox -Identity $box.PrimarySmtpAddress -ForwardingSmtpAddress "smtp:$v" -DeliverToMailboxAndForward ($kc -eq 0); Write-Success "Done." } catch { Write-ErrorMsg "$_" } } } }
        4 { if (Confirm-Action "Remove forwarding?") { try { Set-Mailbox -Identity $box.PrimarySmtpAddress -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false; Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
        5 { $nv = -not $box.HiddenFromAddressListsEnabled; if (Confirm-Action "Set hidden to $nv?") { try { Set-Mailbox -Identity $box.PrimarySmtpAddress -HiddenFromAddressListsEnabled $nv; Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
        6 { $im = Read-UserInput "Internal message"; $em = Read-UserInput "External (Enter for same)"; if ([string]::IsNullOrWhiteSpace($em)) { $em = $im }; if ($im -and (Confirm-Action "Set auto-reply?")) { try { Set-MailboxAutoReplyConfiguration -Identity $box.PrimarySmtpAddress -AutoReplyState Enabled -InternalMessage $im -ExternalMessage $em; Write-Success "Done." } catch { Write-ErrorMsg "$_" } } }
    }
    Pause-ForUser
}

function Remove-SharedMailboxFlow {
    $box = Find-SharedMailbox; if ($null -eq $box) { Pause-ForUser; return }
    Write-Warn "This permanently deletes the mailbox and ALL contents!"
    if (Confirm-Action "DELETE '$($box.DisplayName)'?") {
        $check = Read-UserInput "Type email to confirm"; if ($check -eq $box.PrimarySmtpAddress) { try { Remove-Mailbox -Identity $box.PrimarySmtpAddress -Confirm:$false; Write-Success "Deleted." } catch { Write-ErrorMsg "$_" } } else { Write-Warn "Mismatch." }
    }
    Pause-ForUser
}

function Find-SharedMailbox {
    $sm = Show-Menu -Title "Find by" -Options @("Name","Email") -BackLabel "Cancel"; if ($sm -eq -1) { return $null }
    $si = Read-UserInput $(if ($sm -eq 0) { "Mailbox name" } else { "Mailbox email" }); if ([string]::IsNullOrWhiteSpace($si)) { return $null }
    try {
        $boxes = @(if ($sm -eq 0) { Get-Mailbox -RecipientTypeDetails SharedMailbox -Filter "DisplayName -like '*$si*'" -ResultSize 50 } else { Get-Mailbox -RecipientTypeDetails SharedMailbox -Filter "PrimarySmtpAddress -like '*$si*'" -ResultSize 50 })
        if ($boxes.Count -eq 0) { Write-ErrorMsg "None found."; return $null }
        if ($boxes.Count -eq 1) { return $boxes[0] }
        $sel = Show-Menu -Title "Select" -Options ($boxes | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }) -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }; return $boxes[$sel]
    } catch { Write-ErrorMsg "$_"; return $null }
}

function Show-MailboxPermissions { param([string]$Id, [string]$Name)
    Write-InfoMsg "Permissions on '$Name':"
    try { $p = Get-MailboxPermission -Identity $Id -ErrorAction Stop | Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-*" -and -not $_.IsInherited }
        if ($p.Count -eq 0) { Write-InfoMsg "  (none)" } else { $p | ForEach-Object { Write-Host "    - $($_.User) [$($_.AccessRights -join ', ')]" -ForegroundColor White } }
    } catch { Write-Warn "$_" }
}

function Add-SharedMailboxAccessLoop { param([string]$Id, [string]$Name)
    while ($true) {
        $ui = Read-UserInput "User to grant access (or 'done')"; if ($ui -match '^done$') { break }
        try {
            $tu = if ($ui -match '@') { Get-MgUser -UserId $ui -ErrorAction Stop } else {
                $f = @(Get-MgUser -Search "displayName:$ui" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; continue }; if ($f.Count -eq 1) { $f[0] } else {
                    $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"; if ($sel -eq -1) { continue }; $f[$sel] } }
            $upn = $tu.UserPrincipalName
            if (Confirm-Action "Grant Full Access to $($tu.DisplayName)?") { try { Add-MailboxPermission -Identity $Id -User $upn -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop; Write-Success "Granted." } catch { Write-ErrorMsg "$_" } }
            $pc = Show-Menu -Title "Send permissions?" -Options @("Send As","Send on Behalf","Both","None") -BackLabel "Skip"
            if ($pc -ne -1 -and $pc -ne 3) {
                if ($pc -eq 0 -or $pc -eq 2) { if (Confirm-Action "Send As?") { try { Add-RecipientPermission -Identity $Id -Trustee $upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop; Write-Success "Granted." } catch { Write-ErrorMsg "$_" } } }
                if ($pc -eq 1 -or $pc -eq 2) { if (Confirm-Action "Send on Behalf?") { try { Set-Mailbox -Identity $Id -GrantSendOnBehalfTo @{Add=$upn} -ErrorAction Stop; Write-Success "Granted." } catch { Write-ErrorMsg "$_" } } }
            }
        } catch { Write-ErrorMsg "$_" }
    }
}


# ============ CalendarAccess.ps1 ============
# ============================================================
#  CalendarAccess.ps1 - Calendar Permission Management (MS Graph + EXO)
# ============================================================

function Start-CalendarAccessManagement {
    Write-SectionHeader "Calendar Access Management"
    if (-not (Connect-ForTask "CalendarAccess")) { Pause-ForUser; return }

    Write-InfoMsg "Identify the calendar OWNER."
    $owner = Resolve-UserIdentity -PromptText "Enter calendar owner name or email"
    if ($null -eq $owner) { Pause-ForUser; return }
    $calId = "$($owner.UserPrincipalName):\Calendar"

    Write-SectionHeader "Current Permissions"
    try { $perms = Get-MailboxFolderPermission -Identity $calId -ErrorAction Stop
        foreach ($p in $perms) { $pu = if ($p.User.DisplayName) { $p.User.DisplayName } else { $p.User.ToString() }; Write-Host "    $pu  -  $($p.AccessRights -join ', ')" -ForegroundColor White }
    } catch { Write-ErrorMsg "Could not read permissions: $_"; Pause-ForUser; return }

    $action = Show-Menu -Title "Action" -Options @("Add access","Remove access") -BackLabel "Cancel"
    if ($action -eq -1) { return }

    if ($action -eq 0) {
        $si = Read-UserInput "User to grant access (name or email)"
        try {
            $tu = if ($si -match '@') { Get-MgUser -UserId $si -ErrorAction Stop } else {
                $f = @(Get-MgUser -Search "displayName:$si" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; Pause-ForUser; return }
                if ($f.Count -eq 1) { $f[0] } else {
                    $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                    if ($sel -eq -1) { Pause-ForUser; return }; $f[$sel]
                }
            }
        } catch { Write-ErrorMsg "$_"; Pause-ForUser; return }

        $pl = Show-Menu -Title "Access level" -Options @("Reviewer (read-only)","Editor (full edit)","Author (create + edit own)","Contributor (create only)") -BackLabel "Cancel"
        if ($pl -eq -1) { Pause-ForUser; return }
        $accessMap = @("Reviewer","Editor","Author","Contributor"); $ar = $accessMap[$pl]

        if (Confirm-Action "Grant $ar to $($tu.DisplayName) on $($owner.DisplayName) calendar?") {
            try {
                try { Add-MailboxFolderPermission -Identity $calId -User $tu.UserPrincipalName -AccessRights $ar -ErrorAction Stop; Write-Success "Granted: $ar" }
                catch { if ($_.Exception.Message -match "already exists") { Set-MailboxFolderPermission -Identity $calId -User $tu.UserPrincipalName -AccessRights $ar -ErrorAction Stop; Write-Success "Updated to: $ar" } else { throw $_ } }
            } catch { Write-ErrorMsg "Failed: $_" }
        }
    } else {
        $removable = @($perms | Where-Object { $_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" -and $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" })
        if ($removable.Count -eq 0) { Write-InfoMsg "No custom permissions."; Pause-ForUser; return }
        $labels = $removable | ForEach-Object { $pu = if ($_.User.DisplayName) { $_.User.DisplayName } else { $_.User.ToString() }; "$pu ($($_.AccessRights -join ', '))" }
        $sel = Show-MultiSelect -Title "Remove" -Options $labels
        foreach ($idx in $sel) {
            $p = $removable[$idx]; $pu = if ($p.User.DisplayName) { $p.User.DisplayName } else { $p.User.ToString() }
            if (Confirm-Action "Remove access for '$pu'?") { try { Remove-MailboxFolderPermission -Identity $calId -User $pu -Confirm:$false -ErrorAction Stop; Write-Success "Removed." } catch { Write-ErrorMsg "$_" } }
        }
    }
    Pause-ForUser
}


# ============ UserProfile.ps1 ============
# ============================================================
#  UserProfile.ps1 - View / Edit User Profile (Microsoft Graph)
# ============================================================

function Start-UserProfileManagement {
    Write-SectionHeader "User Profile Management"
    if (-not (Connect-ForTask "UserProfile")) { Pause-ForUser; return }

    $user = Resolve-UserIdentity
    if ($null -eq $user) { Pause-ForUser; return }

    # Editable fields: label -> Graph property
    $editableFields = [ordered]@{
        "Display Name"    = "DisplayName"
        "First Name"      = "GivenName"
        "Last Name"       = "Surname"
        "Job Title"       = "JobTitle"
        "Department"      = "Department"
        "Company Name"    = "CompanyName"
        "Office Location" = "OfficeLocation"
        "Street Address"  = "StreetAddress"
        "City"            = "City"
        "State"           = "State"
        "Postal Code"     = "PostalCode"
        "Country"         = "Country"
        "Usage Location"  = "UsageLocation"
        "Mobile Phone"    = "MobilePhone"
        "Mail Nickname"   = "MailNickname"
    }

    # Read-only fields
    $readOnlyFields = [ordered]@{
        "UPN"             = "UserPrincipalName"
        "Mail"            = "Mail"
        "Business Phones" = "BusinessPhones"
        "Account Enabled" = "AccountEnabled"
        "User Type"       = "UserType"
        "Object ID"       = "Id"
    }

    $keepGoing = $true
    while ($keepGoing) {
        # Refresh profile
        try {
            $profile = Get-MgUser -UserId $user.Id -Property ($editableFields.Values + $readOnlyFields.Values + @("Id")) -ErrorAction Stop
        } catch { Write-ErrorMsg "Could not load profile: $_"; Pause-ForUser; return }

        # Manager
        $mgrDisplay = "(not set)"
        try {
            $mgr = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
            if ($mgr) { $mgrDisplay = "$($mgr.AdditionalProperties['displayName']) ($($mgr.AdditionalProperties['userPrincipalName']))" }
        } catch {}

        # Display
        $b = $script:Box
        Write-Host ""
        Write-Host ("  " + $b.DTL + [string]::new($b.DH, 2) + " Profile: $($profile.DisplayName) " + [string]::new($b.DH, 20) + $b.DTR) -ForegroundColor $script:Colors.Title
        Write-Host ""

        $idx = 1
        foreach ($label in $editableFields.Keys) {
            $prop = $editableFields[$label]
            $val = $profile.$prop
            $valStr = if ($null -eq $val -or ([string]::IsNullOrWhiteSpace("$val"))) { "(not set)" } else { "$val" }
            Write-Host "  [" -NoNewline -ForegroundColor $script:Colors.Accent
            Write-Host ("{0,2}" -f $idx) -NoNewline -ForegroundColor $script:Colors.Highlight
            Write-Host "]" -NoNewline -ForegroundColor $script:Colors.Accent
            Write-Host (" {0,-18}" -f $label) -NoNewline -ForegroundColor $script:Colors.Info
            Write-Host ": " -NoNewline
            Write-Host $valStr -ForegroundColor $(if ($valStr -eq "(not set)") { "DarkGray" } else { "White" })
            $idx++
        }

        # Manager row (editable via separate menu)
        Write-Host "  [" -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host " M" -NoNewline -ForegroundColor $script:Colors.Highlight
        Write-Host "]" -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host (" {0,-18}" -f "Manager") -NoNewline -ForegroundColor $script:Colors.Info
        Write-Host ": " -NoNewline
        Write-Host $mgrDisplay -ForegroundColor $(if ($mgrDisplay -eq "(not set)") { "DarkGray" } else { "White" })

        Write-Host ""
        # Read-only
        foreach ($label in $readOnlyFields.Keys) {
            $prop = $readOnlyFields[$label]
            $val = $profile.$prop
            if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) { $valStr = ($val -join "; ") }
            else { $valStr = if ($null -eq $val -or ([string]::IsNullOrWhiteSpace("$val"))) { "(not set)" } else { "$val" } }
            Write-Host "     " -NoNewline
            Write-Host (" {0,-18}" -f $label) -NoNewline -ForegroundColor $script:Colors.Info
            Write-Host ": " -NoNewline
            Write-Host $valStr -ForegroundColor "Gray"
        }

        Write-Host ""
        Write-Host ("  " + $b.DBL + [string]::new($b.DH, 60) + $b.DBR) -ForegroundColor $script:Colors.Title
        Write-InfoMsg "Numbered fields are editable. Enter number, 'M' for manager, or 'done'."

        $input = Read-UserInput "Action"

        if ($input -match '^done$') { $keepGoing = $false; continue }

        if ($input -match '^[Mm]$') {
            # Manager management
            Write-StatusLine "Current Manager" $mgrDisplay "White"
            $ma = Show-Menu -Title "Manager" -Options @("Set / change","Remove") -BackLabel "Cancel"
            if ($ma -eq 0) {
                $mi = Read-UserInput "New manager name or email"
                if ($mi) {
                    try {
                        $nm = if ($mi -match '@') { Get-MgUser -UserId $mi -ErrorAction Stop } else {
                            $f = @(Get-MgUser -Search "displayName:$mi" -ConsistencyLevel eventual -ErrorAction Stop)
                            if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; Pause-ForUser; continue }
                            if ($f.Count -eq 1) { $f[0] } else {
                                $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                                if ($sel -eq -1) { Pause-ForUser; continue }; $f[$sel]
                            }
                        }
                        if (Confirm-Action "Set manager to $($nm.DisplayName)?") {
                            Set-MgUserManagerByRef -UserId $user.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($nm.Id)" } -ErrorAction Stop
                            Write-Success "Manager set."
                        }
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            } elseif ($ma -eq 1) {
                if (Confirm-Action "Remove manager?") {
                    try { Remove-MgUserManagerByRef -UserId $user.Id -ErrorAction Stop; Write-Success "Removed." }
                    catch { Write-ErrorMsg "Failed: $_" }
                }
            }
            Pause-ForUser; continue
        }

        if ($input -match '^\d+$') {
            $fi = [int]$input
            $editKeys = @($editableFields.Keys)
            if ($fi -ge 1 -and $fi -le $editKeys.Count) {
                $label = $editKeys[$fi - 1]
                $prop = $editableFields[$label]
                $curVal = $profile.$prop
                Write-StatusLine "Field" $label "Cyan"
                Write-StatusLine "Current" $(if ($curVal) { "$curVal" } else { "(not set)" }) "White"
                $newVal = Read-UserInput "New value (or 'clear')"
                if ($newVal -eq 'clear') { $newVal = "" }
                if (Confirm-Action "Update '$label'?" "Old: $(if ($curVal) { $curVal } else { '(empty)' })`nNew: $(if ($newVal) { $newVal } else { '(cleared)' })") {
                    try {
                        $body = @{}; $body[$prop] = if ([string]::IsNullOrWhiteSpace($newVal)) { $null } else { $newVal }
                        Update-MgUser -UserId $user.Id -BodyParameter $body -ErrorAction Stop
                        Write-Success "'$label' updated."
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            } else { Write-ErrorMsg "Invalid number." }
            Pause-ForUser; continue
        }

        Write-ErrorMsg "Enter a field number, 'M', or 'done'."
        Pause-ForUser
    }
    Write-Success "Profile management complete."
}


# ============ Main.ps1 ============
# ============================================================
#  Main.ps1 - M365 Administration Tool - Entry Point
# ============================================================
#  Usage:  powershell -ExecutionPolicy Bypass -File Main.ps1
# ============================================================


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
            9 {
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

