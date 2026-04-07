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
        Handles 403 errors by offering to reconnect Graph with proper consent.
        Returns a user object or $null.
    #>
    param([string]$PromptText = "Enter user name or email")

    $searchInput = Read-UserInput $PromptText
    if ([string]::IsNullOrWhiteSpace($searchInput)) { return $null }

    $userProps = "Id,DisplayName,UserPrincipalName,JobTitle,Department,GivenName,Surname,CompanyName,OfficeLocation,StreetAddress,City,State,PostalCode,Country,UsageLocation,BusinessPhones,MobilePhone,Mail,MailNickname,AccountEnabled,UserType"

    # Allow up to 2 attempts (first try + retry after reconnect)
    for ($attempt = 1; $attempt -le 2; $attempt++) {

        Write-InfoMsg "Searching for user..."

        try {
            if ($searchInput -match '@') {
                $user = Get-MgUser -UserId $searchInput -Property $userProps -ErrorAction Stop
            } else {
                $users = @(Get-MgUser -Search "displayName:$searchInput" -ConsistencyLevel eventual -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department" -ErrorAction Stop)
                if ($users.Count -eq 0) {
                    $users = @(Get-MgUser -Filter "startsWith(displayName,'$searchInput')" -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department" -ErrorAction Stop)
                }
                if ($users.Count -eq 0) {
                    Write-ErrorMsg "No users found matching '$searchInput'."
                    return $null
                }
                if ($users.Count -eq 1) {
                    $user = Get-MgUser -UserId $users[0].Id -Property $userProps -ErrorAction Stop
                } else {
                    Write-Warn "Multiple users found:"
                    $names = $users | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
                    $sel = Show-Menu -Title "Select User" -Options $names -BackLabel "Cancel"
                    if ($sel -eq -1) { return $null }
                    $user = Get-MgUser -UserId $users[$sel].Id -Property $userProps -ErrorAction Stop
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
            $errMsg = "$_"

            # Detect 403 / insufficient privileges
            if ($errMsg -match "403|Forbidden|Authorization_RequestDenied|Insufficient privileges") {
                Write-ErrorMsg "Microsoft Graph returned 403 Forbidden."
                Write-Host ""
                Write-Warn "The current Graph session lacks the required permissions."
                Write-InfoMsg "This usually means admin consent was not granted for the app."
                Write-Host ""

                if ($attempt -eq 1) {
                    $fix = Show-Menu -Title "How to fix?" -Options @(
                        "Reconnect Graph with fresh consent prompt",
                        "Show manual admin consent instructions"
                    ) -BackLabel "Cancel"

                    if ($fix -eq 0) {
                        if (Reconnect-GraphWithConsent) {
                            Write-InfoMsg "Retrying search..."
                            continue   # retry the for loop
                        } else {
                            Write-ErrorMsg "Reconnect failed."
                            return $null
                        }
                    }
                    elseif ($fix -eq 1) {
                        Write-Host ""
                        Write-InfoMsg "An Azure AD admin needs to grant consent:"
                        Write-Host ""
                        Write-Host "  1. Go to: https://entra.microsoft.com" -ForegroundColor White
                        Write-Host "  2. App registrations > Microsoft Graph PowerShell" -ForegroundColor White
                        Write-Host "  3. API permissions > Grant admin consent" -ForegroundColor White
                        Write-Host ""
                        Write-Host "  Required permissions:" -ForegroundColor $script:Colors.Info
                        Write-Host "    User.ReadWrite.All" -ForegroundColor White
                        Write-Host "    Group.ReadWrite.All" -ForegroundColor White
                        Write-Host "    Directory.ReadWrite.All" -ForegroundColor White
                        Write-Host "    Organization.Read.All" -ForegroundColor White
                        Write-Host ""
                        Write-InfoMsg "After granting consent, restart the tool."
                        return $null
                    }
                    else { return $null }
                }
            }

            Write-ErrorMsg "Could not find user: $errMsg"
            return $null
        }
    }
    return $null
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
