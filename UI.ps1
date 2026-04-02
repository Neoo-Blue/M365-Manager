# ============================================================
#  UI.ps1 - Shared TUI helpers (colors, menus, confirmations)
# ============================================================

# ---- Color Palette ----
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

# ---- Box-drawing chars built at runtime so encoding never matters ----
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
    $title = "            M365 Administration Tool  v1.0"
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
    param(
        [string]$Label,
        [string]$Value,
        [string]$Color = "White"
    )
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

function Read-SecureInput {
    param([string]$Prompt)
    Write-Host ""
    Write-Host "  $Prompt" -ForegroundColor $script:Colors.Prompt -NoNewline
    Write-Host ": " -ForegroundColor $script:Colors.Prompt -NoNewline
    return (Read-Host -AsSecureString)
}

function Confirm-Action {
    param(
        [string]$Message,
        [string]$Details = ""
    )

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
    param(
        [string]$Title,
        [string[]]$Options,
        [string]$BackLabel = "Back to Main Menu"
    )
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
            if ($num -eq 0)                   { return -1 }
            if ($num -ge 1 -and $num -le $Options.Count) { return ($num - 1) }
        }
        Write-ErrorMsg "Invalid selection. Please try again."
    }
}

function Show-MultiSelect {
    param(
        [string]$Title,
        [string[]]$Options,
        [string]$Prompt = "Enter selection(s) (e.g. 1,3,5)"
    )
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
                if ($idx -ge 1 -and $idx -le $Options.Count) {
                    $indices += ($idx - 1)
                } else { $valid = $false; break }
            } else { $valid = $false; break }
        }
        if ($valid -and $indices.Count -gt 0) { return $indices }
        Write-ErrorMsg "Invalid input. Use numbers separated by commas (e.g. 1,3,5)."
    }
}

function Show-UserDataTable {
    param(
        [hashtable]$Data,
        [string[]]$FieldOrder
    )

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
    param(
        [hashtable]$Data,
        [string[]]$FieldOrder
    )

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
    param([string]$PromptText = "Enter user name or email")

    $userInput = Read-UserInput $PromptText
    if ([string]::IsNullOrWhiteSpace($userInput)) { return $null }

    Write-InfoMsg "Searching for user..."

    try {
        if ($userInput -match '@') {
            $user = Get-AzureADUser -ObjectId $userInput -ErrorAction Stop
        } else {
            $users = Get-AzureADUser -SearchString $userInput -ErrorAction Stop
            if ($users.Count -eq 0) {
                Write-ErrorMsg "No users found matching '$userInput'."
                return $null
            }
            if ($users.Count -eq 1) {
                $user = $users[0]
            } else {
                Write-Warn "Multiple users found:"
                $names = $users | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
                $sel = Show-Menu -Title "Select User" -Options $names -BackLabel "Cancel"
                if ($sel -eq -1) { return $null }
                $user = $users[$sel]
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
