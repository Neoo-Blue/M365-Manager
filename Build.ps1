# ============================================================
#  Build.ps1 - Package M365Admin into a single .exe
# ============================================================
#  Run this from the M365Admin folder:
#    powershell -ExecutionPolicy Bypass -File Build.ps1
#
#  Output: M365Admin.exe in the same folder
# ============================================================

$ErrorActionPreference = "Stop"
$buildDir = $PSScriptRoot
if (-not $buildDir) { $buildDir = Split-Path -Parent $MyInvocation.MyCommand.Path }

Write-Host ""
Write-Host "=== M365 Admin Tool - Build to EXE ===" -ForegroundColor Cyan
Write-Host ""

# ---- Step 1: Install PS2EXE if missing ----
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "[1/4] Installing ps2exe module..." -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force
} else {
    Write-Host "[1/4] ps2exe module found." -ForegroundColor Green
}

# ---- Step 2: Merge all scripts into one file ----
Write-Host "[2/4] Merging scripts..." -ForegroundColor Yellow

# Order matters: UI first (shared helpers), then Auth, then feature modules, Main logic last
$scriptOrder = @(
    "UI.ps1",
    "Auth.ps1",
    "Onboard.ps1",
    "Offboard.ps1",
    "License.ps1",
    "Archive.ps1",
    "SecurityGroup.ps1",
    "DistributionList.ps1",
    "SharedMailbox.ps1",
    "CalendarAccess.ps1",
    "UserProfile.ps1",
    "Reports.ps1"
)

$merged = @()

# Add header
$merged += @"
# ============================================================
#  M365 Administration Tool v2.0 (Compiled)
#  Built: $(Get-Date -Format 'yyyy-MM-dd HH:mm')
# ============================================================

"@

# Merge each module
foreach ($file in $scriptOrder) {
    $path = Join-Path $buildDir $file
    if (Test-Path $path) {
        $merged += "# ============ $file ============"
        $content = Get-Content $path -Raw

        # Remove dot-source lines (they reference external files we're merging)
        $content = $content -replace '^\.\s+".*\\[^"]+\.ps1"', '# (merged)'
        $content = $content -replace "^\.\s+'\$ScriptRoot\\[^']+\.ps1'", '# (merged)'

        $merged += $content
        $merged += ""
        Write-Host "    + $file" -ForegroundColor Gray
    } else {
        Write-Host "    ! $file not found, skipping" -ForegroundColor DarkYellow
    }
}

# Add the Main.ps1 logic (without the dot-source lines)
$mainPath = Join-Path $buildDir "Main.ps1"
$mainContent = Get-Content $mainPath -Raw

# Remove the dot-source block and ScriptRoot detection (already merged)
$mainContent = $mainContent -replace '(?ms)\$ScriptRoot = \$PSScriptRoot.*?# ---- Dot-source all modules ----\s*', ''
$mainContent = $mainContent -replace '^\.\s+".*\\[^"]+\.ps1"\s*$', ''
$mainContent = $mainContent -replace '^\.\s+"\$ScriptRoot\\[^"]+"\s*$', ''

# Remove lines that start with dot-source patterns
$mainLines = $mainContent -split "`n" | Where-Object { $_ -notmatch '^\.\s+"\$ScriptRoot' }
$mainContent = $mainLines -join "`n"

$merged += "# ============ Main.ps1 ============"
$merged += $mainContent

# Write merged file
$mergedPath = Join-Path $buildDir "M365Admin_Merged.ps1"
$merged | Out-File -FilePath $mergedPath -Encoding UTF8 -Force
Write-Host "    Merged file: $mergedPath" -ForegroundColor Green

# ---- Step 3: Compile to EXE ----
Write-Host "[3/4] Compiling to EXE..." -ForegroundColor Yellow

$exePath = Join-Path $buildDir "M365Admin.exe"

$ps2exeParams = @{
    InputFile   = $mergedPath
    OutputFile  = $exePath
    NoConsole   = $false
    Title       = "M365 Administration Tool"
    Description = "Microsoft 365 Admin TUI - User, License, Group, Mailbox Management"
    Company     = "IT Administration"
    Version     = "2.0.0.0"
    Copyright   = "Internal Tool"
    RequireAdmin = $false
}

try {
    Invoke-PS2EXE @ps2exeParams
    Write-Host ""
    Write-Host "[4/4] Build complete!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Output: $exePath" -ForegroundColor Cyan
    Write-Host "  Size:   $([math]::Round((Get-Item $exePath).Length / 1KB)) KB" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  You can distribute this single .exe file." -ForegroundColor White
    Write-Host "  Users just double-click it - no PowerShell scripts needed." -ForegroundColor White
} catch {
    Write-Host ""
    Write-Host "  EXE compilation failed: $_" -ForegroundColor Red
    Write-Host "  The merged script is still available at:" -ForegroundColor Yellow
    Write-Host "  $mergedPath" -ForegroundColor White
    Write-Host "  You can compile it manually: Invoke-PS2EXE -InputFile '$mergedPath' -OutputFile '$exePath'" -ForegroundColor Gray
}

# ---- Cleanup merged file (optional, keep for debugging) ----
# Remove-Item $mergedPath -Force

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
