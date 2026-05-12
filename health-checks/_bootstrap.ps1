# ============================================================
#  Shared bootstrap for health-check scripts.
#  Dot-source modules + flip non-interactive mode + return the
#  repo root so the calling script can resolve paths.
#
#  Usage in a check script:
#     param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings')
#     $root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive `
#                 -Modules 'UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1','LicenseOptimizer.ps1'
# ============================================================
param(
    [switch]$NonInteractive,
    [string[]]$Modules = @('UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1')
)

$root = Split-Path $PSScriptRoot -Parent
foreach ($m in $Modules) {
    $p = Join-Path $root $m
    if (Test-Path -LiteralPath $p) { . $p }
}
if ($NonInteractive -and (Get-Command Set-NonInteractiveMode -ErrorAction SilentlyContinue)) {
    Set-NonInteractiveMode -Enabled $true
}
return $root
