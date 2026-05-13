# Shared helper -- emit a structured health-result-<name>-<ts>.json
# next to the audit log so Get-HealthResults can find it.
param(
    [Parameter(Mandatory)][string]$CheckName,
    [Parameter(Mandatory)][string]$Status,           # clean | findings | failure
    [int]$FindingCount = 0,
    $Findings = @(),
    [string]$Note = ''
)
$dir = if (Get-Command Get-AuditLogDirectory -ErrorAction SilentlyContinue) { Get-AuditLogDirectory } else { (Get-Location).Path }
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$path  = Join-Path $dir ("health-result-{0}-{1}.json" -f $CheckName, $stamp)
$obj = [ordered]@{
    checkName     = $CheckName
    startedAtUtc  = (Get-Date).ToUniversalTime().ToString('o')
    completedAtUtc= (Get-Date).ToUniversalTime().ToString('o')
    status        = $Status
    findingCount  = $FindingCount
    note          = $Note
    findings      = $Findings
}
$obj | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $path -Encoding UTF8 -Force
Write-Host "  result: $path"
return $path
