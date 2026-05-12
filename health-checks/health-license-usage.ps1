param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings')

$root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive `
            -Modules 'UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1','LicenseOptimizer.ps1'

Write-Host "==> health-license-usage"
try {
    if (-not (Connect-ForTask 'Report')) {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'license-usage' -Status 'failure' -Note 'connect failed' | Out-Null
        exit 1
    }
    $r = Get-LicenseUtilizationReport -DaysInactive 60
    $findings = @()
    if ($r) {
        $findings = @{
            inactiveCount      = $r.Inactive.Count
            hoarderCount       = $r.Hoarders.Count
            paidUnassignedRows = $r.Unassigned.Count
            downgradeCount     = $r.Downgrade.Count
            dashboardPath      = $r.DashboardPath
        }
    }
    $count = ($r.Inactive.Count + $r.Hoarders.Count + $r.Downgrade.Count)
    $status = if ($count -gt 0) { 'findings' } else { 'clean' }
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'license-usage' -Status $status -FindingCount $count -Findings $findings | Out-Null
} catch {
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'license-usage' -Status 'failure' -Note $_.Exception.Message | Out-Null
    exit 1
}
