param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings')

$root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive `
            -Modules 'UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1','GuestUsers.ps1'

Write-Host "==> health-stale-guests"
try {
    if (-not (Connect-ForTask 'GuestUsers')) {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'stale-guests' -Status 'failure' -Note 'connect failed' | Out-Null
        exit 1
    }
    $stale = @(Get-StaleGuests -DaysSinceSignIn 90)
    $findings = ($stale | Select-Object UPN,DisplayName,DaysSinceSignIn,Domains | Select-Object -First 200)
    $status = if ($stale.Count -gt 0) { 'findings' } else { 'clean' }
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'stale-guests' -Status $status -FindingCount $stale.Count -Findings $findings | Out-Null
} catch {
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'stale-guests' -Status 'failure' -Note $_.Exception.Message | Out-Null
    exit 1
}
