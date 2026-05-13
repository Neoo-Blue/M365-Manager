param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings')

$root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive `
            -Modules 'UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1','TeamsManager.ps1'

Write-Host "==> health-orphaned-teams"
try {
    if (-not (Connect-ForTask 'Teams')) {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'orphaned-teams' -Status 'failure' -Note 'connect failed' | Out-Null
        exit 1
    }
    $orphans = @(Get-OrphanedTeams)
    $single  = @(Get-SingleOwnerTeams)
    $count   = $orphans.Count + $single.Count
    $findings = @{
        orphans     = ($orphans | Select-Object TeamId,DisplayName | Select-Object -First 200)
        singleOwner = ($single  | Select-Object TeamId,DisplayName,OwnerUPN | Select-Object -First 200)
        totals      = @{ orphans = $orphans.Count; singleOwner = $single.Count }
    }
    $status = if ($count -gt 0) { 'findings' } else { 'clean' }
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'orphaned-teams' -Status $status -FindingCount $count -Findings $findings | Out-Null
} catch {
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'orphaned-teams' -Status 'failure' -Note $_.Exception.Message | Out-Null
    exit 1
}
