param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings')

# Loads BreakGlass.ps1 from Phase 4 Commit C; falls back to an
# empty list if the module isn't present yet.
$root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive `
            -Modules 'UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1','SignInLookup.ps1','BreakGlass.ps1'

Write-Host "==> health-breakglass-signins"
try {
    if (-not (Get-Command Get-BreakGlassAccounts -ErrorAction SilentlyContinue)) {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'breakglass-signins' -Status 'failure' -Note 'BreakGlass module not loaded' | Out-Null
        exit 1
    }
    if (-not (Connect-ForTask 'Report')) {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'breakglass-signins' -Status 'failure' -Note 'connect failed' | Out-Null
        exit 1
    }
    $accounts = @(Get-BreakGlassAccounts)
    $hits = New-Object System.Collections.ArrayList
    foreach ($a in $accounts) {
        $signs = @(Search-SignIns -UPN $a.UPN -From ((Get-Date).AddDays(-1)) -To (Get-Date) -MaxResults 50)
        if ($signs.Count -gt 0) {
            foreach ($s in $signs) {
                [void]$hits.Add([PSCustomObject]@{
                    breakGlassUpn = $a.UPN
                    timeUtc       = $s.TimeUtc
                    app           = $s.App
                    ip            = $s.IpAddress
                    outcome       = $s.Outcome
                })
            }
        }
    }
    $status = if ($hits.Count -gt 0) { 'findings' } else { 'clean' }
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'breakglass-signins' -Status $status -FindingCount $hits.Count -Findings $hits | Out-Null
} catch {
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'breakglass-signins' -Status 'failure' -Note $_.Exception.Message | Out-Null
    exit 1
}
