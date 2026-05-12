param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings')

$root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive `
            -Modules 'UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1'

Write-Host "==> health-conditional-access-conflicts"
try {
    if (-not (Connect-ForTask 'Report')) {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'conditional-access-conflicts' -Status 'failure' -Note 'connect failed' | Out-Null
        exit 1
    }

    $policies = @()
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -ErrorAction Stop
        $policies = @($resp.value)
    } catch {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'conditional-access-conflicts' -Status 'failure' -Note ("Graph CA query failed: " + $_.Exception.Message) | Out-Null
        exit 1
    }

    $findings = New-Object System.Collections.ArrayList

    # 1. Disabled policies that look security-critical (tags / name)
    foreach ($p in $policies) {
        $name = [string]$p.displayName
        $isCritical = ($name -match '(?i)mfa|conditional|legacy|require|block|baseline')
        if ($p.state -ne 'enabled' -and $isCritical) {
            [void]$findings.Add([PSCustomObject]@{
                kind   = 'disabled-critical'
                name   = $name
                state  = $p.state
                policyId = $p.id
            })
        }
    }

    # 2. "All users" excludes with a very large excluded group set
    foreach ($p in $policies) {
        if ($p.conditions.users.includeUsers -contains 'All') {
            $exGroups = @($p.conditions.users.excludeGroups)
            $exUsers  = @($p.conditions.users.excludeUsers)
            if (($exGroups.Count + $exUsers.Count) -ge 3) {
                [void]$findings.Add([PSCustomObject]@{
                    kind = 'all-users-with-large-exclusion'
                    name = $p.displayName
                    excludedGroups = $exGroups.Count
                    excludedUsers  = $exUsers.Count
                    policyId = $p.id
                })
            }
        }
    }

    # 3. Missing legacy-auth block (any enabled policy that
    #    blocks clientAppTypes "exchangeActiveSync" / "other")
    $hasLegacyBlock = $false
    foreach ($p in $policies) {
        if ($p.state -ne 'enabled') { continue }
        $apps = @($p.conditions.clientAppTypes)
        $grant = @($p.grantControls.builtInControls)
        if (($apps -contains 'exchangeActiveSync' -or $apps -contains 'other') -and ($grant -contains 'block')) {
            $hasLegacyBlock = $true; break
        }
    }
    if (-not $hasLegacyBlock) {
        [void]$findings.Add([PSCustomObject]@{ kind = 'no-legacy-auth-block'; note = 'No enabled CA policy blocks exchangeActiveSync/other client app types.' })
    }

    $count = $findings.Count
    $status = if ($count -gt 0) { 'findings' } else { 'clean' }
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'conditional-access-conflicts' -Status $status -FindingCount $count -Findings ($findings | Select-Object -First 50) | Out-Null
} catch {
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'conditional-access-conflicts' -Status 'failure' -Note $_.Exception.Message | Out-Null
    exit 1
}
