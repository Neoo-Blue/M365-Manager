param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings')

$root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive `
            -Modules 'UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1','MFAManager.ps1'

Write-Host "==> health-mfa-gaps"
try {
    if (-not (Connect-ForTask 'Report')) {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'mfa-gaps' -Status 'failure' -Note 'connect failed' | Out-Null
        exit 1
    }
    $noMfa     = @(Get-UsersWithNoMfa        -Max 500)
    $onlyPhone = @(Get-UsersWithOnlyPhoneMfa -Max 500)
    $count     = $noMfa.Count + $onlyPhone.Count
    $findings  = @{
        noMfaUsers     = ($noMfa     | Select-Object -First 50 | ForEach-Object { $_.UserPrincipalName })
        onlyPhoneUsers = ($onlyPhone | Select-Object -First 50 | ForEach-Object { $_.UserPrincipalName })
        totals         = @{ noMfa = $noMfa.Count; onlyPhone = $onlyPhone.Count }
    }
    $status = if ($count -gt 0) { 'findings' } else { 'clean' }
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'mfa-gaps' -Status $status -FindingCount $count -Findings $findings | Out-Null
} catch {
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'mfa-gaps' -Status 'failure' -Note $_.Exception.Message | Out-Null
    exit 1
}
