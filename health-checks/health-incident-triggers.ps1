param([switch]$NonInteractive, [string]$Output='file', [string]$NotifyOn='findings', [string[]]$UPNs)

$root = & "$PSScriptRoot/_bootstrap.ps1" -NonInteractive:$NonInteractive `
            -Modules 'UI.ps1','Auth.ps1','Audit.ps1','Preview.ps1','Notifications.ps1','SignInLookup.ps1','UnifiedAuditLog.ps1','SharePoint.ps1','MFAManager.ps1','IncidentResponse.ps1','IncidentTriggers.ps1'

Write-Host "==> health-incident-triggers"
try {
    if (-not (Connect-ForTask 'Report')) {
        & "$PSScriptRoot/_writeresult.ps1" -CheckName 'incident-triggers' -Status 'failure' -Note 'connect failed' | Out-Null
        exit 1
    }

    # Scope: pass a list of UPNs to scan, OR scan all enabled users
    # by default. Large tenants should ALWAYS pass -UPNs to keep the
    # 15-minute scheduler interval realistic.
    $args = @{ NonInteractive = $NonInteractive }
    if ($UPNs -and $UPNs.Count -gt 0) { $args.UPNs = $UPNs } else { $args.All = $true }

    $findings = @(Invoke-IncidentDetectors @args)

    $status = if ($findings.Count -gt 0) { 'findings' } else { 'clean' }
    & "$PSScriptRoot/_writeresult.ps1" `
        -CheckName 'incident-triggers' `
        -Status $status `
        -FindingCount $findings.Count `
        -Findings @{
            findingCount   = $findings.Count
            byTriggerType  = ($findings | Group-Object -Property { $_.TriggerType } | ForEach-Object { @{ type = $_.Name; count = $_.Count } })
            bySeverity     = ($findings | Group-Object -Property { $_.Severity }     | ForEach-Object { @{ severity = $_.Name; count = $_.Count } })
            details        = ($findings | Select-Object -First 25)
        } | Out-Null
} catch {
    & "$PSScriptRoot/_writeresult.ps1" -CheckName 'incident-triggers' -Status 'failure' -Note $_.Exception.Message | Out-Null
    exit 1
}
