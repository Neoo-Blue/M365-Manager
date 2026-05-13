# ============================================================
#  Audit.ps1 — General audit log for state-mutating operations
#
#  Distinct from the AI assistant's per-session log (mark-*.log).
#  Every Invoke-Action call (Preview.ps1), every Bulk* operation,
#  and any other module that wants a permanent record writes here
#  via Write-AuditEntry.
#
#  Log path:
#    Windows : %LOCALAPPDATA%\M365Manager\audit\session-<ts>-<pid>.log
#    POSIX   : ~/.m365manager/audit/session-<ts>-<pid>.log
#  Directory mode 0700 on POSIX; inherits user-only NTFS ACL from
#  %LOCALAPPDATA% on Windows.
# ============================================================

$script:AuditLogPath = $null

function Get-AuditLogPath {
    if ($script:AuditLogPath -and (Test-Path -LiteralPath (Split-Path $script:AuditLogPath -Parent))) {
        return $script:AuditLogPath
    }
    $base = $null
    $onWindows = $false
    if ($env:LOCALAPPDATA) {
        $base = Join-Path $env:LOCALAPPDATA 'M365Manager\audit'
        $onWindows = $true
    } elseif ($env:HOME) {
        $base = Join-Path $env:HOME '.m365manager/audit'
    } else {
        $base = Join-Path (Get-Location).Path 'audit'
    }
    $createdNow = $false
    try {
        if (-not (Test-Path -LiteralPath $base)) {
            New-Item -ItemType Directory -Path $base -Force | Out-Null
            $createdNow = $true
        }
    } catch { return $null }
    if ($createdNow -and -not $onWindows -and (Get-Command chmod -ErrorAction SilentlyContinue)) {
        try { & chmod 700 $base 2>$null | Out-Null } catch {}
    }
    $name = "session-{0}-{1}.log" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'), $PID
    $script:AuditLogPath = Join-Path $base $name
    return $script:AuditLogPath
}

function Write-AuditEntry {
    <#
        Append one timestamped line to the session audit log. The
        line is tagged with MODE=PREVIEW or MODE=LIVE so reviewers
        can grep distinct sets of entries.
    #>
    param(
        [Parameter(Mandatory)][string]$EventType,
        [Parameter(Mandatory)][string]$Detail
    )
    $path = Get-AuditLogPath
    if (-not $path) { return }
    $modeTag = if ($script:PreviewMode) { 'MODE=PREVIEW' } else { 'MODE=LIVE' }
    $line = "[{0}] [{1}] [{2}] {3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $EventType, $modeTag, $Detail
    try { Add-Content -LiteralPath $path -Value $line -ErrorAction Stop } catch {}
}

function Write-AuditBanner {
    <#
        Write a multi-line banner so anyone tailing the log knows
        which session started, in which mode, against which tenant.
        Called from Main.ps1 right after the mode picker.
    #>
    $path = Get-AuditLogPath
    if (-not $path) { return }
    $modeLabel = if ($script:PreviewMode) { 'PREVIEW (no tenant changes will be made)' } else { 'LIVE (tenant changes WILL be made)' }
    $tenant = 'unknown'
    if ($script:SessionState) {
        $tenant = "$($script:SessionState.TenantMode) / $($script:SessionState.TenantName)"
    }
    $sep = ('-' * 60)
    $banner = @(
        $sep,
        "M365 Manager session START",
        ("Mode    : " + $modeLabel),
        ("Tenant  : " + $tenant),
        ("PID     : " + $PID),
        ("Started : " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')),
        $sep
    )
    try { Add-Content -LiteralPath $path -Value $banner -ErrorAction Stop } catch {}
}
