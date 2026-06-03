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
#
#  Format: JSON-per-line (NDJSON / JSONL). One complete JSON object
#  per line. See docs/audit-format.md for the field reference.
#  Pre-Phase-2 logs are human-readable text; AuditViewer.ps1's
#  parser handles both shapes transparently.
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
    $tenantSlug = ''
    if ($script:SessionState -and $script:SessionState.TenantName -and $script:SessionState.TenantName -ne 'Own Tenant') {
        $tenantSlug = '-' + (($script:SessionState.TenantName -replace '[^A-Za-z0-9]+','_').ToLower())
    }
    $name = "session-{0}-{1}{2}.log" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'), $PID, $tenantSlug
    $script:AuditLogPath = Join-Path $base $name
    return $script:AuditLogPath
}

function Reset-AuditLogPath {
    <# Phase 6: called by Switch-Tenant so audit entries land in a
       tenant-suffixed file after the switch, not in the prior
       tenant's session log. #>
    $script:AuditLogPath = $null
}

function Get-AuditLogDirectory {
    if ($env:LOCALAPPDATA) { return Join-Path $env:LOCALAPPDATA 'M365Manager\audit' }
    if ($env:HOME) { return Join-Path $env:HOME '.m365manager/audit' }
    return Join-Path (Get-Location).Path 'audit'
}

function New-AuditEntryId {
    return [guid]::NewGuid().ToString()
}

function Write-AuditEntry {
    <#
        Append one JSON-line entry to the session audit log.

        Backward-compatible: callers that pass only -EventType and
        -Detail still work; the record is built with just those two
        fields populated and the rest left null. New structured
        callers (Invoke-Action, undo flows, etc.) pass the full set.

        Returns the generated entryId so callers (e.g. Invoke-Action)
        can correlate ERROR / OK entries with the original PROPOSE /
        EXEC, and so Undo can reference an entry by id.
    #>
    param(
        [Parameter(Mandatory)][string]$EventType,
        [Parameter(Mandatory)][string]$Detail,
        [string]$ActionType,
        [hashtable]$Target,
        [string]$Result,
        [string]$ErrorMessage,
        [hashtable]$Reverse,
        [string]$NoUndoReason,
        [string]$EntryId
    )
    $path = Get-AuditLogPath
    if (-not $path) { return $null }

    if (-not $EntryId) { $EntryId = New-AuditEntryId }
    $modeTag = if ($script:PreviewMode) { 'PREVIEW' } else { 'LIVE' }

    # Phase 6: structured tenant fingerprint -- every entry records
    # the human name + the AAD GUID so cross-tenant audit forensics
    # can match either field cleanly. Legacy callers that only had
    # a domain still land that value under .tenant.domain.
    $tenant = $null
    if ($script:SessionState) {
        $tenant = [ordered]@{
            name   = [string]$script:SessionState.TenantName
            id     = [string]$script:SessionState.TenantId
            domain = [string]$script:SessionState.TenantDomain
            mode   = [string]$script:SessionState.TenantMode
        }
    }

    $record = [ordered]@{
        ts            = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        entryId       = $EntryId
        mode          = $modeTag
        event         = $EventType
        description   = $Detail
        actionType    = $ActionType
        target        = $Target
        result        = $Result
        error         = $ErrorMessage
        tenant        = $tenant
        session       = $PID
        reverse       = $Reverse
        noUndoReason  = $NoUndoReason
    }
    try {
        $json = $record | ConvertTo-Json -Depth 8 -Compress
        Add-Content -LiteralPath $path -Value $json -ErrorAction Stop
    } catch {}
    return $EntryId
}

function Write-AuditBanner {
    <#
        Write a top-of-session marker so anyone tailing the log knows
        which session started, in which mode, against which tenant.
        Emitted as a JSON line with event=SESSION_START so the viewer
        can group entries by session.
    #>
    $modeLabel = if ($script:PreviewMode) { 'PREVIEW' } else { 'LIVE' }
    $tenant = 'unknown'
    if ($script:SessionState) {
        $tenant = "$($script:SessionState.TenantMode) / $($script:SessionState.TenantName)"
    }
    Write-AuditEntry -EventType 'SESSION_START' -Detail "M365 Manager session start ($modeLabel) -- $tenant" -ActionType 'SessionStart' -Target @{ pid = $PID; mode = $modeLabel; tenant = $tenant } -Result 'info' | Out-Null
}
