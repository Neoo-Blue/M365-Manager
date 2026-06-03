# ============================================================
#  Preview.ps1 — Dry-run / preview-mode wrapper
#
#  Invoke-Action is the single point of execution for every
#  state-mutating Graph / EXO / AzureAD call. In LIVE mode it runs
#  the scriptblock and returns its value; in PREVIEW mode it logs
#  "[PREVIEW] Would: <description>" to host + audit log and returns
#  the StubReturn value (or $null) WITHOUT executing.
#
#  $script:PreviewMode is the session-scoped flag controlled by
#  the Main.ps1 mode picker (and toggled per-call by -WhatIf on
#  Invoke-BulkOnboard / Invoke-BulkOffboard).
#
#  Phase 2: Invoke-Action now also accepts -ActionType (a stable
#  short tag like "AssignLicense") and -Target (free-form hashtable
#  with the operands) so the JSONL audit log can be filtered cleanly
#  by AuditViewer.ps1 and the undo subsystem can dispatch by
#  actionType.
#
#  -Reverse / -ReverseDescription / -NoUndoReason are accepted and
#  threaded through to the audit record; they are consumed by
#  Undo.ps1 (commit B). When -NoUndoReason is set, the audit record
#  carries no reverse recipe and the entry is flagged as
#  non-reversible in the viewer.
# ============================================================

if ($null -eq (Get-Variable -Name PreviewMode -Scope Script -ErrorAction SilentlyContinue)) {
    $script:PreviewMode = $false
}

function Set-PreviewMode {
    param([bool]$Enabled)
    $script:PreviewMode = $Enabled
}

function Get-PreviewMode { return [bool]$script:PreviewMode }

function Use-PreviewModeScope {
    <#
        Run $Action with $script:PreviewMode forced to $Enabled for
        the duration. Restores the previous value on exit. Used by
        Invoke-Bulk* when -WhatIf is passed so the rest of the
        session keeps its own mode.
    #>
    param([Parameter(Mandatory)][bool]$Enabled, [Parameter(Mandatory)][scriptblock]$Action)
    $prev = $script:PreviewMode
    $script:PreviewMode = $Enabled
    try { & $Action }
    finally { $script:PreviewMode = $prev }
}

function Invoke-Action {
    <#
        Run a state-mutating action, gated by $script:PreviewMode.

        -Description : short human-readable string ("Block sign-in
                       for $upn") -- shown to operator and written
                       to the audit log.
        -Action      : scriptblock containing the mutation cmdlet.
        -ActionType  : optional stable short tag (e.g. "BlockSignIn",
                       "AssignLicense"). Used by AuditViewer for
                       filtering and by Undo for handler dispatch.
        -Target      : optional hashtable of operands (e.g.
                       @{ userUpn=$u; groupId=$g }). Goes into the
                       JSONL audit record. Treat keys as stable
                       contract for downstream tooling.
        -Critical    : when set, re-throws on failure in LIVE mode
                       (default is warn+continue+return null). Use
                       for steps the rest of the workflow depends
                       on, e.g. account creation.
        -StubReturn  : value returned when in PREVIEW mode. Useful
                       when downstream code reads a property
                       (e.g. $newUser.Id) -- pass a PSCustomObject
                       with the fields the caller needs.
        -ReverseType / -ReverseDescription / -ReverseTarget
                     : optional inverse-op recipe, consumed by
                       Undo.ps1 (commit B). If omitted, the entry
                       has reverse=null.
        -NoUndoReason: optional human-readable explanation for why
                       this op cannot be undone (e.g. "User deletion
                       is irreversible"). Mutually exclusive with
                       reverse-* params.

        Always returns either the action's result (LIVE), the stub
        (PREVIEW), or $null (LIVE failure, non-critical).
    #>
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][scriptblock]$Action,
        [string]$ActionType,
        [hashtable]$Target,
        [switch]$Critical,
        $StubReturn = $null,
        [string]$ReverseType,
        [string]$ReverseDescription,
        [hashtable]$ReverseTarget,
        [string]$NoUndoReason
    )

    $entryId = New-AuditEntryId
    $reverse = $null
    if ($ReverseType -and -not $NoUndoReason) {
        $reverse = @{
            type        = $ReverseType
            description = $ReverseDescription
            target      = if ($ReverseTarget) { $ReverseTarget } else { $Target }
        }
    }

    if ($script:PreviewMode) {
        Write-Host ("  [PREVIEW] Would: {0}" -f $Description) -ForegroundColor Yellow
        Write-AuditEntry -EventType 'PREVIEW' -Detail $Description `
            -ActionType $ActionType -Target $Target -Result 'preview' `
            -Reverse $reverse -NoUndoReason $NoUndoReason -EntryId $entryId | Out-Null
        return $StubReturn
    }

    try {
        $result = & $Action
        Write-AuditEntry -EventType 'EXEC' -Detail $Description `
            -ActionType $ActionType -Target $Target -Result 'success' `
            -Reverse $reverse -NoUndoReason $NoUndoReason -EntryId $entryId | Out-Null
        return $result
    } catch {
        $msg = $_.Exception.Message
        Write-AuditEntry -EventType 'EXEC' -Detail $Description `
            -ActionType $ActionType -Target $Target -Result 'failure' `
            -ErrorMessage $msg -Reverse $null -NoUndoReason $NoUndoReason -EntryId $entryId | Out-Null
        if ($Critical) { throw }
        Write-ErrorMsg ("{0} failed: {1}" -f $Description, $msg)
        return $null
    }
}
