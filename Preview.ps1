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
                       Single-statement is typical; multi-statement
                       is fine.
        -Critical    : when set, re-throws on failure in LIVE mode
                       (default is warn+continue+return null). Use
                       for steps the rest of the workflow depends
                       on, e.g. account creation.
        -StubReturn  : value returned when in PREVIEW mode. Useful
                       when downstream code reads a property
                       (e.g. $newUser.Id) -- pass a PSCustomObject
                       with the fields the caller needs.

        Always returns either the action's result (LIVE), the stub
        (PREVIEW), or $null (LIVE failure, non-critical).
    #>
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][scriptblock]$Action,
        [switch]$Critical,
        $StubReturn = $null
    )

    if ($script:PreviewMode) {
        Write-Host ("  [PREVIEW] Would: {0}" -f $Description) -ForegroundColor Yellow
        Write-AuditEntry -EventType 'PREVIEW' -Detail $Description
        return $StubReturn
    }

    Write-AuditEntry -EventType 'EXEC' -Detail $Description
    try {
        return (& $Action)
    } catch {
        $msg = $_.Exception.Message
        Write-AuditEntry -EventType 'ERROR' -Detail ("{0} :: {1}" -f $Description, $msg)
        if ($Critical) { throw }
        Write-ErrorMsg ("{0} failed: {1}" -f $Description, $msg)
        return $null
    }
}
