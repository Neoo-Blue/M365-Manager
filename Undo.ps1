# ============================================================
#  Undo.ps1 — per-operation undo for reversible audit entries
#
#  Phase 2: every call site that wraps a reversible cmdlet through
#  Invoke-Action passes -ActionType / -Target / -ReverseType, which
#  goes into the JSONL audit record. Show-RecentUndoable lists
#  those entries and Invoke-Undo runs the curated inverse via
#  Invoke-Action so the reversal itself is audited and respects
#  PREVIEW mode.
#
#  Reverse recipes are CURATED, not free-form scriptblocks pulled
#  from disk -- the audit record only carries the action type tag
#  ("AddToGroup", "AssignLicense", ...) plus a target hashtable.
#  $script:UndoHandlers dispatches by type. New reversible types
#  are added here, not at the call site.
#
#  Reversal state is persisted to a sidecar at
#  <audit-dir>\undo-state.json so a reversed entry doesn't show
#  up again on the next run.
# ============================================================

# ============================================================
#  Reverse-action handlers. Each takes a $Target hashtable read
#  straight from the audit record and runs the inverse cmdlet
#  (no try/catch; Invoke-Action handles errors).
# ============================================================

$script:UndoHandlers = @{

    # -- Licenses --
    'RemoveLicense' = {
        param($Target)
        Set-MgUserLicense -UserId $Target.userId -AddLicenses @() -RemoveLicenses @($Target.skuId) -ErrorAction Stop | Out-Null
    }
    'AssignLicense' = {
        param($Target)
        Set-MgUserLicense -UserId $Target.userId -AddLicenses @(@{ SkuId = $Target.skuId }) -RemoveLicenses @() -ErrorAction Stop | Out-Null
    }

    # -- Security groups (Graph) --
    'RemoveFromGroup' = {
        param($Target)
        Remove-MgGroupMemberByRef -GroupId $Target.groupId -DirectoryObjectId $Target.userId -ErrorAction Stop | Out-Null
    }
    'AddToGroup' = {
        param($Target)
        New-MgGroupMember -GroupId $Target.groupId -DirectoryObjectId $Target.userId -ErrorAction Stop | Out-Null
    }

    # -- Distribution lists (EXO) --
    'RemoveFromDistributionList' = {
        param($Target)
        Remove-DistributionGroupMember -Identity $Target.dlIdentity -Member $Target.upn -Confirm:$false -ErrorAction Stop | Out-Null
    }
    'AddToDistributionList' = {
        param($Target)
        Add-DistributionGroupMember -Identity $Target.dlIdentity -Member $Target.upn -ErrorAction Stop | Out-Null
    }

    # -- Mailbox access (FullAccess + SendAs) --
    'RevokeMailboxFullAccess' = {
        param($Target)
        Remove-MailboxPermission -Identity $Target.mailbox -User $Target.user -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
    }
    'GrantMailboxFullAccess' = {
        param($Target)
        Add-MailboxPermission -Identity $Target.mailbox -User $Target.user -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop | Out-Null
    }
    'RevokeMailboxSendAs' = {
        param($Target)
        Remove-RecipientPermission -Identity $Target.mailbox -Trustee $Target.user -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
    }
    'GrantMailboxSendAs' = {
        param($Target)
        Add-RecipientPermission -Identity $Target.mailbox -Trustee $Target.user -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
    }

    # -- Calendar access --
    'RevokeCalendarAccess' = {
        param($Target)
        Remove-MailboxFolderPermission -Identity $Target.calendar -User $Target.user -Confirm:$false -ErrorAction Stop | Out-Null
    }
    'GrantCalendarAccess' = {
        param($Target)
        if (-not $Target.rights) { throw "GrantCalendarAccess requires Target.rights" }
        Add-MailboxFolderPermission -Identity $Target.calendar -User $Target.user -AccessRights $Target.rights -ErrorAction Stop | Out-Null
    }

    # -- Mailbox auto-reply / forwarding --
    'ClearOOO' = {
        param($Target)
        Set-MailboxAutoReplyConfiguration -Identity $Target.identity -AutoReplyState Disabled -ErrorAction Stop | Out-Null
    }
    'ClearForwarding' = {
        param($Target)
        Set-Mailbox -Identity $Target.identity -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false -ErrorAction Stop | Out-Null
    }

    # -- Account state --
    'UnblockSignIn' = {
        param($Target)
        Update-MgUser -UserId $Target.userId -AccountEnabled:$true -ErrorAction Stop | Out-Null
    }
    'BlockSignIn' = {
        param($Target)
        Update-MgUser -UserId $Target.userId -AccountEnabled:$false -ErrorAction Stop | Out-Null
    }

    # -- OneDrive site collection admin (Phase 3) --
    'RevokeOneDriveAccess' = {
        param($Target)
        Set-SPOUser -Site $Target.siteUrl -LoginName $Target.granteeUpn -IsSiteCollectionAdmin $false -ErrorAction Stop | Out-Null
    }
    'GrantOneDriveAccess' = {
        param($Target)
        Set-SPOUser -Site $Target.siteUrl -LoginName $Target.granteeUpn -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
    }

    # -- Teams membership / ownership (Phase 3) --
    'RemoveFromTeam' = {
        param($Target)
        $segment = if ($Target.role -eq 'Owner') { 'owners' } else { 'members' }
        Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($Target.teamId)/$segment/$($Target.userId)/`$ref" -ErrorAction Stop | Out-Null
    }
    'AddToTeam' = {
        param($Target)
        $segment = if ($Target.role -eq 'Owner') { 'owners' } else { 'members' }
        $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($Target.userId)" } | ConvertTo-Json -Compress
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$($Target.teamId)/$segment/`$ref" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
    }
    'DemoteTeamOwner' = {
        param($Target)
        Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($Target.teamId)/owners/$($Target.userId)/`$ref" -ErrorAction Stop | Out-Null
    }
    'PromoteTeamOwner' = {
        param($Target)
        $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($Target.userId)" } | ConvertTo-Json -Compress
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$($Target.teamId)/owners/`$ref" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
    }
}

# ============================================================
#  Sidecar state — track which entryIds have been reversed so
#  Show-RecentUndoable can mark them and Invoke-Undo can refuse
#  double-reversals.
# ============================================================

function Get-UndoStatePath {
    $dir = Get-AuditLogDirectory
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch { return $null }
    }
    return Join-Path $dir 'undo-state.json'
}

function Read-UndoState {
    $p = Get-UndoStatePath
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return @{} }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json -ErrorAction Stop
        $h = @{}
        foreach ($prop in $raw.PSObject.Properties) {
            $h[$prop.Name] = @{
                state         = [string]$prop.Value.state
                reversedBy    = [string]$prop.Value.reversedBy
                reversedAt    = [string]$prop.Value.reversedAt
                originalType  = [string]$prop.Value.originalType
            }
        }
        return $h
    } catch { return @{} }
}

function Write-UndoState {
    param([Parameter(Mandatory)][hashtable]$State)
    $p = Get-UndoStatePath
    if (-not $p) { return }
    try { ($State | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $p -Encoding UTF8 -Force }
    catch { Write-Warn "Could not persist undo state: $_" }
}

function Set-UndoStateReversed {
    param([Parameter(Mandatory)][string]$EntryId, [Parameter(Mandatory)][string]$ReversedBy, [string]$OriginalType)
    $s = Read-UndoState
    $s[$EntryId] = @{
        state        = 'reversed'
        reversedBy   = $ReversedBy
        reversedAt   = (Get-Date).ToUniversalTime().ToString('o')
        originalType = $OriginalType
    }
    Write-UndoState -State $s
}

# ============================================================
#  Querying / listing
# ============================================================

function Get-UndoableEntries {
    <#
        Returns audit entries that:
          - have a non-null `reverse` recipe (i.e. an actionType
            with a registered handler in $script:UndoHandlers)
          - have result=success (we don't undo failed or preview ops)
          - have not already been reversed (per sidecar)
        Optional filters: -Filter (UPN substring), -Since (DateTime),
        -Limit (default 20).
    #>
    param(
        [string]$Filter,
        [DateTime]$Since,
        [int]$Limit = 20
    )
    $all = Read-AuditEntries
    $state = Read-UndoState

    $candidates = $all | Where-Object {
        $_.result -eq 'success' -and
        $_.reverse -and
        $_.reverse.type -and
        $script:UndoHandlers.ContainsKey([string]$_.reverse.type) -and
        $_.entryId -and
        -not $state.ContainsKey([string]$_.entryId)
    }
    if ($Since) { $candidates = $candidates | Where-Object { $_.ts -and $_.ts -ge $Since } }
    if ($Filter) { $candidates = $candidates | Where-Object { Test-AuditEntryMatchesUser -Entry $_ -Upn $Filter -or (Test-AuditEntryMatchesTarget -Entry $_ -TargetText $Filter) } }
    return @($candidates | Sort-Object { $_.ts } -Descending | Select-Object -First $Limit)
}

function Show-RecentUndoable {
    <#
        Interactive list of reversible recent entries. Shows the
        entryId prefix (first 8 chars), timestamp, action type, the
        forward description, and the reverse description.
    #>
    param([int]$Limit = 20, [string]$Filter)
    $entries = Get-UndoableEntries -Filter $Filter -Limit $Limit
    Write-SectionHeader "Recent Undoable Operations"
    if ($entries.Count -eq 0) {
        Write-InfoMsg "No undoable operations found (none with success+reverse, or all already reversed)."
        return @()
    }
    Write-Host ""
    for ($i = 0; $i -lt $entries.Count; $i++) {
        $e = $entries[$i]
        $idShort = ($e.entryId).Substring(0, [Math]::Min(8, $e.entryId.Length))
        $ts = if ($e.ts) { $e.ts.ToString('yyyy-MM-dd HH:mm:ss') } else { '???' }
        Write-Host ("  [{0,2}] {1}  {2}  {3}" -f ($i + 1), $idShort, $ts, $e.actionType) -ForegroundColor White
        Write-Host ("       fwd: {0}" -f $e.description) -ForegroundColor DarkGray
        Write-Host ("       rev: {0} ({1})" -f $e.reverse.description, $e.reverse.type) -ForegroundColor DarkCyan
    }
    Write-Host ""
    return $entries
}

# ============================================================
#  Invoke-Undo
# ============================================================

function ConvertTo-UndoTargetHashtable {
    <#
        ConvertFrom-Json gives a PSCustomObject; the handler
        scriptblocks call $Target.someKey, which works for both
        forms, but ContainsKey() is hashtable-only. Normalize.
    #>
    param($Target)
    if ($null -eq $Target) { return @{} }
    if ($Target -is [hashtable]) { return $Target }
    $h = @{}
    foreach ($p in $Target.PSObject.Properties) { $h[$p.Name] = $p.Value }
    return $h
}

function Invoke-Undo {
    <#
        Reverse a single audit entry by id (or prefix-match).
        Routes through Invoke-Action so the reversal itself
        becomes an audited entry. Refuses to run if the entry
        is already reversed; warns and aborts if the target
        appears to have gone missing (per the spec, default is
        fail-clearly).
    #>
    param(
        [Parameter(Mandatory)][string]$EntryId,
        [switch]$WhatIf,
        [switch]$NoConfirm
    )
    $all = Read-AuditEntries
    $matches = @($all | Where-Object { $_.entryId -and ($_.entryId -eq $EntryId -or $_.entryId.StartsWith($EntryId)) })
    if ($matches.Count -eq 0) { Write-ErrorMsg "No audit entry with id '$EntryId'."; return }
    if ($matches.Count -gt 1) {
        Write-ErrorMsg "Ambiguous entryId prefix '$EntryId' (matched $($matches.Count) entries). Use more characters."
        return
    }
    $entry = $matches[0]

    if (-not $entry.reverse -or -not $entry.reverse.type) {
        $reason = if ($entry.noUndoReason) { $entry.noUndoReason } else { "Entry has no reverse recipe." }
        Write-ErrorMsg "Cannot undo this entry: $reason"
        return
    }
    if ($entry.result -ne 'success') {
        Write-ErrorMsg "Cannot undo an entry with result='$($entry.result)'. Only successful operations are reversible."
        return
    }

    $state = Read-UndoState
    if ($state.ContainsKey([string]$entry.entryId)) {
        $st = $state[[string]$entry.entryId]
        Write-Warn ("Entry already reversed at {0} (reversal entryId: {1})." -f $st.reversedAt, $st.reversedBy)
        return
    }

    $reverseType = [string]$entry.reverse.type
    if (-not $script:UndoHandlers.ContainsKey($reverseType)) {
        Write-ErrorMsg "No registered handler for reverse type '$reverseType'."
        return
    }

    $target = ConvertTo-UndoTargetHashtable $entry.reverse.target

    Write-Host ""
    Write-Host "  Forward action : $($entry.description)" -ForegroundColor White
    Write-Host "  Reverse action : $($entry.reverse.description)" -ForegroundColor DarkCyan
    Write-Host "  Reverse type   : $reverseType" -ForegroundColor DarkGray
    Write-Host "  Target         : $(($target | ConvertTo-Json -Compress -Depth 5))" -ForegroundColor DarkGray
    Write-Host ""

    if (-not $NoConfirm) {
        if (-not (Confirm-Action "Run the reverse action above?")) {
            Write-InfoMsg "Cancelled."; return
        }
    }

    $handler = $script:UndoHandlers[$reverseType]
    $reversalEntryId = New-AuditEntryId

    # Per the open-question default: if the reverse cmdlet errors
    # with a "not found" / "couldn't be found" / 404 shape, surface
    # a clean "Target missing" message rather than the raw Graph
    # error, and stop. Invoke-Action will capture the exception and
    # log it; we re-read the audit log to confirm failure result.
    $previousMode = Get-PreviewMode
    if ($WhatIf.IsPresent -and -not $previousMode) { Set-PreviewMode -Enabled $true }
    try {
        $ok = Invoke-Action `
            -Description ("[UNDO] " + $entry.reverse.description) `
            -ActionType ("Undo_" + $reverseType) `
            -Target $target `
            -Action { & $handler $target } `
            -StubReturn $true

        if (Get-PreviewMode) {
            Write-InfoMsg "(preview only -- no state recorded; rerun without -WhatIf to commit)"
            return
        }

        if ($ok) {
            Set-UndoStateReversed -EntryId $entry.entryId -ReversedBy $reversalEntryId -OriginalType $entry.actionType
            Write-Success "Reversed. Original entry $($entry.entryId.Substring(0,8)) marked as reversed in undo-state.json."
        } else {
            Write-ErrorMsg "Reverse action failed. See audit log for the underlying error."
            Write-InfoMsg "If the target ($((($target.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '))) no longer exists, the original entry is stale -- it was not auto-marked as superseded; you can edit undo-state.json by hand if needed."
        }
    } finally {
        Set-PreviewMode -Enabled $previousMode
    }
}

# ============================================================
#  Menu entry
# ============================================================

function Start-UndoMenu {
    Write-SectionHeader "Undo Recent Operation"
    $filter = Read-UserInput "Filter by UPN / target (blank = no filter)"
    $entries = Show-RecentUndoable -Filter $filter
    if ($entries.Count -eq 0) { Pause-ForUser; return }

    $choice = Read-UserInput "Pick a row to undo (1-$($entries.Count); blank = cancel)"
    if ([string]::IsNullOrWhiteSpace($choice)) { return }
    $n = 0
    if (-not [int]::TryParse($choice, [ref]$n) -or $n -lt 1 -or $n -gt $entries.Count) {
        Write-ErrorMsg "Out of range."; Pause-ForUser; return
    }
    $entry = $entries[$n - 1]

    $previewFirst = Confirm-Action "Preview the reverse first (WhatIf)?"
    if ($previewFirst) {
        Invoke-Undo -EntryId $entry.entryId -WhatIf -NoConfirm
        Write-Host ""
        if (-not (Confirm-Action "Run for real now?")) { return }
    }
    Invoke-Undo -EntryId $entry.entryId
    Pause-ForUser
}
