# ============================================================
#  TeamsManager.ps1 — Teams membership / ownership management
#
#  Microsoft Graph endpoints used:
#    GET  /users/{id}/joinedTeams                  -- list teams
#    GET  /groups/{id}/owners                      -- owner roster
#    GET  /groups/{id}/members
#    POST /groups/{id}/owners/$ref                 -- promote
#    POST /groups/{id}/members/$ref                -- add member
#    DELETE /groups/{id}/owners/{userId}/$ref      -- demote
#    DELETE /groups/{id}/members/{userId}/$ref     -- remove
#    GET  /teams                                   -- tenant scan
#
#  Required scopes (already requested in Auth.ps1 Phase 3):
#    Group.ReadWrite.All, TeamMember.ReadWrite.All,
#    TeamSettings.ReadWrite.All
# ============================================================

function ConvertTo-TeamRecord {
    param($Raw, [string]$Role)
    [PSCustomObject]@{
        TeamId      = [string]$Raw.id
        DisplayName = [string]$Raw.displayName
        Description = [string]$Raw.description
        Visibility  = [string]$Raw.visibility
        Role        = $Role
    }
}

function Get-UserTeams {
    <#
        Joined teams for a user with the user's role on each
        (Member or Owner). Role is computed by checking owners
        on each team (joinedTeams alone doesn't carry the role).
    #>
    param([Parameter(Mandatory)][string]$UPN)
    try {
        $teams = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN/joinedTeams" -ErrorAction Stop).value)
    } catch {
        Write-ErrorMsg "Could not list teams for $UPN -- $($_.Exception.Message)"
        return @()
    }
    # Resolve user id once
    $userId = $null
    try { $userId = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN?`$select=id" -ErrorAction Stop).id } catch {}
    $out = New-Object System.Collections.ArrayList
    foreach ($t in $teams) {
        $role = 'Member'
        if ($userId) {
            try {
                $owners = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($t.id)/owners?`$select=id" -ErrorAction Stop).value
                if ($owners | Where-Object { [string]$_.id -eq $userId }) { $role = 'Owner' }
            } catch {}
        }
        [void]$out.Add((ConvertTo-TeamRecord -Raw $t -Role $role))
    }
    return @($out)
}

function Get-UserOwnedTeams {
    param([Parameter(Mandatory)][string]$UPN)
    return @(Get-UserTeams -UPN $UPN | Where-Object Role -eq 'Owner')
}

function Resolve-TeamIdentifier {
    <#
        Accepts either a TeamId (GUID) or a display name. If the
        name is ambiguous, lets the operator pick. Returns the
        team id string or $null.
    #>
    param([Parameter(Mandatory)][string]$IdOrName)
    if ($IdOrName -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        return $IdOrName
    }
    try {
        $escaped = $IdOrName -replace "'", "''"
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$escaped' and resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,description&`$count=true" -Headers @{ ConsistencyLevel = 'eventual' } -ErrorAction Stop
        $candidates = @($resp.value)
        if ($candidates.Count -eq 0) {
            Write-Warn "No team named '$IdOrName' found."
            return $null
        }
        if ($candidates.Count -eq 1) { return [string]$candidates[0].id }
        $labels = $candidates | ForEach-Object { "$($_.displayName)  ($($_.id))" }
        $sel = Show-Menu -Title "Multiple teams match '$IdOrName'" -Options $labels -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }
        return [string]$candidates[$sel].id
    } catch {
        Write-ErrorMsg "Team lookup failed: $($_.Exception.Message)"
        return $null
    }
}

function Add-UserToTeam {
    <#
        Adds a user to a team as Member or Owner. Wrapped in
        Invoke-Action with the inverse recipe RemoveFromTeam.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][string]$TeamId,
        [switch]$AsOwner
    )
    $userId = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN?`$select=id" -ErrorAction Stop).id
    $segment = if ($AsOwner) { 'owners' } else { 'members' }
    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$userId" } | ConvertTo-Json -Compress

    return Invoke-Action `
        -Description ("Add {0} to Team {1} as {2}" -f $UPN, $TeamId, $(if ($AsOwner) {'Owner'} else {'Member'})) `
        -ActionType 'AddToTeam' `
        -Target @{ userId = [string]$userId; userUpn = $UPN; teamId = $TeamId; role = $(if ($AsOwner) {'Owner'} else {'Member'}) } `
        -ReverseType 'RemoveFromTeam' `
        -ReverseDescription ("Remove {0} from Team {1}" -f $UPN, $TeamId) `
        -Action {
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/$segment/`$ref" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                $true
            } catch {
                if ($_.Exception.Message -match 'object references already exist|already a member|already exists') { 'already' } else { throw }
            }
        }
}

function Remove-UserFromTeam {
    param(
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][string]$TeamId,
        [switch]$AsOwner
    )
    $userId = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN?`$select=id" -ErrorAction Stop).id
    $segment = if ($AsOwner) { 'owners' } else { 'members' }
    return Invoke-Action `
        -Description ("Remove {0} from Team {1} ({2})" -f $UPN, $TeamId, $segment) `
        -ActionType 'RemoveFromTeam' `
        -Target @{ userId = [string]$userId; userUpn = $UPN; teamId = $TeamId; role = $(if ($AsOwner) {'Owner'} else {'Member'}) } `
        -ReverseType 'AddToTeam' `
        -ReverseDescription ("Re-add {0} to Team {1}" -f $UPN, $TeamId) `
        -Action {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/$segment/$userId/`$ref" -ErrorAction Stop | Out-Null
            $true
        }
}

function Set-TeamOwnership {
    <#
        Promote a member to owner, or demote an owner to member.
        Reverse pair: PromoteTeamOwner <-> DemoteTeamOwner.
        Promotion: POST /owners/$ref (idempotent if already an owner).
        Demotion: DELETE /owners/{userId}/$ref (leaves the member
                  membership intact).
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][string]$TeamId,
        [Parameter(Mandatory)][ValidateSet('Promote','Demote')][string]$Direction
    )
    $userId = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN?`$select=id" -ErrorAction Stop).id
    if ($Direction -eq 'Promote') {
        $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$userId" } | ConvertTo-Json -Compress
        return Invoke-Action `
            -Description ("Promote {0} to Owner of Team {1}" -f $UPN, $TeamId) `
            -ActionType 'PromoteTeamOwner' `
            -Target @{ userId = [string]$userId; userUpn = $UPN; teamId = $TeamId } `
            -ReverseType 'DemoteTeamOwner' `
            -ReverseDescription ("Demote {0} from Owner of Team {1}" -f $UPN, $TeamId) `
            -Action {
                try { Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/owners/`$ref" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null; $true }
                catch { if ($_.Exception.Message -match 'already exist') { 'already' } else { throw } }
            }
    } else {
        return Invoke-Action `
            -Description ("Demote {0} from Owner of Team {1}" -f $UPN, $TeamId) `
            -ActionType 'DemoteTeamOwner' `
            -Target @{ userId = [string]$userId; userUpn = $UPN; teamId = $TeamId } `
            -ReverseType 'PromoteTeamOwner' `
            -ReverseDescription ("Re-promote {0} to Owner of Team {1}" -f $UPN, $TeamId) `
            -Action { Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$TeamId/owners/$userId/`$ref" -ErrorAction Stop | Out-Null; $true }
    }
}

# ============================================================
#  Tenant-wide reports
# ============================================================

function Get-AllTeams {
    <#
        Paginates /teams. Returns minimal field set so the report
        helpers can layer their own enrichment.
    #>
    $out = New-Object System.Collections.ArrayList
    $uri = "https://graph.microsoft.com/v1.0/teams?`$select=id,displayName,description,visibility&`$top=500"
    try {
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($t in $resp.value) { [void]$out.Add($t) }
            $uri = $resp.'@odata.nextLink'
        } while ($uri)
    } catch { Write-ErrorMsg "Could not enumerate Teams: $($_.Exception.Message)"; return @() }
    return @($out)
}

function Get-OrphanedTeams {
    <#
        Teams with zero owners. Heavy: one /owners call per team.
        Operators are expected to run this off-hours on big tenants.
    #>
    Write-InfoMsg "Scanning tenant Teams for zero-owner state..."
    $teams = Get-AllTeams
    $hits = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($t in $teams) {
        $i++
        Write-Progress -Activity "Orphan scan" -Status $t.displayName -PercentComplete (($i / [Math]::Max(1, $teams.Count)) * 100)
        try {
            $owners = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($t.id)/owners?`$select=id" -ErrorAction Stop).value
            if (-not $owners -or @($owners).Count -eq 0) {
                [void]$hits.Add([PSCustomObject]@{ TeamId = $t.id; DisplayName = $t.displayName; Visibility = $t.visibility; OwnerCount = 0 })
            }
        } catch {}
    }
    Write-Progress -Activity "Orphan scan" -Completed
    return @($hits | Sort-Object DisplayName)
}

function Get-SingleOwnerTeams {
    Write-InfoMsg "Scanning tenant Teams for single-owner SPOF risk..."
    $teams = Get-AllTeams
    $hits = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($t in $teams) {
        $i++
        Write-Progress -Activity "Single-owner scan" -Status $t.displayName -PercentComplete (($i / [Math]::Max(1, $teams.Count)) * 100)
        try {
            $owners = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($t.id)/owners?`$select=id,userPrincipalName,displayName" -ErrorAction Stop).value)
            if ($owners.Count -eq 1) {
                $o = $owners[0]
                [void]$hits.Add([PSCustomObject]@{ TeamId = $t.id; DisplayName = $t.displayName; OwnerUPN = $o.userPrincipalName; OwnerName = $o.displayName })
            }
        } catch {}
    }
    Write-Progress -Activity "Single-owner scan" -Completed
    return @($hits | Sort-Object DisplayName)
}

function Get-TeamsWithGuests {
    Write-InfoMsg "Scanning tenant Teams for guest membership..."
    $teams = Get-AllTeams
    $hits = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($t in $teams) {
        $i++
        Write-Progress -Activity "Guest scan" -Status $t.displayName -PercentComplete (($i / [Math]::Max(1, $teams.Count)) * 100)
        try {
            $members = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($t.id)/members?`$select=id,userType,userPrincipalName" -ErrorAction Stop).value)
            $guests = @($members | Where-Object { $_.userType -eq 'Guest' })
            if ($guests.Count -gt 0) {
                [void]$hits.Add([PSCustomObject]@{ TeamId = $t.id; DisplayName = $t.displayName; GuestCount = $guests.Count; Guests = (($guests | ForEach-Object { $_.userPrincipalName }) -join '; ') })
            }
        } catch {}
    }
    Write-Progress -Activity "Guest scan" -Completed
    return @($hits | Sort-Object -Property GuestCount -Descending)
}

# ============================================================
#  Bulk
# ============================================================

function Invoke-BulkTeamsMembership {
    <#
        CSV columns: UPN, TeamId or TeamName, Action (Add|Remove|
        Promote|Demote), Role (Member|Owner, only relevant for Add).
        Validate-then-execute pattern (cf. BulkOnboard).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [switch]$WhatIf)
    if (-not (Test-Path -LiteralPath $Path)) { Write-ErrorMsg "CSV not found: $Path"; return }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { Write-Warn "Empty CSV."; return }

    if (-not (Connect-ForTask 'Teams')) { Write-ErrorMsg "Could not connect."; return }

    $errors = @()
    $normalized = @()
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]; $row = $i + 2
        $upn = [string]$r.UPN
        $idOrName = if ($r.TeamId) { [string]$r.TeamId } elseif ($r.TeamName) { [string]$r.TeamName } else { '' }
        $act = [string]$r.Action
        $role = if ($r.Role) { [string]$r.Role } else { 'Member' }
        # ${row} delimiter is required -- "$row:" parses as a drive
        # reference and breaks the whole script at tokenize time.
        if ([string]::IsNullOrWhiteSpace($upn))      { $errors += "Row ${row}: missing UPN" }
        if ([string]::IsNullOrWhiteSpace($idOrName)) { $errors += "Row ${row}: missing TeamId/TeamName" }
        if ($act -notin @('Add','Remove','Promote','Demote')) { $errors += "Row ${row}: invalid Action '$act'" }
        $normalized += [PSCustomObject]@{ UPN = $upn; IdOrName = $idOrName; Action = $act; Role = $role; RowNum = $row }
    }
    if ($errors.Count -gt 0) {
        $errors | ForEach-Object { Write-ErrorMsg $_ }
        return
    }
    Write-Success "Validation passed: $($normalized.Count) row(s)."

    $previousMode = Get-PreviewMode
    if ($WhatIf.IsPresent -and -not $previousMode) { Set-PreviewMode -Enabled $true }
    try {
        $results = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $normalized.Count; $i++) {
            $r = $normalized[$i]
            Write-Progress -Activity "Bulk Teams" -Status "$($r.UPN) -- $($r.Action)" -PercentComplete (($i / $normalized.Count) * 100)
            $teamId = Resolve-TeamIdentifier -IdOrName $r.IdOrName
            $entry = [ordered]@{ UPN = $r.UPN; TeamRequested = $r.IdOrName; TeamId = $teamId; Action = $r.Action; Role = $r.Role; Status = ''; Reason = '' }
            if (-not $teamId) { $entry.Status = 'Failed'; $entry.Reason = 'team not found'; [void]$results.Add([PSCustomObject]$entry); continue }
            try {
                switch ($r.Action) {
                    'Add'     { Add-UserToTeam -UPN $r.UPN -TeamId $teamId -AsOwner:($r.Role -eq 'Owner') | Out-Null }
                    'Remove'  { Remove-UserFromTeam -UPN $r.UPN -TeamId $teamId -AsOwner:($r.Role -eq 'Owner') | Out-Null }
                    'Promote' { Set-TeamOwnership -UPN $r.UPN -TeamId $teamId -Direction Promote | Out-Null }
                    'Demote'  { Set-TeamOwnership -UPN $r.UPN -TeamId $teamId -Direction Demote  | Out-Null }
                }
                $entry.Status = if (Get-PreviewMode) { 'Preview' } else { 'Success' }
            } catch { $entry.Status = 'Failed'; $entry.Reason = $_.Exception.Message }
            [void]$results.Add([PSCustomObject]$entry)
        }
        Write-Progress -Activity "Bulk Teams" -Completed
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $out = Join-Path (Split-Path -Parent (Resolve-Path $Path)) ("bulk-teams-$stamp.csv")
        $results | Export-Csv -LiteralPath $out -NoTypeInformation -Force
        Write-Success "Result CSV: $out"
    } finally {
        Set-PreviewMode -Enabled $previousMode
    }
}

# ============================================================
#  Offboard integration -- ownership transfer
# ============================================================

function Invoke-TeamsOffboardTransfer {
    <#
        For an offboarding leaver:
          - every team where they are the SOLE owner: prompt for
            a successor UPN, promote them, then remove the leaver
          - every team where they are a NON-sole owner: demote
            (other owners remain)
          - every team where they are just a member: remove
        Returns a summary record for the offboard result CSV.
    #>
    param(
        [Parameter(Mandatory)][string]$LeaverUPN,
        [string]$TeamsSuccessorUPN
    )
    if (-not (Connect-ForTask 'Teams')) { return [PSCustomObject]@{ Note = 'Teams connect failed' } }

    $owned = Get-UserOwnedTeams -UPN $LeaverUPN
    $all   = Get-UserTeams -UPN $LeaverUPN
    $memberOnly = @($all | Where-Object Role -eq 'Member')

    $soleOwnerActions = 0; $coOwnerActions = 0; $memberActions = 0; $failures = 0
    foreach ($t in $owned) {
        $owners = @()
        try { $owners = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($t.TeamId)/owners?`$select=id,userPrincipalName" -ErrorAction Stop).value) } catch {}
        $isSole = (@($owners).Count -le 1)

        if ($isSole) {
            $successor = $TeamsSuccessorUPN
            if (-not $successor) { $successor = Read-UserInput "New owner for sole-owner team '$($t.DisplayName)' (UPN, blank to skip)" }
            if ([string]::IsNullOrWhiteSpace($successor)) { continue }
            try {
                Set-TeamOwnership -UPN $successor.Trim() -TeamId $t.TeamId -Direction Promote | Out-Null
                Remove-UserFromTeam -UPN $LeaverUPN -TeamId $t.TeamId -AsOwner | Out-Null
                Remove-UserFromTeam -UPN $LeaverUPN -TeamId $t.TeamId          | Out-Null
                $soleOwnerActions++
            } catch { $failures++ }
        } else {
            try {
                Set-TeamOwnership -UPN $LeaverUPN -TeamId $t.TeamId -Direction Demote | Out-Null
                Remove-UserFromTeam -UPN $LeaverUPN -TeamId $t.TeamId | Out-Null
                $coOwnerActions++
            } catch { $failures++ }
        }
    }
    foreach ($t in $memberOnly) {
        try { Remove-UserFromTeam -UPN $LeaverUPN -TeamId $t.TeamId | Out-Null; $memberActions++ } catch { $failures++ }
    }
    return [PSCustomObject]@{
        LeaverUPN        = $LeaverUPN
        SoleOwnerActions = $soleOwnerActions
        CoOwnerActions   = $coOwnerActions
        MemberRemovals   = $memberActions
        Failures         = $failures
    }
}

# ============================================================
#  Menu
# ============================================================

function Start-TeamsMenu {
    while ($true) {
        $sel = Show-Menu -Title "Teams Management" -Options @(
            "View user's teams",
            "Add user to team...",
            "Remove user from team...",
            "Promote / demote owner...",
            "Reports: orphaned teams (zero owners)",
            "Reports: single-owner teams (SPOF)",
            "Reports: teams with guests",
            "Bulk Teams membership from CSV..."
        ) -BackLabel "Back"
        switch ($sel) {
            0 { $u = Read-UserInput "User UPN"; if ($u) { Get-UserTeams -UPN $u | Format-Table -AutoSize; Pause-ForUser } }
            1 {
                $u = Read-UserInput "User UPN"; if (-not $u) { continue }
                $t = Read-UserInput "Team id or name"; if (-not $t) { continue }
                $tid = Resolve-TeamIdentifier -IdOrName $t; if (-not $tid) { Pause-ForUser; continue }
                $asOwner = (Show-Menu -Title "Role" -Options @("Member","Owner") -BackLabel "Cancel") -eq 1
                Add-UserToTeam -UPN $u -TeamId $tid -AsOwner:$asOwner | Out-Null
                Pause-ForUser
            }
            2 {
                $u = Read-UserInput "User UPN"; if (-not $u) { continue }
                $t = Read-UserInput "Team id or name"; if (-not $t) { continue }
                $tid = Resolve-TeamIdentifier -IdOrName $t; if (-not $tid) { Pause-ForUser; continue }
                Remove-UserFromTeam -UPN $u -TeamId $tid | Out-Null
                Pause-ForUser
            }
            3 {
                $u = Read-UserInput "User UPN"; if (-not $u) { continue }
                $t = Read-UserInput "Team id or name"; if (-not $t) { continue }
                $tid = Resolve-TeamIdentifier -IdOrName $t; if (-not $tid) { Pause-ForUser; continue }
                $dirSel = Show-Menu -Title "Direction" -Options @("Promote to owner","Demote from owner") -BackLabel "Cancel"
                if ($dirSel -eq -1) { continue }
                $dir = if ($dirSel -eq 0) { 'Promote' } else { 'Demote' }
                Set-TeamOwnership -UPN $u -TeamId $tid -Direction $dir | Out-Null
                Pause-ForUser
            }
            4 { $rows = Get-OrphanedTeams;     $rows | Format-Table -AutoSize; if ($rows.Count -gt 0 -and (Confirm-Action "Export CSV?")) { $p = Join-Path (Get-AuditLogDirectory) ("teams-orphan-$(Get-Date -Format yyyyMMdd-HHmmss).csv"); $rows | Export-Csv -LiteralPath $p -NoTypeInformation -Force; Write-Success "CSV: $p" } ; Pause-ForUser }
            5 { $rows = Get-SingleOwnerTeams;  $rows | Format-Table -AutoSize; if ($rows.Count -gt 0 -and (Confirm-Action "Export CSV?")) { $p = Join-Path (Get-AuditLogDirectory) ("teams-single-$(Get-Date -Format yyyyMMdd-HHmmss).csv"); $rows | Export-Csv -LiteralPath $p -NoTypeInformation -Force; Write-Success "CSV: $p" } ; Pause-ForUser }
            6 { $rows = Get-TeamsWithGuests;   $rows | Format-Table -AutoSize; if ($rows.Count -gt 0 -and (Confirm-Action "Export CSV?")) { $p = Join-Path (Get-AuditLogDirectory) ("teams-guests-$(Get-Date -Format yyyyMMdd-HHmmss).csv"); $rows | Export-Csv -LiteralPath $p -NoTypeInformation -Force; Write-Success "CSV: $p" } ; Pause-ForUser }
            7 {
                $p = Read-UserInput "Path to CSV (UPN, TeamId/TeamName, Action, Role)"
                if (-not $p) { continue }
                $dry = Confirm-Action "Run as DRY-RUN first?"
                Invoke-BulkTeamsMembership -Path $p.Trim('"').Trim("'") -WhatIf:$dry
                Pause-ForUser
            }
            -1 { return }
        }
    }
}
