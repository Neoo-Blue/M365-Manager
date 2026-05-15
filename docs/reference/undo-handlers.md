# Undo handler dispatch table

Every reverse type registered in `$script:UndoHandlers` (`Undo.ps1`). The audit log's `reverse.type` field must match one of these for `Invoke-Undo -EntryId X` to dispatch successfully.

## How dispatch works

1. `Invoke-Undo -EntryId X`:
   - Reads the audit entry by id.
   - Pulls `entry.reverse.type` + `entry.reverse.target`.
   - Looks up `$script:UndoHandlers[type]` (a scriptblock).
   - Invokes it with `-Target <hashtable>`.
2. The handler returns truthy on success; the undo machinery writes its own audit entry (`event=UNDO`, `actionType=Undo-<original>`).
3. The original entry's id gets stamped in `<stateDir>\undo-state.json` so it can't be reversed twice.

## Identity

| Reverse type | Forward emitter | Notes |
|---|---|---|
| `UnblockSignIn` | `BlockSignIn` (any flow) | `PATCH /users/<id>` with `accountEnabled=true`. |
| `BlockSignIn` | `UnblockSignIn` | Mirror of above. |
| `RemoveLicense` | `AssignLicense` (any flow) | `Set-MgUserLicense -RemoveLicenses @(@{ SkuId = <id> })`. |
| `AssignLicense` | `RemoveLicense` | `-AddLicenses @(@{ SkuId = <id> })`. |

## Groups + DLs

| Reverse type | Forward emitter | Notes |
|---|---|---|
| `RemoveFromGroup` | `AddToGroup` | `DELETE /groups/<g>/members/<u>/$ref`. |
| `AddToGroup` | `RemoveFromGroup` | `POST /groups/<g>/members/$ref`. |
| `RemoveFromDistributionList` | `AddToDistributionList` | EXO `Remove-DistributionGroupMember`. |
| `AddToDistributionList` | `RemoveFromDistributionList` | EXO `Add-DistributionGroupMember`. |

## Mailbox permissions

| Reverse type | Forward emitter | Notes |
|---|---|---|
| `RevokeMailboxFullAccess` | `GrantMailboxFullAccess` | `Remove-MailboxPermission ... -AccessRights FullAccess`. |
| `GrantMailboxFullAccess` | `RevokeMailboxFullAccess` | `Add-MailboxPermission ... -AccessRights FullAccess`. |
| `RevokeMailboxSendAs` | `GrantMailboxSendAs` | `Remove-RecipientPermission ... -AccessRights SendAs`. |
| `GrantMailboxSendAs` | `RevokeMailboxSendAs` | `Add-RecipientPermission ... -AccessRights SendAs`. |

## Calendar permissions

| Reverse type | Forward emitter | Notes |
|---|---|---|
| `RevokeCalendarAccess` | `GrantCalendarAccess` | `Remove-MailboxFolderPermission`. |
| `GrantCalendarAccess` | `RevokeCalendarAccess` | `Add-MailboxFolderPermission`. |

## Mailbox state

| Reverse type | Forward emitter | Notes |
|---|---|---|
| `ClearOOO` | `SetOOO` (offboard flow) | `Set-MailboxAutoReplyConfiguration -AutoReplyState Disabled`. |
| `ClearForwarding` | `SetForwarding` | `Set-Mailbox -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false`. |
| `SetForwarding` | `ClearForwarding` | Restores prior forwarding from the audit entry's target. |

## OneDrive + SharePoint

| Reverse type | Forward emitter | Notes |
|---|---|---|
| `RevokeOneDriveAccess` | `GrantOneDriveAccess` | `Remove-SPOSiteOwner` on the personal site. |
| `GrantOneDriveAccess` | `RevokeOneDriveAccess` | `Add-SPOSiteOwner`. |
| `RemoveSiteOwner` | `AddSiteOwner` | `Set-SPOUser -RemoveCollectionAdministrator`. |
| `AddSiteOwner` | `RemoveSiteOwner` | `Set-SPOUser -IsSiteCollectionAdmin $true`. |

## Teams

| Reverse type | Forward emitter | Notes |
|---|---|---|
| `RemoveFromTeam` | `AddToTeam` | `DELETE /groups/<t>/members/<u>/$ref`. |
| `AddToTeam` | `RemoveFromTeam` | `POST /groups/<t>/members/$ref` or `/owners/$ref`. |
| `DemoteTeamOwner` | `PromoteTeamOwner` | Removes owner role. |
| `PromoteTeamOwner` | `DemoteTeamOwner` | Adds owner role. |

## Phase 7 incident response

| Reverse type | Forward emitter | Notes |
|---|---|---|
| `EnableInboxRule` | `Incident:DisableInboxRule` | `PATCH ...messageRules/<id>` with `isEnabled=true`. |
| `Incident:DisableInboxRule` | (Phase 7 step 6) | Forward emitter; reverse is `EnableInboxRule`. |

Phase 7's "snapshot + contain + audit" steps that are NOT reversible:
- `Incident:RevokeSessions` — sessions are gone.
- `Incident:RevokeAuthMethods` — must re-enroll.
- `Incident:ForcePasswordChange` — can't restore the prior password.
- `Incident:QuarantineSentMail` — compliance purge is permanent.

These all populate `noUndoReason` explicitly so `Invoke-Undo` prints the why and exits.

## Adding a handler

For a new reversible mutation:

1. Implement the inverse logic as a scriptblock in `Undo.ps1`:

   ```powershell
   $script:UndoHandlers['MyNewReverse'] = {
       param([hashtable]$Target)
       # use $Target.userId, $Target.someField etc.
       Invoke-MgGraphRequest -Method DELETE -Uri "..." -ErrorAction Stop
       $true
   }
   ```

2. Emit the matching `-ReverseType` from your forward `Invoke-Action`:

   ```powershell
   Invoke-Action -Description "..." -ActionType 'MyForward' `
       -Target @{ userId = $id; someField = $val } `
       -ReverseType 'MyNewReverse' `
       -ReverseDescription "Undo MyForward for $id" `
       -Action { ... }
   ```

3. Add a Pester test in `tests/Undo.Tests.ps1`. The existing pattern uses a mocked Graph + verifies the handler runs the right cmdlet.

The full list of registered handlers is 24 today. See `Undo.ps1` for the canonical definitions.

## See also

- [`audit-format.md`](audit-format.md) — the `reverse` recipe shape in the audit log.
- [`../guides/audit-and-undo.md`](../guides/audit-and-undo.md) — operator walkthrough.
- [`../developer/adding-a-module.md`](../developer/adding-a-module.md) — full pattern for a new reversible mutation.
