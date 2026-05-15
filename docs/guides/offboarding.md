# Offboarding flow (12 steps)

Order matches Phase 3 spec. Each step is independently skippable via per-operator confirm in single-user mode, and via column-level flags in bulk mode.

```
+-------------------------------------------------------------+
|  USER OFFBOARDING -- canonical 12-step orchestration         |
+-------------------------------------------------------------+

   leaver UPN  +  optional context (forward to / OOO / etc.)
        |
        v
  +--------------+   prevent re-auth even with the device
  | 0  MFA       |   Remove-AllAuthMethods (MFAManager.ps1)
  +--------------+
        |
        v
  +--------------+   revoke sessions + AccountEnabled=$false
  | 1  Sign-in   |
  +--------------+
        |
        v
  +--------------+   set OOO + forward outbound mail
  | 2  Mail rules|
  +--------------+
        |
        v
  +--------------+   memberOf scan, filter: securityEnabled
  | 3  Sec grps  |     AND NOT mailEnabled AND NOT Unified
  +--------------+
        |
        v
  +--------------+   memberOf scan, filter: mailEnabled
  | 4  DLs       |     AND NOT team-backed (Step 7 handles those)
  +--------------+
        |
        v
  +--------------+   direct assignments only; group-assigned
  | 5  Licenses  |     SKUs left intact (remove from the group
  +--------------+     instead)
        |
        v
  +--------------+   Set-Mailbox -Type Shared (conditional)
  | 6  Shared mb |
  +--------------+
        |
        +---- (single-user only) 6b: grant delegates Full Access / SendAs
        |
        v
  +--------------+   sole owner: promote successor + remove
  | 7  Teams     |   co-owner: demote + remove
  +--------------+   member: remove
        |
        v
  +--------------+   UAL scan for last 365d of share events
  | 8  SP shares |     -> revoke each (Y/A/N confirm in single,
  +--------------+      auto in bulk if RevokeShares=true)
        |
        v
  +--------------+   Get-SPOSite owner filter -> grant successor
  | 9  OneDrive  |     SCA + retention extension intent +
  +--------------+     handoff email with top-20 recent files
        |
        v
  +--------------+   HTML recap of all step outcomes
  | 10 Mgr email |     (uses OneDrive notifier's Send-Offboard-
  +--------------+      ManagerSummary helper)
        |
        v
  +--------------+   one OFFBOARD_COMPLETE JSONL audit line
  | 11 Audit sum |     summarizing every step's outcome
  +--------------+
        |
        v
     done
```

## Bulk-mode column-to-step map

| CSV column           | Default | Affects step  |
|----------------------|---------|---------------|
| `UserPrincipalName`  | required | all |
| `ForwardTo`          | empty   | Step 2 (mail forwarding) |
| `ConvertToShared`    | false   | Step 6 |
| `HandoffOneDriveTo`  | empty   | Step 9 (UPN of OneDrive successor) |
| `RemoveFromAllGroups`| false   | Steps 3 + 4 (security groups + DLs) |
| `Reason`             | empty   | Step 2 (OOO message body) |
| `TeamsSuccessor`     | empty   | Step 7 (UPN for sole-owner Teams; prompts per team in single-user mode if blank) |
| `RevokeShares`       | true    | Step 8 |
| `NotifyManager`      | true    | Steps 9 + 10 |

Mandatory Phase 3 connectivity:

- Step 0  : Graph (`UserAuthenticationMethod.ReadWrite.All`)
- Step 1  : Graph
- Step 2  : Graph + EXO
- Steps 3 + 4 : Graph
- Step 5  : Graph
- Step 6  : EXO
- Step 7  : Graph (`TeamMember.ReadWrite.All` for $ref calls)
- Step 8  : EXO (`Search-UnifiedAuditLog`)
- Step 9  : SPO (admin role) + Graph (`Sites.FullControl.All`, `Mail.Send`)
- Step 10 : Graph (`Mail.Send`)
- Step 11 : local file write to `audit/`

Soft-fail rules: if SPO is not connectable (license / role missing), Steps 8 and 9 emit a warning and continue with the remaining steps. The result row records the failure in its `Reason` column.

## Reversibility map

| Step | Action type             | Reversible? | Inverse handler              |
|------|-------------------------|-------------|------------------------------|
| 0    | RevokeAuthMethod        | no          | (re-register manually)       |
| 1    | RevokeSignInSessions    | no          | n/a                          |
| 1    | BlockSignIn             | yes         | UnblockSignIn                |
| 2    | SetOOO                  | yes         | ClearOOO                     |
| 2    | SetForwarding           | yes         | ClearForwarding              |
| 3    | RemoveFromGroup         | yes         | AddToGroup                   |
| 4    | RemoveFromGroup         | yes         | AddToGroup                   |
| 5    | RemoveLicense           | yes         | AssignLicense                |
| 6    | ConvertToShared         | no          | re-license required          |
| 7    | RemoveFromTeam          | yes         | AddToTeam                    |
| 7    | DemoteTeamOwner         | yes         | PromoteTeamOwner             |
| 7    | PromoteTeamOwner        | yes         | DemoteTeamOwner              |
| 8    | RevokeExternalShare     | no          | recipient consent needed     |
| 9    | GrantOneDriveAccess     | yes         | RevokeOneDriveAccess         |
| 9    | ExtendOneDriveRetention | no          | intent-only record           |
| 10   | SendOffboardSummary     | no          | email sent                   |

See `docs/reference/audit-format.md` for the JSONL audit record shape and `Undo.ps1` for the full dispatch table.
