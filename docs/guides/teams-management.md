# Teams management

Membership + ownership management for Microsoft Teams, with offboarding hooks for sole-owner handoff.

## Menu

**Slot 17 → Teams Management**

```
  1. List a user's teams
  2. Add a user to a team (member or owner)
  3. Remove a user from a team
  4. Promote / demote owner
  5. Orphan-team report (teams with zero owners)
  6. Single-owner classification report
  7. Bulk Teams membership from CSV
```

## Prereqs

- Graph scopes: `Team.ReadBasic.All`, `TeamMember.ReadWrite.All`, `Group.ReadWrite.All`.
- Role: `Teams Administrator` or higher.

## Listing a user's teams

```powershell
Get-UserTeams -UPN alice@contoso.com | Format-Table

# Output:
# TeamId    DisplayName            UserRole  OtherOwners
# t-001     Engineering            Owner     2
# t-002     Sales NA               Member    4
# t-003     Project Phoenix        Owner     0       <-- sole owner!
```

`OtherOwners=0` means Alice is the team's only owner. If she leaves without a handoff the team is orphaned and members lose ownership escalation.

## Add / remove / promote / demote

Each operation is its own `Invoke-Action` with a `reverse` recipe:

| Forward | Reverse | Notes |
|---|---|---|
| `AddToTeam` (member or owner) | `RemoveFromTeam` | Reverse drops back to no membership. |
| `RemoveFromTeam` | `AddToTeam` | Reverse re-adds at the original role (captured in target). |
| `PromoteTeamOwner` | `DemoteTeamOwner` | |
| `DemoteTeamOwner` | `PromoteTeamOwner` | |

The undo handlers (`Undo.ps1`) walk the team membership API via Graph DELETE / POST on `/groups/{id}/members/$ref` and `/groups/{id}/owners/$ref`.

## Orphan + single-owner reports

**Slot 17 → option 5** lists every team with zero owners. Adopt by promoting a member, or archive the team.

**Slot 17 → option 6** lists every team with exactly one owner. Useful before mass offboarding — you want to identify these BEFORE the owner leaves.

```
  SINGLE-OWNER TEAMS
  ------------------
  TeamId    DisplayName            Sole Owner               Members
  t-003     Project Phoenix        alice@contoso.com        14
  t-008     Internal QBR Q1        carol@contoso.com         5
  t-011     Vendor X integration   bob@contoso.com           3
```

## Offboarding hook

`Invoke-TeamsOffboardTransfer -LeaverUPN <upn>` (called by the 12-step offboard flow's step 10) walks every team the leaver owns or co-owns. For sole-owner teams, the operator is prompted to:

1. Pick a successor UPN (auto-promoted to owner before the leaver is removed).
2. Or accept that the team will be orphaned (audit entry warns).

For co-owner teams, the leaver is simply removed; remaining owners keep the team running.

```
[Teams handoff for alice@contoso.com]
  Team 'Project Phoenix' (sole owner)
    Promote successor? Enter UPN, or blank to orphan: dave@contoso.com
    [+] Promoted dave@contoso.com to owner.
    [+] Removed alice@contoso.com.

  Team 'Sales NA' (member only)
    [+] Removed alice@contoso.com.

  Team 'Engineering' (co-owner, 2 other owners)
    [+] Removed alice@contoso.com.
```

Every action audited; reversible per entry.

## Bulk membership CSV

**Slot 17 → option 7.**

CSV columns: `UPN, TeamId or TeamName, Action (Add|Remove|Promote|Demote), Role (Member|Owner — for Add)`.

```csv
UPN,TeamId,TeamName,Action,Role
alice@contoso.com,t-001,,Add,Owner
bob@contoso.com,,Engineering,Remove,
carol@contoso.com,,Engineering,Promote,
```

Either `TeamId` or `TeamName` must be set; the validator catches missing.

Bulk validates first (validates the team exists, the user exists, the action is one of `Add/Remove/Promote/Demote`); per-row failures don't halt the batch; result CSV written next to input.

## Common failures

| Symptom | Cause + fix |
|---|---|
| `Cannot add member to team: insufficient privileges` | The connecting account isn't a Teams Administrator or doesn't own the team. Adjust permissions or have the team owner add. |
| `Multiple teams match 'Sales'` | The `TeamName` lookup uses contains-match; refine the name or use `TeamId`. |
| `Failed to promote: user is not a member` | Promote requires the user be a member first. The flow auto-adds-then-promotes; if this fires, the auto-add step failed silently. |
| `Cannot remove the only owner` | Microsoft 365 doesn't allow demoting / removing the only owner. Promote a successor first (as the offboard hook does). |

Full troubleshooting at [`../operations/troubleshooting.md`](../operations/troubleshooting.md).

## See also

- [`offboarding.md`](offboarding.md) — step 10 handles Teams handoff automatically.
- [`audit-and-undo.md`](audit-and-undo.md) — every Teams action is reversible.
- [`../reference/csv-formats.md`](../reference/csv-formats.md) — bulk Teams CSV schema.
