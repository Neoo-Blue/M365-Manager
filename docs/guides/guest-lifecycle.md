# Guest user lifecycle

Discovery, recertification campaigns, and full teardown of B2B guest users.

## Menu

**Slot 19 → Guest Users**

```
  1. List all guests
  2. Stale-guest report (no sign-in for N days)
  3. Guests grouped by inviter / domain
  4. Recertification campaign (start / progress / apply decisions)
  5. Remove a guest (full teardown)
  6. Bulk guest removal from CSV
```

## Prereqs

- `User Administrator` or `Guest Inviter` + `User Administrator`.
- For full teardown: `Group Administrator` + `Teams Administrator` for cross-resource cleanup.

## Discovery

```powershell
Get-Guests | Format-Table

# Fields:
# UPN, DisplayName, Mail, Domains (external domains they're tied to),
# CreatedUtc, AgeDays, LastSignInUtc, DaysSinceSignIn, AccountEnabled, ExternalUserState
```

The 9999 sentinel for `DaysSinceSignIn` means "never signed in." Sort by `DaysSinceSignIn` descending for the stalest first.

## Stale guests

```powershell
Get-StaleGuests -DaysSinceSignIn 90
```

Default threshold: 90 days. Override via the parameter or the `StaleGuestDays` config key (resolved via `Get-EffectiveConfig` — per-tenant override supported).

Pivot views for triage:

```powershell
Get-GuestsByInviter            # who invited the most stale guests
Get-GuestsByDomain             # which external domains have the most exposure
```

## Recertification campaigns

For periodic guest review (often quarterly). The workflow:

1. **Start a campaign**: pulls stale guests + emails each guest's "manager" (the inviter, or a designated approver) asking yes/no.
2. **Operator collects replies** out of band (the tool doesn't have an inbox webhook).
3. **Apply decisions**: tool walks each pending record + the operator records yes/no/needs-more-info per row.

State file at `<stateDir>\guest-recerts.json` (per-tenant). Each record:

```jsonc
{
  "campaignId":    "recert-2026-Q1",
  "upn":           "external@partner.example",
  "invitedBy":     "bob@contoso.com",
  "decision":      null,        // null | "Keep" | "Remove" | "NeedsInfo"
  "decisionUtc":   null,
  "decidedBy":     null,
  "removedIncident": null       // populated when Remove-Guest ran
}
```

Apply decisions interactively:

```powershell
Show-PendingRecerts -CampaignId recert-2026-Q1
# Walks each pending row, prompts for decision, records.
```

When a decision is `Remove`, the tool offers to run the full teardown (see below).

## Full teardown

`Remove-Guest` (slot 19 → option 5) does four things in order:

1. **Revoke outbound shares** the guest created. Walks UAL, `SharingSet` operations by this user in the last 365 days, revokes each. Audited per share.
2. **Remove from groups**. Walks `/users/<id>/memberOf`, removes one group at a time. Each `Invoke-Action` with `reverseType=AddToGroup`.
3. **Remove from teams**. Uses `Invoke-TeamsOffboardTransfer` to handle sole-owner teams gracefully.
4. **Delete the user**. `DELETE /users/<id>`. Lands in `/directory/deletedItems` for 30 days; restore via `POST /directory/deletedItems/<id>/restore` if needed (the audit entry's `noUndoReason` documents this).

```powershell
Remove-Guest -UPN external@partner.example -Reason "Recert campaign: vendor decommissioned"
```

The `-Reason` is mandatory and lands in every audit entry.

## Bulk removal

For mass guest cleanup:

```powershell
Invoke-BulkGuestRemoval -Path guests-to-remove.csv
```

CSV columns: `UPN, Reason`. Sample at `templates/bulk-guest-removal-sample.csv`.

Validate-first / per-row / result CSV pattern from Phase 1.

## Common scenarios

### "We just acquired Vendor X's tenant and need to remove their guests"

```powershell
Get-Guests -Domain vendorx.com | Export-Csv .\vendorx-guests.csv -NoTypeInformation
# Review the CSV, add a Reason column with "M&A: vendor consolidated"
Invoke-BulkGuestRemoval -Path .\vendorx-guests.csv
```

### "Quarterly stale-guest cleanup"

```powershell
# Get stale guests
Get-StaleGuests -DaysSinceSignIn 180 | Export-Csv .\stale-q1.csv -NoTypeInformation
# Send a campaign for borderline cases
Start-RecertCampaign -CampaignId "recert-2026-Q1" -Path .\stale-q1.csv
# Wait for replies, then:
Show-PendingRecerts -CampaignId "recert-2026-Q1"
```

### "Compromised vendor account"

See [`incident-response.md`](incident-response.md) — the playbook handles compromised Guest users with the appropriate severity gate (Medium by default; Critical if AiTM signature).

## Common failures

| Symptom | Cause + fix |
|---|---|
| `User is not a guest` | The UPN belongs to a member user, not a guest. Verify via `Get-MgUser -UserId <upn>` — `userType` should be `Guest`. |
| `Cannot enumerate guest memberships: insufficient privileges` | The connecting account needs `Directory.Read.All`. Re-consent or use a more permissive account. |
| `Get-UserOutboundShares returns empty (but you know there are shares)` | UAL not enabled OR the guest pre-dates the audit retention window. Audit step is best-effort. |
| `Delete user: object not found` | The guest was already removed (e.g. concurrent operator). Audit entry is logged as `result=failure` with explanation. |

## See also

- [`offboarding.md`](offboarding.md) — equivalent flow for member users.
- [`incident-response.md`](incident-response.md) — compromised-guest playbook.
- [`sharepoint-management.md`](sharepoint-management.md) — share-revocation primitives.
- [`teams-management.md`](teams-management.md) — guest-as-team-member handling.
