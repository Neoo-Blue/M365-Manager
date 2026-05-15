# OneDrive handoff

Transfer a leaver's OneDrive to their manager (or a designated successor) before offboarding tears down the account.

## Why this matters

OneDrive content lives on the user's personal site (`/personal/<upn>`). When the user is deleted, the OneDrive enters a 90-day retention window — readable by an admin via SPO, but invisible to peers. Without a handoff, knowledge work created in OneDrive (drafts, scratch decks, vendor-shared files) becomes orphaned.

A handoff:

- Adds a target user as site owner on the OneDrive.
- Optionally extends retention beyond the default 90 days.
- Audits both actions with reverse recipes.

## Menu

**Slot 4 → Mailbox Archiving** *(despite the menu name, this also routes to OneDrive flows; the menu label is from before Phase 3.)*

Or directly:

```powershell
# Manual handoff
Grant-OneDriveAccess -LeaverUPN alice@contoso.com -TargetUPN bob@contoso.com

# Orchestrated handoff (used by the 12-step offboard)
Invoke-OneDriveHandoff -LeaverUPN alice@contoso.com -TargetUPN bob@contoso.com -RetentionDays 365
```

## Prereqs

- `SharePoint Administrator`.
- SPO admin URL configured (see [`sharepoint-management.md`](sharepoint-management.md)).
- The leaver still has an active OneDrive (function returns gracefully if not).

## Recent-files awareness

Before transferring, the tool surfaces what's in scope:

```
[OneDrive handoff for alice@contoso.com -> bob@contoso.com]

  Site URL    : https://contoso-my.sharepoint.com/personal/alice_contoso_com
  Storage     : 12.4 GB / 1 TB
  Last modified: 2026-05-13 (1 day ago)
  Recent files (last 7 days):
    - Q1-customers-draft.xlsx  (modified 2026-05-13)
    - vendor-contract-v2.docx  (modified 2026-05-12)
    - meeting-notes.txt         (modified 2026-05-10)
    ... (47 more)
```

Helpful context — "is there anything in here?" — before committing.

## What the orchestration does

`Invoke-OneDriveHandoff` runs three Invoke-Action steps:

1. **Grant access**: `Grant-OneDriveAccess`. `actionType=GrantOneDriveAccess`, `reverseType=RevokeOneDriveAccess`. Equivalent to `Add-SPOSiteOwner` on the personal site.
2. **Extend retention** (optional, `-RetentionDays N`): sets `RetentionEndDate` on the user object via Graph. Capped at tenant policy max. `noUndoReason` set because retention extension by a specific date isn't auto-reversible.
3. **Audit summary email** (optional): if Notifications is configured, sends `Send-OneDriveHandoffSummary` to the target user with the site URL + recent-files snapshot.

```
[+] Granted bob@contoso.com access to alice's OneDrive
[+] Retention extended to 2027-05-14
[+] Summary email sent to bob@contoso.com
[+] entryIds: a1b2c3d4-... (grant), b2c3d4e5-... (retention), c3d4e5f6-... (email)
```

## Revoking access

After the target has copied what they need, revoke:

```powershell
Revoke-OneDriveAccess -LeaverUPN alice@contoso.com -TargetUPN bob@contoso.com
```

Or `Invoke-Undo -EntryId <grant-entry-id>` to use the audit-driven reverse.

## During offboarding

The 12-step canonical offboard (see [`offboarding.md`](offboarding.md)) wraps this:

- Step 7: `Invoke-OneDriveHandoff` against the `HandoffOneDriveTo` UPN from the bulk CSV column.
- The handoff fires BEFORE step 12 (the actual user removal) so the target has access at the moment the leaver disappears.

## Bulk

For mass offboarding scenarios, the bulk offboard CSV's `HandoffOneDriveTo` column drives this per-row. See [`offboarding.md`](offboarding.md).

## Common failures

| Symptom | Cause + fix |
|---|---|
| `User has no OneDrive provisioned` | Common for guest accounts and accounts that never signed in. Skip the handoff step; no content to transfer. |
| `Add-SPOSiteOwner: Cannot find the URL` | The personal site URL changed (rare; happens after UPN rename). Re-resolve via `Get-SPOSite -Identity (Get-MgUser ...).userPrincipalName`. |
| `RetentionEndDate cannot exceed tenant policy max` | Your tenant's retention policy caps the date. Pick a sooner date or update the policy in Compliance Center. |
| `Email send failed` (handoff summary) | Notifications block not configured, or `DefaultEmailFrom` lacks send permission. Test with `Test-NotificationChannels`. |

## See also

- [`offboarding.md`](offboarding.md) — the 12-step flow that uses this.
- [`sharepoint-management.md`](sharepoint-management.md) — share lifecycle.
- [`notifications.md`](notifications.md) — `Send-OneDriveHandoffSummary` setup.
