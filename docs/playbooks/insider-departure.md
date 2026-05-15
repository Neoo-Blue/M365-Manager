# Insider departure

A regular employee (or contractor, or vendor user) is leaving. Standard offboarding — but with the legal / compliance + knowledge-transfer steps the tool helps automate.

## When to use this playbook

- Voluntary resignation with notice period.
- End of contract / consulting engagement.
- Layoff (legal-sensitive — see variations).
- Vendor consolidation (vendor's user is going away).

For hostile departures (terminations / litigation hold / suspected exfiltration), see [`compromised-account.md`](compromised-account.md) and consider invoking it pre-emptively as if the account were compromised.

## Day-N preparation (>2 weeks before departure)

1. **Calendar the offboard** — exact effective time + day. Most companies cut access at end-of-day final-day-of-work.
2. **Identify the successor** for handoffs:
   - Manager (for OneDrive content + email forward).
   - Team mate for active customer relationships (CRM / email aliases).
   - Co-owner promotion targets for any Teams the user solely owns.
3. **Generate a manifest** so nothing is missed:
   ```powershell
   $upn = "alice@contoso.com"
   $manifest = @{
       Licenses    = (Get-MgUser -UserId $upn -Property AssignedLicenses).AssignedLicenses
       Groups      = (Get-MgUserMemberOf -UserId $upn).Value
       Teams       = (Get-UserTeams -UPN $upn)
       SharedMboxes = (Get-MailboxPermission -Identity ... | Where-Object User -eq $upn)
       OutboundShares = (Get-UserOutboundShares -UPN $upn -LookbackDays 180)
   }
   $manifest | ConvertTo-Json -Depth 8 | Set-Content alice-pre-offboard.json
   ```
   Useful both for confirming handoffs went where they should, and for a post-departure spot-check.

## 1-2 days before

1. **Identify Teams the user solely owns** (use `Slot 17 → option 6` from the menu):
   ```powershell
   Get-SingleOwnerTeams | Where-Object { $_.Owners[0] -eq "alice@contoso.com" }
   ```
   For each: pick a successor and pre-promote them so the offboard doesn't orphan the team.

2. **Brief the successor** on what they'll inherit:
   - OneDrive content (point them at the recent-files preview).
   - Forwarded mail (set expectations for 30-day forwarding window).
   - Specific shared mailboxes.

3. **Schedule the run** so notifications + audit logs are timestamped clearly. Off-hours is fine; the affected user isn't online anyway.

## Day-of (execution)

### Option A — single user, interactive

Menu **Slot 2 → Offboard User**, fill the prompts:

```
> UPN              : alice@contoso.com
> Reason            : Resigned, last day 2026-05-14
> ForwardTo         : bob@contoso.com    (manager)
> ConvertToShared   : Y                    (keeps mailbox accessible to bob)
> HandoffOneDriveTo : bob@contoso.com    (manager)
> RemoveFromAllGroups: Y
```

The tool runs the canonical 12-step flow. See [`../guides/offboarding.md`](../guides/offboarding.md) for the per-step detail.

### Option B — bulk (multiple departures, e.g. layoff)

CSV at `departures-2026-05-14.csv`:

```csv
UPN,ForwardTo,ConvertToShared,HandoffOneDriveTo,RemoveFromAllGroups,Reason
alice@contoso.com,bob@contoso.com,yes,bob@contoso.com,yes,RIF 2026-05
charlie@contoso.com,dave@contoso.com,yes,dave@contoso.com,yes,RIF 2026-05
```

Run in PREVIEW first:

```powershell
Invoke-BulkOffboard -Path .\departures-2026-05-14.csv -WhatIf
```

Review the PREVIEW audit output + the result CSV. Then run for real:

```powershell
Invoke-BulkOffboard -Path .\departures-2026-05-14.csv
```

Result CSV next to the input.

## Post-departure (within 24h)

1. **Verify handoffs landed** — open `bob@contoso.com`'s OneDrive, confirm they can see alice's content. Open Teams, confirm sole-owner promotions went through.
2. **Audit the audit log** — look for any `failure` results in the offboard run:
   ```powershell
   Read-AuditEntries | Where-Object { $_.target.userUpn -eq 'alice@contoso.com' -and $_.result -eq 'failure' }
   ```
3. **Send the manager a summary**. If `Notifications.ps1` is configured, the tool already sent `Send-OffboardManagerSummary`. Verify it landed.

## Day-15 + Day-30 + Day-90 checkpoints

| Day | Task |
|---|---|
| 15 | Verify forwarding is still routing to `bob`. Bob acknowledges receipt. |
| 30 | Remove the forwarding (if temporary). Verify the mailbox conversion to Shared persisted. |
| 90 | Decommission. Remove the Shared mailbox entirely (if no longer needed). Delete the user account. |

Step 90 (user deletion) is operator-driven — the offboard flow doesn't delete the account on day-of because deleted users go to `/directory/deletedItems` for 30 days, and you may need to restore for an unexpected handoff question.

To delete on day 90:

```powershell
Remove-MgUser -UserId alice@contoso.com -ErrorAction Stop
# Audit entry will be written with noUndoReason explaining the 30-day soft-delete window.
```

## Variations

### Layoff / RIF

Legal and HR are stakeholders. The offboard flow itself is the same — but the *coordination* changes:

- Don't run the offboard until HR has formally notified the affected employees.
- For very large layoffs, batch + stagger so the audit log is readable.
- Notifications are likely manual (HR-controlled emails) — set `Notifications.DryRunNotifications=true` for the bulk run to suppress automated alerts.

### Termination for cause + suspected exfiltration

Treat as a compromised account. Run [`compromised-account.md`](compromised-account.md) at `High` severity with `-QuarantineSentMail` if there's evidence of phishing-from-internal. Legal hold on the snapshot dir — DO NOT delete artifacts for the duration of any litigation.

### Vendor / contractor departure

Same flow but:
- Vendor users are usually Guests. Use the guest-removal flow from [`../guides/guest-lifecycle.md`](../guides/guest-lifecycle.md) instead of regular offboard.
- Don't forward mail to a successor — the vendor's mail doesn't live in your tenant.
- Audit outbound shares from the vendor (they may have shared files OUT to other vendors).

### Long-tenure executive

OneDrive content may be voluminous (decades-old). Pre-handoff:

1. Set `RetentionEndDate` extended to 3-5 years instead of the default 90 days.
2. Notify legal/compliance — exec content often has hold requirements.
3. Consider a content review BEFORE deletion, not after.

## See also

- [`../guides/offboarding.md`](../guides/offboarding.md) — the 12-step flow detail.
- [`../guides/guest-lifecycle.md`](../guides/guest-lifecycle.md) — for vendor / Guest user offboarding.
- [`../guides/teams-management.md`](../guides/teams-management.md) — sole-owner handoff mechanics.
- [`compromised-account.md`](compromised-account.md) — when departure is hostile.
