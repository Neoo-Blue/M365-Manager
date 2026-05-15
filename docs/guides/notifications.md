# Notifications

Email + Teams webhook dispatcher. Used by health checks, scheduled audits, the incident response playbook, and any operator-driven message that wants to reach the security or operations team.

## Setup

**Slot 20 → License & Cost → Notifications setup**

Or directly: `Start-NotificationsSetup`.

The wizard collects:

1. **DefaultEmailFrom** — UPN to send from. Leave blank to use the operator's mailbox via Graph `/me/sendMail`.
2. **SecurityTeamRecipients** — emails routed for `Critical` severity.
3. **OperationsTeamRecipients** — emails routed for `Warning` / `Info`.
4. **TeamsWebhookSecurity** — incoming-webhook URL for the security channel. Encrypted DPAPI on save.
5. **TeamsWebhookOperations** — same for operations channel.
6. **DryRunNotifications** — boolean. When true, every send writes an audit line but does NOT actually deliver. Useful for testing health checks without spamming the team.

State persists to `ai_config.json` under the `Notifications` block.

## Sending

```powershell
Send-Notification -Channels SecurityTeam -Severity Critical `
    -Subject "Unusual sign-in detected for alice@contoso.com" `
    -Body "<html>...</html>"
```

`-Channels` is one of:

- `SecurityTeam` — email recipients + Teams security webhook.
- `OperationsTeam` — email recipients + Teams operations webhook.
- `Both` — fire both.

`-Severity`: `Critical` (red), `Warning` (yellow), `Info` (gray) — drives the Teams adaptive card theme.

## Atomic send fns

For specific use cases (e.g. offboarding handoff summaries, guest recertification campaigns), the tool ships helper functions that build the body and call `Send-Notification` internally:

| Helper | Body |
|---|---|
| `Send-OneDriveHandoffSummary` | OneDrive site URL + recent files + retention end date. |
| `Send-OffboardManagerSummary` | Steps that ran on the leaver's offboard. |
| `Send-GuestRecertEmail` | Per-guest recertification HTML with yes/no instructions. |

All three delegate to `Send-Email` when `Notifications.ps1` is loaded; if not loaded, they fall back to `/me/sendMail` directly (so older deploys without notifications configured still get the same emails — they just don't honor DryRun).

## Test channels

```powershell
Test-NotificationChannels
# Sends a low-severity test to each configured channel.
# Returns @{ Email=$true/$false; Teams=$true/$false } per channel.
```

Run this after first setup + after any DefaultEmailFrom / webhook URL change.

## DryRun mode

`Notifications.DryRunNotifications=true` writes one audit line per send (`actionType=Notify`, `result=preview`) without actually delivering. Pair with scheduled health checks during initial setup so the team doesn't get pages until the routing is verified.

## Security model

- **Webhook URLs are sensitive.** Anyone with the URL can post to that Teams channel. Encrypted at rest with the `DPAPI:` prefix.
- **SMTP password (if using SMTP)** — DPAPI-encrypted same way.
- **Send-Email uses Graph `/me/sendMail` by default** — so the message is sent FROM the operator's mailbox unless `DefaultEmailFrom` overrides. Operators with shared "noreply" mailboxes can set `DefaultEmailFrom` to that mailbox and grant Send-As permission.

## Used by

- Phase 4 health checks: `health-mfa-gaps.ps1`, `health-stale-guests.ps1`, etc. all call `Send-Notification` on findings.
- Phase 7 incident response: step 12 routes to SecurityTeam Critical.
- Phase 7 detector framework: every finding alerts the team before any escalation decision.

## Common failures

| Symptom | Cause + fix |
|---|---|
| `Send-MgUserMail: Forbidden` | The operator's mailbox doesn't permit Graph `/me/sendMail`. Set `DefaultEmailFrom` to a mailbox you have Send-As on. |
| `Teams webhook returned 410 Gone` | The webhook URL was deleted in the Teams channel. Regenerate + re-save. |
| `Recipients list is empty` | The Notifications block has empty `SecurityTeamRecipients`. Re-run `Start-NotificationsSetup`. |
| `Test-NotificationChannels returns $false for Email but Teams works` | Likely the `DefaultEmailFrom` mailbox lacks send permission. Use the operator's mailbox (leave blank) or grant Send-As. |

## See also

- [`scheduled-checks.md`](scheduled-checks.md) — the most common notification source.
- [`incident-response.md`](incident-response.md) — step 12 (notify) + the auto-detection framework's per-finding alerts.
- [`../getting-started/configuration.md`](../getting-started/configuration.md) — `Notifications` block reference.
