# Mass phishing response

A coordinated phishing campaign hit multiple accounts simultaneously. Multiple users clicked. You need to contain, audit, and notify in parallel — not one at a time.

## When to use this playbook

- Email security gateway flags multiple accounts received the same phishing message.
- 2+ users report receiving phishing within the same window.
- Detector framework surfaces `SuspiciousInboxRule` on 2+ accounts within a short interval.
- A user reports a teammate's account just sent them a phishing message.

## Decision tree

```
                  +--------------------------+
                  |  Phishing reported       |
                  +--------------------------+
                              |
                  +-----------v------------+
                  | How many accounts hit? |
                  +-----------+------------+
                              |
              +---------------+---------------+
              |                               |
        +-----v-----+                    +----v-----+
        |   = 1     |                    |   >= 2   |
        +-----+-----+                    +----+-----+
              |                               |
+-------------v--------------+   +------------v---------------+
| compromised-account.md     |   | THIS PLAYBOOK (bulk + IR)   |
| (single Invoke-Compromised |   | (Invoke-BulkIncidentResp,   |
| AccountResponse)           |   |  cross-tenant if MSP)       |
+----------------------------+   +-----------------------------+
```

## Step 1 — Stop the spread (minutes 0-5)

If the phishing came from a compromised internal account, ANY second-tier victim that clicks will become tier 3. Stop the spread first.

1. **Identify patient zero** — the account whose mail propagated the phish. Usually the earliest sender in the campaign.
2. **Block sign-in for patient zero immediately** (full incident response is fine, but if you're in a hurry, just the block):
   ```powershell
   Update-MgUser -UserId patient-zero@contoso.com -AccountEnabled $false
   Revoke-MgUserSignInSession -UserId <patient-zero-id>
   ```
3. **Search sent mail** to identify additional victims:
   ```powershell
   Search-UAL -UserId patient-zero@contoso.com -From (Get-Date).AddDays(-1) -Operations @('Send') | Format-Table CreationDate, ObjectId
   ```
   Or from Exchange Admin Center → Mail flow → Message trace.

## Step 2 — Compile the victim list (minutes 5-15)

Pull the recipients of patient zero's outbound phishing. From the UAL search above, expand `AuditData.ToRecipients`.

Then add anyone who reported receiving the phish from external sources (some campaigns spoof the company; mail from outside-the-tenant phishing isn't visible in your UAL).

Build a CSV at `victims.csv`:

```csv
UPN,Severity,Reason,QuarantineSentMail
alice@contoso.com,High,Phishing campaign 2026-05-14,false
bob@contoso.com,High,Phishing campaign 2026-05-14,false
charlie@contoso.com,Critical,Phishing campaign + AiTM signature on inbox rule,true
patient-zero@contoso.com,Critical,Sent the phishing,true
```

`Severity` per victim:
- `Low` if they received but didn't click (forensic capture only).
- `Medium` if they received + you're not sure if they clicked.
- `High` if they clicked + entered credentials.
- `Critical` if MFA was satisfied + you see suspicious activity post-click (AiTM session theft).

`QuarantineSentMail=true` only for Critical — the operator will be prompted to type `PURGE` even with the flag set.

## Step 3 — Bulk incident response (minutes 15-45)

```powershell
Invoke-BulkIncidentResponse -Path victims.csv
```

What happens for each victim:

1. Snapshot (forensic baseline).
2. Containment per severity (block / revoke sessions / wipe MFA / force password change).
3. Cleanup (disable inbox rules + clear forwarding).
4. Audit (24h sign-ins + 7d sent mail + 7d shares).
5. Notify the security team.
6. Generate an incident report.

For Critical rows, the operator gets prompted at the quarantine step.

The bulk flow writes:
- A result CSV next to `victims.csv` with `Status` per row.
- An aggregate HTML at `<stateDir>/<tenant>/incidents/bulk-<ts>/index.html` linking every sub-incident.
- Per-victim incident reports under `<stateDir>/<tenant>/incidents/<INC-...>/report.html`.

## Step 4 — Communicate (minutes 45-60)

Once containment is done, the operator (out-of-band, NOT via the affected accounts):

1. **Call each affected user** — tell them their account is locked, here's what happened, here's how to recover (issue a TAP, walk them through MFA re-enrollment).
2. **Notify recipients of the phishing** — anyone who got the message from patient zero should know it was phishing.
3. **Update the incident tickets** in your ticket-tracking with the incident ids.

## Step 5 — Audit reach (within 24h)

For each Critical victim, audit downstream systems they had access to:

- CRM / Salesforce / HubSpot
- GitHub / source repos
- Finance systems (banking portals, payment processors)
- Partner portals
- Cloud admin consoles (AWS, GCP, etc.) — federated SSO means an M365 compromise often gives the attacker access here too

For each system, look at the same 24-hour window pre-compromise + 24h post.

## Step 6 — Close out

Once recovery is complete per victim:

```powershell
Close-Incident -Id INC-2026-05-14-aaaa -Resolution "Verified user identity, re-enrolled MFA, monitored downstream systems for 24h, no further activity."
```

For false-positive victims (those who received but didn't click):

```powershell
Close-Incident -Id INC-2026-05-14-bbbb -Resolution "Received phishing but did not click; user confirmed via phone." -FalsePositive
```

`-FalsePositive` triggers `Undo-Incident` to unblock + restore.

## Step 7 — Compliance handoff

Bundle every incident for compliance retention:

```powershell
foreach ($id in (Get-Incidents -Days 1 | Where-Object { $_.reason -like 'Phishing campaign 2026-05-14*' }).id) {
    Export-Incident -Id $id -Path "/secure/phishing-2026-05-14/$id.zip"
}
```

Each ZIP contains snapshot + 3 audits + report + filtered audit-log slice.

## Step 8 — Post-incident review

Within a week:

1. **How did patient zero get phished?** Email gateway gap, missing CA policy, user training gap? Fix.
2. **Why did N victims click?** Improve phishing simulation training.
3. **Did the detector framework fire?** If you had auto-detection running with appropriate thresholds, `Detect-SuspiciousInboxRule` should have flagged patient zero. If it didn't, tune thresholds OR add the missing detector.
4. **Were notifications timely?** If the security team got the alert hours after the fact, fix the notification routing.

## Variations

### MSP / cross-tenant phishing

A single campaign hit multiple customer tenants. Use `Invoke-AcrossTenants` to scope the bulk response per tenant:

```powershell
Invoke-AcrossTenants -Tenants @('Contoso','Fabrikam') -Script {
    Invoke-BulkIncidentResponse -Path victims.csv
}
```

Each tenant's incidents land in its own `<stateDir>/<tenant>/incidents/` — no cross-contamination.

### Spear-phishing (one targeted account)

If only one account was hit and it was a high-value target (CFO, IT admin), escalate severity to `Critical` and run `compromised-account.md` directly rather than bulk. Trade speed for thoroughness on the audit + downstream-systems review.

## See also

- [`compromised-account.md`](compromised-account.md) — single-account playbook.
- [`incident-triggers.md`](incident-triggers.md) — auto-detection framework.
- [`../guides/incident-response.md`](../guides/incident-response.md) — operator reference for the playbook.
- [`../guides/tabletop-exercises.md`](../guides/tabletop-exercises.md) — `phishing-campaign` scenario for IR-team drills.
