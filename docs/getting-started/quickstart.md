# 10-minute quickstart

A guided hands-on tour. By the end you'll have:

- Connected to a tenant
- Registered a tenant profile
- Run one report (read-only)
- Run one mutation in PREVIEW mode
- Viewed the audit log
- Chatted with the AI assistant (optional)

Read end-to-end first, then execute. Every command below is meant to be **typed as-is** — substitute your tenant's UPNs only where placeholders are marked.

## Prereqs

You've completed [`installation.md`](installation.md). The tool launches, you have admin or read-only access to at least one M365 tenant.

## 1. Launch in PREVIEW mode

```cmd
Launch.bat
```

At the mode picker: **PREVIEW**. The banner will be **yellow** above the main menu — that's how you can always tell at a glance which mode you're in.

```
  [ PREVIEW MODE -- dry-run, no tenant changes ]
```

## 2. Connect to a tenant

At the tenant picker, choose your option:

- **My own organization (direct admin)** — your day-job tenant.
- **A customer tenant (GDAP partner access)** — if you're an MSP / partner.

A browser pops open. Consent to the requested scopes (first time only). After OAuth completes you'll see:

```
  Tenant: Contoso (contoso.onmicrosoft.com)
    Graph: OK    EXO: OK    SCC: OK
```

If any of those shows `---` instead of `OK`, the matching service is unreachable. The tool will still run; cmdlets that need that service will fail individually. See [`../operations/troubleshooting.md`](../operations/troubleshooting.md).

## 3. Register a tenant profile

If this is the same tenant you'll often work in, register a profile so you can switch back to it instantly later. From the main menu:

- **Slot 21 → Tenants** → "Register a new tenant"
- Name it (e.g. `Contoso`)
- The tool reads the tenant ID + primary domain from the active connection
- Authentication mode: pick **Interactive** for now (cert + client-secret modes are covered in [tenant-setup.md](tenant-setup.md))

You'll see:

```
  [+] Registered tenant 'Contoso' (Interactive).
```

The profile lives at `<stateDir>\tenants.json` (plaintext metadata; secrets go through DPAPI in a separate file). See [`../concepts/multi-tenant.md`](../concepts/multi-tenant.md).

## 4. Run a read-only report

From the main menu:

- **Slot 11 → Reporting** → "Tenant overview"

Output (truncated):

```
  Tenant overview -- Contoso
    Active users          : 1247
    Disabled users        : 38
    Licensed users        : 1198
    Guest users           : 64
    Distinct SKUs         : 9
    Active Teams          : 213
```

Nothing was changed. This is a pure Graph read.

## 5. Run one mutation in PREVIEW

We'll use a license operation that's safe to "do" because PREVIEW only logs the intent. Pick a test account in your tenant.

- **Slot 3 → Add / Remove License**
- Enter the test UPN
- Pick "Remove a license"
- Pick any SKU the user has

Expected output:

```
  [PREVIEW] Would remove license 'SPE_E3' from <test-user@yourdomain.com>
            entryId: a1b2c3d4-...
            reverseType: AssignLicense
```

The `[PREVIEW]` tag means **nothing happened in the tenant**. The audit log got a line; if you switch to LIVE later and call `Invoke-Undo -EntryId a1b2c3d4-...`, you'll see "this entry was preview only, no reversal needed".

## 6. View the audit log

- **Slot 14 → Audit & Reporting** → "Open audit log viewer"

You'll see a table of every action this session. Filter / sort / view detail. Press `E` to export to CSV or HTML. The full reference for the line format is at [`../reference/audit-format.md`](../reference/audit-format.md).

## 7. (Optional) Talk to the AI assistant

If you ran the AI setup during install (option 99 → `/config`), try a few chat commands:

- `/about` — diagnostic snapshot. Shows which provider you're on, plan-mode state, cost totals.
- `/tools` — list every tool the AI can call.
- `Show me the top 5 most-licensed users.` — natural-language request. The AI proposes a tool call (`Get-LicenseAssignments` or similar); you approve `[Y]es` / `[A]ll` / `[N]o` / `[Q]uit`.
- `/quit` — exits the assistant, auto-saves the chat under `<stateDir>\chat-sessions\`.

See [`../concepts/ai-assistant.md`](../concepts/ai-assistant.md) for how Mark works conceptually.

## 8. Quit cleanly

From the main menu: select the `-` option (back/quit). The tool disconnects every service, writes a `SESSION_END` audit line, and exits.

```
  All sessions disconnected.
  Goodbye!
```

## What you learned

- PREVIEW vs LIVE — colored banners, mode picker on launch, `/dryrun` to flip mid-chat.
- Tenant profiles — `<stateDir>\tenants.json`, instant switching.
- Audit log — every action logged with `entryId`, `actionType`, `target`, and (where applicable) a `reverse` recipe.
- Read vs mutation — `Reports`, `Audit & Reporting`, `Sign-in lookup` etc. are pure reads; everything that says "Manage" or "Onboard" / "Offboard" mutates and goes through `Invoke-Action`.
- The AI assistant lives at option 99 but is hidden from the menu list — see [`../concepts/ai-assistant.md`](../concepts/ai-assistant.md).

## Next steps

| If you want to… | Read this |
|---|---|
| Run a real onboard | [`../guides/onboarding.md`](../guides/onboarding.md) |
| Offboard a user | [`../guides/offboarding.md`](../guides/offboarding.md) |
| Respond to a compromised account | [`../guides/incident-response.md`](../guides/incident-response.md) |
| Set up many tenants for MSP work | [tenant-setup.md](tenant-setup.md) |
| Tune the config | [configuration.md](configuration.md) |
| Connect notifications | [`../guides/notifications.md`](../guides/notifications.md) |
