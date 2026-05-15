# Onboarding users

Add one user or a batch of users to your tenant with the right licenses, group memberships, distribution lists, and shared-mailbox access from day one.

## When to use this

- A new hire is starting. You want them productive on day one.
- A consultant is joining a project and needs a temporary tenant identity.
- HR drops you a CSV every Friday with the week's new hires.

## Prereqs

- You're connected to the target tenant (see [`../getting-started/tenant-setup.md`](../getting-started/tenant-setup.md)).
- Your account has `User Administrator` + `License Administrator` + the right Exchange / SharePoint admin roles for any group / shared mailbox assignments.
- For role templates: `templates/role-<slug>.json` files exist for the roles you want to use.

## Three onboarding modes

### 1. Single user — interactive (menu)

**Slot 1 → Onboard New User**

```
> First name        : Alice
> Last name         : Smith
> User principal    : alice.smith@contoso.com   (defaults to first.last@tenant)
> Display name      : Alice Smith               (defaults to "First Last")
> Job title         : Sales Account Executive
> Department        : Sales - North America
> Office location   : Seattle
> Usage location    : US                        (required for license assignment)
> Manager UPN       : bob.manager@contoso.com   (optional, skipped if not found)
> Apply role-based onboarding template? Y/N: Y
  Available templates:
    1. sales-rep
    2. engineer
    3. exec-assistant
    4. contractor
    5. default
  > Pick template: 1

> Generate password (Y) or enter manually (N)? Y
> Issue Temporary Access Pass for first sign-in? Y

[Applying template: sales-rep]
[+] User created      : alice.smith@contoso.com
[+] License assigned  : SPE_E3
[+] Group added       : Sales-NorthAmerica
[+] Group added       : All-Employees
[+] DL added          : sales-announce@contoso.com
[+] TAP issued        : 60-minute single-use code
                          Code: XXXX-XXXX-XXXX
                          Deliver via SMS/phone -- not email.
[+] entryId           : a1b2c3d4-...
```

Every line corresponds to an audit log entry. Reverse via `Invoke-Undo -EntryId a1b2c3d4-...` if needed.

### 2. Single user — replicate from existing

If a new hire is replacing someone who left, or joining a team where you want their access to match a peer:

**Slot 1 → Onboard New User** → "Replicate from existing user"

The tool reads the source user's licenses + group memberships + DL memberships + manager and applies the same set to the new user. You still get prompted for first name / last name / UPN.

Common gotcha: replicating from a still-active user duplicates license consumption. The tool warns when the source has assigned licenses you're about to clone.

### 3. Bulk from CSV

For >2 users or for HR integration:

```powershell
Invoke-BulkOnboard -Path users.csv [-Template sales-rep] [-WhatIf]
```

CSV columns (sample at `templates/bulk-onboard-sample.csv`):

| Column | Required? | Notes |
|---|---|---|
| `FirstName` | yes | |
| `LastName` | yes | |
| `DisplayName` | no | Defaults to `FirstName LastName`. |
| `UserPrincipalName` (or `UPN`) | yes | Sign-in name + primary email. |
| `Manager` | no | UPN of an existing user; skipped with a warning if not found. |
| `Department`, `JobTitle`, `OfficeLocation` | no | |
| `UsageLocation` | yes | ISO country code (`US`, `GB`, `DE`, ...). Required for license assignment. |
| `Template` | no | Role-template key (`sales-rep`, `engineer`, etc.). Per-row value wins over the `-Template` flag. |
| `Password` | no | If blank, a 16-char random password is generated. Either way, ForceChangePasswordNextSignIn is set. |
| `IssueTAP` | no | `true` / `yes` / `1` issues a 60-minute single-use TAP via `/authentication/temporaryAccessPassMethods`. |

The bulk flow follows the validate-first pattern:

1. **Parse + validate every row.** Validation errors print + the batch aborts before any tenant call.
2. **Confirmation.** "Onboard N user(s)?"
3. **Per-row execution.** Per-row failure is recorded; the batch continues.
4. **Result CSV.** Written next to the input as `bulk-onboard-<ts>.csv` with `Status` ∈ `{Success, PartialSuccess, Failed, Preview}` + `Reason`.

## Role templates

Templates live at `templates/role-<slug>.json`. They declare what to assign — licenses, groups, distribution lists, shared mailboxes, calendar permissions — so the operator doesn't have to remember the set per role.

Example (`templates/role-sales-rep.json`):

```jsonc
{
  "name": "sales-rep",
  "description": "Standard sales rep — E3 + Sales channels + CRM shared mailbox.",
  "licenses":           ["SPE_E3"],
  "groups":             ["Sales-NorthAmerica", "All-Employees"],
  "distributionLists":  ["sales-announce@contoso.com"],
  "sharedMailboxes":    [
    { "mailbox": "crm-shared@contoso.com", "permission": "FullAccess" }
  ],
  "contractorExpiryDays": null
}
```

Drop a new JSON into `templates/` and it auto-discovers — no code change required. Schema reference: [`../reference/template-schema.md`](../reference/template-schema.md).

`contractorExpiryDays` (non-null) records `employeeLeaveDateTime` on the new user, surfacing them in stale-user reports as the date approaches. Auto-disable on that date still needs Entra Lifecycle Workflows (out of scope for this tool).

## What gets logged

Each onboard step is its own audit entry, with `actionType` covering the full set: `CreateUser`, `AssignLicense`, `AddToGroup`, `AddToDistributionList`, `GrantMailboxFullAccess`, `IssueTAP`, etc. All reversible via the standard undo dispatch — see [`audit-and-undo.md`](audit-and-undo.md).

## PREVIEW mode

In PREVIEW, every step produces an audit entry with `event=PREVIEW` and `result=preview`. No tenant calls fire. Useful for:

- HR'd "trial onboard": let HR see what the user would receive before going LIVE.
- Bulk CSV review: see the validation result + the plan without committing.

## Common failures

| Symptom | Cause + fix |
|---|---|
| `License assignment failed: License assignment cannot be applied to a user with no usage location` | The user's `UsageLocation` wasn't set (CSV column or interactive prompt). Set it (ISO country code) and retry. |
| `Group not found: 'Sales-NorthAmerica'` | The role template references a group that doesn't exist in this tenant. Check the template; either create the group or remove the reference. |
| `License SKU not subscribed: 'ENTERPRISEPACK'` | Your tenant doesn't have the SKU the template requests. Check available SKUs with `Get-MgSubscribedSku`. |
| `User principal already exists` | The UPN is taken. Pick a different UPN. (The tool doesn't auto-suffix.) |

Full troubleshooting at [`../operations/troubleshooting.md`](../operations/troubleshooting.md).

## See also

- [`../reference/csv-formats.md`](../reference/csv-formats.md) — every bulk CSV schema.
- [`../reference/template-schema.md`](../reference/template-schema.md) — role / site / scheduled-check templates.
- [`offboarding.md`](offboarding.md) — the inverse flow.
- [`audit-and-undo.md`](audit-and-undo.md) — viewing + reversing operations.
