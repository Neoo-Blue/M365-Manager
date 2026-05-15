# CSV formats

Every bulk-flow CSV schema in one place. Each section: columns, required-vs-optional, sample row, sample at `templates/<name>-sample.csv`.

All bulk flows follow the same Phase 1 pattern: **validate every row first** → confirm → per-row execution that does NOT halt on per-row failure → result CSV written next to input with `Status` ∈ `{Success, PartialSuccess, Failed, Preview}` + `Reason`.

## Onboarding

`Invoke-BulkOnboard -Path users.csv [-Template <name>] [-WhatIf]`

Sample: `templates/bulk-onboard-sample.csv`.

| Column | Required | Notes |
|---|---|---|
| `FirstName` | yes | |
| `LastName` | yes | |
| `DisplayName` | no | Defaults to `FirstName LastName`. |
| `UserPrincipalName` (alias `UPN`) | yes | Used as both sign-in name and primary email. |
| `Manager` | no | UPN of an existing user; skipped with a warning if not found. |
| `Department` | no | |
| `JobTitle` | no | |
| `OfficeLocation` | no | |
| `UsageLocation` | yes | ISO country code (`US`, `GB`, `DE`, ...). Required for license assignment. |
| `Template` | no | Role-template key. Per-row value wins over the `-Template` flag. |
| `Password` | no | If blank, 16-char random generated. ForceChangePasswordNextSignIn always set. |
| `IssueTAP` | no | `true`/`yes`/`1` issues a 60-min single-use TAP. |

Result CSV adds: `Status`, `Reason`.

## Offboarding

`Invoke-BulkOffboard -Path leavers.csv [-WhatIf]`

Sample: `templates/bulk-offboard-sample.csv`.

| Column | Required | Notes |
|---|---|---|
| `UserPrincipalName` (alias `UPN`) | yes | |
| `ForwardTo` | no | UPN to forward mail to during the offboard window. |
| `ConvertToShared` | no | `true`/`yes`/`1` to convert mailbox to Shared. |
| `HandoffOneDriveTo` | no | UPN. Triggers `Invoke-OneDriveHandoff` for this row. |
| `RemoveFromAllGroups` | no | `true`/`yes`/`1` to walk every group + DL. |
| `Reason` | yes | Free text; recorded in every step's audit entry. |

## Bulk MFA reset

`Invoke-BulkMfaReset -Path mfa-reset.csv [-WhatIf]`

| Column | Required | Notes |
|---|---|---|
| `UPN` | yes | |
| `IssueTAP` | no | `true`/`yes`/`1` issues a TAP after wipe. |
| `TAPLifetimeMinutes` | no | Default 60; max 480. |
| `Reason` | yes | |

## Teams membership

`Invoke-BulkTeamsMembership -Path teams.csv [-WhatIf]`

| Column | Required | Notes |
|---|---|---|
| `UPN` | yes | |
| `TeamId` | one of TeamId/TeamName | Stable id (preferred). |
| `TeamName` | one of TeamId/TeamName | Display name (contains-match). |
| `Action` | yes | `Add` / `Remove` / `Promote` / `Demote`. |
| `Role` | for `Add` only | `Member` / `Owner`. |

## Guest removal

`Invoke-BulkGuestRemoval -Path guests.csv`

Sample: `templates/bulk-guest-removal-sample.csv`.

| Column | Required | Notes |
|---|---|---|
| `UPN` | yes | Guest UPN. |
| `Reason` | yes | Recorded in every per-step audit entry. |

## Guest recertification campaign

`Start-RecertCampaign -CampaignId <id> -Path guests.csv`

Sample: `templates/guest-recertification-sample.csv`.

| Column | Required | Notes |
|---|---|---|
| `UPN` | yes | |
| `InvitedBy` | no | UPN of the inviter. Used for the manager-prompt email. |
| `ReviewerUpn` | no | Overrides `InvitedBy` for who gets the recert email. |
| `ProjectContext` | no | Free text included in the email body. |

## License remediation

`Invoke-BulkLicenseRemoval -Path license-cleanup.csv [-WhatIf]`

Sample: `templates/license-remediation-sample.csv`.

| Column | Required | Notes |
|---|---|---|
| `UPN` | yes | |
| `SkuId` | yes | Microsoft SKU GUID. |
| `Reason` | yes | |
| `ApprovedBy` | no | Audit metadata. |

## Incident response (Phase 7)

`Invoke-BulkIncidentResponse -Path incidents.csv [-WhatIf]`

Sample: `templates/incidents-bulk-sample.csv`.

| Column | Required | Notes |
|---|---|---|
| `UPN` | yes | |
| `Severity` | yes | `Low` / `Medium` / `High` / `Critical`. |
| `Reason` | no | Free text. |
| `QuarantineSentMail` | no | `true`/`yes`/`1`. Only applies at Critical severity; always operator-confirmed before purging. |

## CSV encoding

The validator + Import-Csv expect:

- **UTF-8** (with or without BOM).
- **Comma-delimited** by default.
- **Header row** matching the columns above (case-insensitive).

If your CSV is from Excel, save as "CSV UTF-8 (Comma delimited) (*.csv)" — the regular "CSV" option emits Windows-1252 and double-byte names get mangled.

## Result CSV

Every bulk flow writes `<input-name>-result-<ts>.csv` next to the input. Columns vary per flow but always include:

- `Row` — 1-based row number in the input.
- `Status` — `Success` / `PartialSuccess` / `Failed` / `Preview`.
- `Reason` — failure reason, or empty on success.

Status meaning:

- **Success** — the row's primary action completed.
- **PartialSuccess** — one of the row's per-row sub-actions failed but the rest succeeded.
- **Failed** — the row's primary action failed.
- **Preview** — PREVIEW mode; nothing was actually done.

## See also

- [`template-schema.md`](template-schema.md) — role / site / scheduled-check JSON schemas.
- [`../guides/onboarding.md`](../guides/onboarding.md) · [`../guides/offboarding.md`](../guides/offboarding.md) · etc. — operator walkthroughs.
