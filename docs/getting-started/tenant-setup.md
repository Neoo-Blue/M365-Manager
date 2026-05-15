# Tenant setup

Connecting to one tenant on demand vs registering a profile for repeated use.

## Two paths

### Path A — connect once

For exploration, one-off scripts, or a single-tenant operator who doesn't need fast switching.

On launch the tool asks:

```
Which tenant are you managing?
  1. My own organization (direct admin)
  2. A customer tenant (GDAP partner access)
```

Pick one. A browser pops, you OAuth, you're connected. The tenant context lasts for the lifetime of this PowerShell process.

### Path B — register a profile

For repeated use, multi-tenant work, or if you want fast `Switch-Tenant -Name <n>` without going through the OAuth browser dance each time.

From the menu: **Slot 21 → Tenants** → "Register a new tenant".

The profile lives at `<stateDir>\tenants.json` (plaintext metadata) with secrets in `<stateDir>\secrets\tenant-<name>.dat` (DPAPI-encrypted; per-user, per-machine, not portable).

## Three authentication modes

| Mode | Use when | Notes |
|---|---|---|
| **Interactive** | Most operators most of the time | Re-prompts in the browser when token expires. Cleanest. |
| **CertThumbprint** | Unattended automation, scheduled tasks | App-only auth, no human in loop. Requires registering an Entra app + uploading a cert. **Preferred** for automation. |
| **ClientSecret** | Same use case, less ideal | App-only auth via shared secret. Same registration path but cert is more secure. Warned at register time. |

### Registering an Interactive profile

The most common path. From the Tenants menu:

```
> Tenant name (any short label): Contoso
> Tenant ID (Azure AD directory id GUID, or 'auto' to detect from current connection): auto
> Authentication mode: 1. Interactive

[+] Registered tenant 'Contoso' (Interactive).
```

That's it. Switching back later:

```powershell
Switch-Tenant -Name Contoso
# or from the assistant:
/tenant Contoso
```

### Registering a CertThumbprint profile (unattended)

Setup work upfront, smooth running afterward.

1. **Create an Entra app registration** in the target tenant:
   - Sign in to https://entra.microsoft.com as a tenant admin.
   - **App registrations** → **New registration** → name it (e.g. `M365 Manager Automation`).
   - **API permissions** → Add the Microsoft Graph + Office 365 Exchange Online application permissions you'll actually use. Start narrow:
     - `User.Read.All`, `Group.Read.All`, `Directory.Read.All` for reports.
     - Add write scopes (`User.ReadWrite.All`, `Group.ReadWrite.All`, `MailboxSettings.ReadWrite`, etc.) only as needed.
   - Click **Grant admin consent**.

2. **Create or upload a certificate**:
   - **Certificates & secrets** → **Certificates** → Upload your `.cer` (public key). Generate a self-signed one for testing:
     ```powershell
     $cert = New-SelfSignedCertificate -Subject "CN=M365 Manager Automation" `
                 -CertStoreLocation "Cert:\CurrentUser\My" `
                 -KeyExportPolicy NonExportable `
                 -KeySpec Signature `
                 -KeyLength 2048 `
                 -HashAlgorithm SHA256 `
                 -NotAfter (Get-Date).AddYears(2)
     Export-Certificate -Cert $cert -FilePath .\m365mgr.cer
     # Upload m365mgr.cer to the Entra app.
     # Capture the thumbprint:
     $cert.Thumbprint
     ```

3. **Note three values**:
   - **Tenant ID** (the directory id, a GUID — visible on the app registration page).
   - **Application (client) ID** (the app id — also visible).
   - **Certificate thumbprint** (from step 2).

4. **Register the profile**:

   ```powershell
   Register-Tenant -Name 'ContosoAuto' -TenantId '<tenant-guid>' `
       -AuthMode CertThumbprint `
       -ClientId '<app-client-id>' `
       -Thumbprint '<cert-thumbprint>'
   ```

5. **Verify** by switching:
   ```powershell
   Switch-Tenant -Name 'ContosoAuto'
   # Should connect app-only without opening a browser.
   ```

### Registering a ClientSecret profile

Same shape, less safe — the secret is stored DPAPI-encrypted but anyone who exfiltrates the secret has full app-permission access.

```powershell
Register-Tenant -Name 'ContosoSecret' -TenantId '<tenant-guid>' `
    -AuthMode ClientSecret `
    -ClientId '<app-client-id>' `
    -ClientSecret (Read-Host -AsSecureString)
```

You'll see a warning:

```
[!] Storing a client_secret on disk -- cert+thumbprint is safer. Rotate this secret regularly.
```

## GDAP / partner mode

If you're an MSP managing customer tenants via GDAP (Granular Delegated Admin Privileges):

1. **Sign in to your partner tenant** at the initial launch.
2. **Pick "A customer tenant (GDAP partner access)"** at the tenant picker.
3. The tool lists every customer tenant your account has delegated rights to.
4. Pick one. The connection runs against the customer.
5. **Register a profile per customer** so you can `Switch-Tenant -Name <customer>` later. The profile stores the customer's tenant ID; auth still goes through your partner credential at runtime.

See [`../concepts/multi-tenant.md`](../concepts/multi-tenant.md) for the full MSP cockpit story.

## Switching tenants

```powershell
# Listing
Get-Tenants
# Or from the chat:
/tenants

# Switching
Switch-Tenant -Name Contoso
# Or from the chat:
/tenant Contoso
```

The audit log filename gets a tenant slug — `session-<ts>-<pid>-contoso.log` vs `session-<ts>-<pid>-fabrikam.log` — so cross-tenant operations stay distinct on disk.

## Removing a profile

```powershell
Remove-Tenant -Name Contoso
```

Removes both the profile from `tenants.json` and the DPAPI secret manifest (if any). Does NOT touch the actual Entra app or revoke any tokens — you'd handle that in the Entra portal.

## Profile portability

`tenants.json` is portable across machines (plaintext metadata only). The DPAPI secret manifests are NOT — they're tied to the user account that encrypted them. If you move the profile to a new machine, re-run `Register-Tenant` to recreate the secret manifest under the new DPAPI scope.

## See also

- [`../concepts/multi-tenant.md`](../concepts/multi-tenant.md) — full multi-tenant model, audit, cross-tenant reports.
- [`../guides/msp-dashboard.md`](../guides/msp-dashboard.md) — portfolio overview HTML.
- [`../operations/permissions.md`](../operations/permissions.md) — Graph scope + role matrix per feature.
