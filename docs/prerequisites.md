# Prerequisites — App Registration with Certificate-based Authentication

This document is the one-time setup guide for running the AroraMSP audit scripts against a Microsoft 365 tenant using **certificate-based app-only authentication**. Allow ~20 minutes the first time.

The TL;DR: register an app, upload a certificate, grant a defined set of read-only permissions, assign the Exchange Administrator role, and run the scripts with three parameters. After that, every run is silent — no prompts, no broker, no token refresh dramas.

---

## Why certificate authentication?

The scripts support three ways to authenticate. They are not equivalent:

| Method | Suitable for | Risk profile |
| --- | --- | --- |
| **Certificate-based** (recommended) | Scheduled tasks, CI runners, ad-hoc engineer-on-laptop runs | Private key stays in the local certificate store; nothing transits the network or gets cached anywhere reusable |
| **Device code** (`-UseDeviceCode`) | Interactive use only when certs are not yet set up, or when running from a host you can't permanently trust | Browser-based; relies on the user's signed-in identity and broker token; subject to broker-persistence quirks on Windows 11 |
| **Client secret** (not supported by these scripts) | — | Long string in plaintext; rotated rarely; leaked secrets are the #1 cause of compromised app registrations |

This is consistent with Microsoft's [security best practices for application registrations](https://learn.microsoft.com/en-us/azure/active-directory/develop/security-best-practices-for-app-registration), which explicitly state: *"Use certificate credentials over password credentials (client secrets) where possible. Certificates are typically more secure than client secrets."*

Why certs win:

- **The private key never leaves the host.** The certificate store on the machine holds the private key under DPAPI; `Connect-MgGraph` reads the thumbprint, the OS signs an assertion, and only the public certificate ever reaches Entra ID.
- **No long-lived bearer credential.** A leaked client secret is a leaked password — it works from anywhere until rotated. A leaked thumbprint is useless without the corresponding private key, which is bound to the local machine and (for non-exportable keys) cannot be copied off it.
- **Auditable.** Each token Entra issues to your app is tied to the certificate fingerprint. Rotating a cert invalidates every prior token immediately — there is no equivalent for password-based auth.
- **Unattended-friendly.** No `Connect-MgGraph` browser dance, no device-code prompt, no broker. Suitable for cron-style scheduled runs.

Device code is supported as a fallback for engineers who haven't yet stood up the app registration or who are running from a machine where they can't install a certificate. It is **not** recommended for routine use.

---

## Step 1 — Generate the certificate

Run this on the machine that will execute the scripts. The cert is created in the **CurrentUser** store, marked **non-exportable** (the private key cannot be copied off the host), and signed with SHA-256 over a 2048-bit RSA key.

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=AroraMSP-AuditScripts" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy NonExportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export the public-key .cer for upload to Entra ID
Export-Certificate -Cert $cert -FilePath "$env:USERPROFILE\Desktop\AroraMSP-AuditScripts.cer"

# Display the thumbprint — you'll pass this to the scripts via -CertificateThumbprint
$cert.Thumbprint
```

Parameter-by-parameter:

| Parameter | Why |
| --- | --- |
| `-Subject "CN=AroraMSP-AuditScripts"` | Friendly name shown in `certmgr.msc`. Must start with `CN=`. |
| `-CertStoreLocation "Cert:\CurrentUser\My"` | Personal store of the current user. Use `LocalMachine\My` if you need other accounts on this host (e.g. a service account) to be able to read the key — that requires elevated PowerShell. |
| `-KeyExportPolicy NonExportable` | Prevents `Export-PfxCertificate` from copying the private key off the host. The certificate cannot be moved between machines. This is the security-relevant flag. |
| `-KeySpec Signature` | The key is for signing assertions, not encryption. |
| `-KeyLength 2048` | Minimum acceptable RSA key length. 4096 is fine but slower. |
| `-HashAlgorithm SHA256` | SHA-1 is deprecated by Entra ID. SHA-384 and SHA-512 are also accepted. |
| `-NotAfter (Get-Date).AddYears(2)` | 2-year validity. Set a calendar reminder; see the [Renewal](#renewal) section. |

After running this, the public key sits at `~/Desktop/AroraMSP-AuditScripts.cer` and the private key is in your **personal** certificate store. The thumbprint output by the last line is what you'll pass to the scripts.

If you want to retrieve the thumbprint later:

```powershell
Get-ChildItem Cert:\CurrentUser\My | Where-Object Subject -eq "CN=AroraMSP-AuditScripts" | Select-Object Thumbprint, NotAfter
```

---

## Step 2 — Register the app in Entra ID

1. Sign in to the [Entra admin center](https://entra.microsoft.com) as a **Global Administrator** or **Application Administrator**. Lower-privileged roles cannot register apps unless the tenant policy *Users can register applications* is enabled.
2. Navigate to **Identity → Applications → App registrations → New registration**.
3. Fill in the form:

| Field | Value | Why |
| --- | --- | --- |
| **Name** | `AroraMSP Audit Scripts` (or any descriptive name) | Shown in audit logs. Make it specific so a future admin reading sign-in logs immediately understands what this app does. |
| **Supported account types** | **Accounts in this organizational directory only (single tenant)** | The app should only be callable by identities in your tenant. Multi-tenant is for SaaS providers; we are not a SaaS provider. |
| **Redirect URI** | *Leave blank* | Redirect URIs are for delegated/interactive flows. Certificate auth is app-only and uses no redirect. |

4. Click **Register**.

Record two values from the **Overview** page:

- **Application (client) ID** — a GUID. Passed to the scripts via `-ClientId`.
- **Directory (tenant) ID** — also a GUID. Passed via `-TenantId`.

Both are non-secret — they identify the app and the tenant, but cannot be used for authentication on their own.

---

## Step 3 — Upload the certificate

1. From the app registration's left navigation: **Certificates & secrets**.
2. Open the **Certificates** tab.
3. Click **Upload certificate**.
4. Browse to the `.cer` file produced in Step 1 (`~/Desktop/AroraMSP-AuditScripts.cer`).
5. Add a description (e.g. `Self-signed cert on <hostname> — expires 2028-05-09`).
6. Click **Add**.

Verify the **Thumbprint** shown in the portal matches the one PowerShell printed in Step 1. If they don't match, you uploaded the wrong file.

> **Why upload only the public key**: Entra ID needs the public key to verify assertions signed by your private key. The private key never leaves your host. This is the entire point of asymmetric crypto in this flow.

---

## Step 4 — Add the Microsoft Graph application permissions

1. From the app registration: **API permissions → Add a permission → Microsoft Graph → Application permissions**.
2. Find and check each of the following, then click **Add permissions**.

| Permission | Used by | Why |
| --- | --- | --- |
| `Directory.Read.All` | Tenant Audit | Read role assignments (Global Admin count, break-glass exclusions). |
| `Policy.Read.All` | Tenant Audit | Read Conditional Access policies (legacy auth block, MFA enforcement, exclusions). |
| `Domain.Read.All` | Tenant Audit | Enumerate verified domains for DMARC, SPF, and DKIM checks. |
| `Organization.Read.All` | Both scripts | Read tenant display name and primary verified domain (used for the EXO `-Organization` parameter). |
| `User.Read.All` | Both scripts | Read sign-in activity (inactive licensed users) and licence assignments. |
| `Group.Read.All` | Tenant Audit | Resolve group membership for CA exclusion analysis. |
| `AuditLog.Read.All` | Tenant Audit | Read sign-in logs. Requires Entra ID P1 or higher to actually populate. |
| `DeviceManagementApps.Read.All` | Tenant Audit | Read App Protection (MAM) policy counts. |
| `DeviceManagementConfiguration.Read.All` | Tenant Audit | Read device compliance policies. |
| `Reports.Read.All` | Mailbox Report | Read the mailbox usage report (drives the Last Activity column). |

> **Application** vs **Delegated**: select **Application** permissions. Application permissions act on behalf of the app (no signed-in user); delegated permissions act on behalf of a signed-in user. Certificate-auth is app-only, so it can only consume application permissions.

> **All these are read-only.** The scripts do not modify tenant configuration. If a future feature needs write access, add it explicitly here — never grant `*.ReadWrite.All` "just in case".

---

## Step 5 — Add the Exchange Online application permission

1. **API permissions → Add a permission → APIs my organization uses**.
2. Search for and select **Office 365 Exchange Online**.
3. Choose **Application permissions**.
4. Find and check **`Exchange.ManageAsApp`**. Click **Add permissions**.

`Exchange.ManageAsApp` is the gate — without it, the EXO module refuses to authenticate the app at all. But it does not by itself authorise any actual EXO operation. That comes from the role assignment in Step 7.

---

## Step 6 — Grant admin consent

1. Back on the **API permissions** page, click **Grant admin consent for `<your tenant>`**.
2. Confirm in the dialog.

After consent, the **Status** column for every permission you added should change to a green tick: **Granted for `<your tenant>`**.

> **What admin consent means**: Application permissions cannot be exercised until a tenant administrator explicitly authorises them. This is the firewall between "permission is requested" and "permission is usable". An app that has been granted `Directory.Read.All` but not consented can register the request but cannot read directory data. Consent is a deliberate human-in-the-loop checkpoint, and it is required for every application permission you add.

> **Why "for `<tenant>`"**: Consent is per-tenant, not per-app-registration. If you ever export this app to another tenant, that tenant's admin must consent independently.

---

## Step 7 — Assign the Exchange Administrator role to the app

`Exchange.ManageAsApp` lets the app *call* Exchange Online. It does not give the app *any* RBAC role within Exchange Online. Without an Exchange role assignment, every EXO cmdlet returns:

```
The role assigned to application <id> isn't supported in this scenario.
```

Fix:

1. In Entra admin center: **Identity → Roles & admins → Roles & administrators**.
2. Search for **Exchange Administrator**.
3. Click the role name → **Add assignments → Select members**.
4. Search for the app registration name (e.g. `AroraMSP Audit Scripts`) and select it.
5. Click **Next**, then **Assign**.

The role assignment is permanent for service principals — there is no PIM eligibility for app principals.

> **Lower-privilege option**: If you only ever run the **Mailbox Report** script (which only reads `Get-EXOMailbox`, `Get-EXOMailboxStatistics`, and the Graph mailbox usage report), you can assign **Global Reader** instead of **Exchange Administrator**. The Tenant Audit script reads transport rules, anti-phishing config, DKIM signing config, and audit settings, all of which are also visible to Global Reader. In short: **Global Reader is sufficient for both scripts** if you want strict least-privilege. Use **Exchange Administrator** only if a future enhancement starts modifying EXO state, which would warrant a separate discussion.

---

## Step 8 — Find and copy the certificate thumbprint

You already have the thumbprint from Step 1. To retrieve it later from PowerShell:

```powershell
Get-ChildItem Cert:\CurrentUser\My |
    Where-Object Subject -eq "CN=AroraMSP-AuditScripts" |
    Select-Object Thumbprint, Subject, NotAfter
```

To retrieve it from `certmgr.msc`:

1. Run `certmgr.msc`.
2. Navigate to **Personal → Certificates**.
3. Double-click the certificate → **Details** tab → scroll to **Thumbprint**.
4. Click the thumbprint value to display the full hex string.
5. Copy it. It's case-insensitive but typically rendered as upper-case hex with no spaces (e.g. `AABBCCDDEEFF00112233445566778899AABBCCDD`).

The thumbprint is **not** secret — it is essentially a fingerprint. Possessing the thumbprint does not let an attacker authenticate; only the corresponding private key can do that. It can be safely placed in a configuration file or scheduled-task definition.

---

## Step 9 — Run the scripts

```powershell
$tenantId   = "00000000-0000-0000-0000-000000000000"  # from Step 2
$clientId   = "11111111-1111-1111-1111-111111111111"  # from Step 2
$thumbprint = "AABBCCDDEEFF00112233445566778899AABBCCDD" # from Step 8

.\scripts\Invoke-AroraMSPTenantAudit.ps1 `
    -TenantId $tenantId `
    -ClientId $clientId `
    -CertificateThumbprint $thumbprint

.\scripts\Invoke-AroraMSPMailboxReport.ps1 `
    -TenantId $tenantId `
    -ClientId $clientId `
    -CertificateThumbprint $thumbprint
```

Either script will:

1. Connect to Microsoft Graph silently using the certificate.
2. Resolve the primary verified domain via `Get-MgOrganization` (required for the EXO `-Organization` parameter — EXO expects a domain, not a tenant GUID).
3. Connect to Exchange Online silently using the same certificate plus the resolved domain.
4. Pull data, build the HTML report, save it to `$OutputDirectory`.

No prompts, no browser, no broker.

---

## Renewal

Self-signed certificates created in Step 1 expire in 2 years. To renew without downtime:

1. Run the same `New-SelfSignedCertificate` command in Step 1 (it generates a new key pair). Note the new thumbprint.
2. Export the new public-key `.cer` and upload it to the existing app registration (Step 3). **Both certificates can coexist** on the registration.
3. Update the script invocation to use the new thumbprint. Verify a successful run.
4. Once verified, delete the old certificate from the app registration. The old private key in your local store can also be deleted at this point.

Set a calendar reminder for ~30 days before expiry. A failed run after expiry produces:

```
AADSTS700027: Client assertion contains an invalid signature.
```

---

## Security notes

This setup follows the [Microsoft Zero Trust](https://learn.microsoft.com/en-us/security/zero-trust/) and least-privilege principles:

- **Verify explicitly**: every API call is signed with the app's certificate; Entra ID validates the signature against the public key it has on file.
- **Use least-privileged access**: only specific read-only Microsoft Graph and Exchange Online permissions are granted. No `*.ReadWrite.All` or `Directory.AccessAsUser.All`. The Exchange role can be downscoped from Exchange Administrator to Global Reader.
- **Assume breach**: if the host running the scripts is compromised, the non-exportable private key cannot be copied off it. Worst case is the attacker calls the scripts from that host, which is a read-only blast radius. Rotating the cert (delete on the app registration) immediately invalidates the credential; no further action on the host is required.

What this design deliberately avoids:

- **Long-lived secrets on disk** — there are none. The private key is held by DPAPI in the certificate store, accessible only to the user who created it.
- **Tokens cached in plain files** — the Graph SDK caches access tokens in memory only when authenticating with a certificate; nothing on disk.
- **Network-routable credentials** — the cert never traverses the wire. Only signed assertions and the resulting bearer tokens (which Entra issues per-call) do.

For audit purposes:

- App sign-ins appear in **Entra admin center → Monitoring → Sign-in logs → Service principal sign-ins** with the app's name and certificate thumbprint.
- Certificate creation, upload, and deletion events are recorded under **Audit logs**.
- Token issuance is logged per-call. Set up alerts on unexpected sign-in patterns if the app is widely deployed.
