# Prerequisites — Certificate-based App Registration

Both scripts authenticate to Microsoft Graph and Exchange Online via an Entra ID app registration with a certificate credential. This is the recommended path: no interactive prompts, no token-persistence quirks, suitable for scheduled runs.

This document walks through the one-time setup. Allow ~15 minutes.

## 1. Create the certificate

On the machine that will run the scripts, generate a self-signed certificate and export the public-key `.cer` for upload to Entra ID. Keep the private key in the local certificate store — the scripts read it via `-CertificateThumbprint`.

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=AroraMSP-AuditScripts" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy NonExportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export public key for upload to Entra ID
Export-Certificate -Cert $cert -FilePath "$env:USERPROFILE\Desktop\AroraMSP-AuditScripts.cer"

# Note the thumbprint — you'll pass this to the script via -CertificateThumbprint
$cert.Thumbprint
```

Record the thumbprint — you'll need it when running the scripts.

## 2. Register the app in Entra ID

1. Sign in to the [Entra admin center](https://entra.microsoft.com) as a Global Administrator or Application Administrator.
2. Go to **Identity → Applications → App registrations → New registration**.
3. **Name**: `AroraMSP Audit Scripts` (or any descriptive name).
4. **Supported account types**: *Accounts in this organizational directory only (single tenant)*.
5. **Redirect URI**: leave blank.
6. Click **Register**.

Record:
- **Application (client) ID** — passed to the scripts via `-ClientId`.
- **Directory (tenant) ID** — passed to the scripts via `-TenantId`.

## 3. Upload the certificate

1. Open the new app registration.
2. **Certificates & secrets → Certificates → Upload certificate**.
3. Upload the `.cer` file you exported in step 1.
4. Confirm the thumbprint shown in the portal matches the one from step 1.

## 4. Add Microsoft Graph application permissions

1. **API permissions → Add a permission → Microsoft Graph → Application permissions**.
2. Add the following:

| Permission | Used by |
| --- | --- |
| `Directory.Read.All` | Tenant Audit |
| `Policy.Read.All` | Tenant Audit (Conditional Access) |
| `Domain.Read.All` | Tenant Audit (DMARC, SPF, DKIM domain enumeration) |
| `Organization.Read.All` | Both scripts (tenant name, primary verified domain) |
| `User.Read.All` | Both scripts (sign-in activity, licence assignments) |
| `Group.Read.All` | Tenant Audit (CA exclusion groups) |
| `AuditLog.Read.All` | Tenant Audit (sign-in activity) |
| `DeviceManagementApps.Read.All` | Tenant Audit (App Protection policies) |
| `DeviceManagementConfiguration.Read.All` | Tenant Audit (compliance policies) |
| `Reports.Read.All` | Mailbox Report (mailbox usage report) |
| `Mail.Read` | Reserved for future mailbox-content checks |

3. After adding all of the above, click **Grant admin consent for `<tenant>`**. Each row should show *Granted for `<tenant>`* in the Status column.

## 5. Add Exchange Online application permission

1. **API permissions → Add a permission → APIs my organization uses**.
2. Search for and select **Office 365 Exchange Online**.
3. **Application permissions → `Exchange.ManageAsApp`**. Add it.
4. Click **Grant admin consent for `<tenant>`** again so the new Exchange permission is consented.

## 6. Assign the Exchange Administrator role to the app

`Exchange.ManageAsApp` only authorises the app to call Exchange Online; it does not grant any actual Exchange RBAC role. Without an Exchange admin role assignment, every EXO cmdlet returns `Error - Operation not allowed.`

1. In Entra admin center, go to **Identity → Roles & admins → Roles & administrators**.
2. Search for **Exchange Administrator**.
3. Click the role → **Add assignments → Select members**.
4. Search for the app registration name (e.g. *AroraMSP Audit Scripts*) and select it.
5. Click **Next** and confirm. The assignment is permanent (no PIM) for service principals.

> **Lower-privilege alternative**: For a read-only mailbox report, the Exchange role *Global Reader* is enough. The Tenant Audit script also reads mail-flow rules, transport config, anti-phish policies, etc., which are still readable under Global Reader. If you only ever run the mailbox report, assign **Global Reader** instead of **Exchange Administrator**.

## 7. Run the scripts

```powershell
$tenantId    = "00000000-0000-0000-0000-000000000000"
$clientId    = "11111111-1111-1111-1111-111111111111"
$thumbprint  = "AABBCCDDEEFF00112233445566778899AABBCCDD"

.\scripts\Invoke-AroraMSPTenantAudit.ps1 `
    -TenantId $tenantId `
    -ClientId $clientId `
    -CertificateThumbprint $thumbprint

.\scripts\Invoke-AroraMSPMailboxReport.ps1 `
    -TenantId $tenantId `
    -ClientId $clientId `
    -CertificateThumbprint $thumbprint
```

No browser prompts; both scripts authenticate silently via the certificate.

## Renewal

Self-signed certs created in step 1 expire after 2 years. To renew:

1. Run the same `New-SelfSignedCertificate` command (it generates a new key pair).
2. Export and upload the new `.cer` to the existing app registration. Both certificates can coexist — the script picks the one whose thumbprint you pass.
3. Once the new cert is verified working, remove the old cert from the app registration.
