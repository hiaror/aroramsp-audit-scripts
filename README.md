# AroraMSP Audit Scripts

## Overview
This repository contains a PowerShell automation suite that produces self-contained HTML reports for **Microsoft 365 tenant security** and **Exchange Online mailbox storage**. Both reports use AroraMSP branding, can be exported to PDF in-browser, and embed a CSV export of the underlying findings.

The scripts are designed for the engineer-in-the-field scenario: connect, run, hand the HTML to a stakeholder. No external dependencies in the output, no SaaS portal, no telemetry.

## Scripts

### Invoke-AroraMSPTenantAudit.ps1
Read-only Microsoft 365 security audit covering identity and access, email security, endpoint management, and licensing. Output is a single self-contained HTML report grouped by category, with PASS / FAIL / WARNING per check, recommendations, summary cards, a print stylesheet, and a CSV export button.

### Invoke-AroraMSPMailboxReport.ps1
Per-mailbox storage and archive report from Exchange Online and Microsoft Graph. Output is a sortable HTML table with summary cards, sorted by total size descending, with shared mailboxes that hold a paid licence highlighted in amber.

## Why This Exists

Off-the-shelf M365 audit tooling is either expensive, opinionated, or hidden behind a SaaS portal. These scripts produce something an engineer can hand straight to a stakeholder: a clean HTML report, brandable, printable, and exportable to CSV when the next person wants to slice the findings differently.

Both scripts are read-only and idempotent. They do not modify tenant configuration, mailbox content, or policy.

## Key Capabilities
- Microsoft Graph PowerShell SDK for tenant data (identity, devices, licensing)
- Exchange Online Management for mail flow, mailbox audit, and DKIM
- Self-contained HTML output (no external dependencies aside from Google Fonts)
- Print stylesheet flips to a clean light theme for PDF output
- CSV export embedded in every report (browser download via inlined JSON)
- Suitable for Global Reader / Mail Recipients roles — no write scopes requested

## Repository Structure
```
.
├── scripts/
│   ├── Invoke-AroraMSPTenantAudit.ps1
│   └── Invoke-AroraMSPMailboxReport.ps1
├── sample-data/
│   └── readme.md
├── .gitignore
└── README.md
```

## Prerequisites
- PowerShell 7.0 or later required. Run `pwsh` to launch PS7 after installing.
- Microsoft.Graph PowerShell module
- ExchangeOnlineManagement module
- Tenant Audit: Global Reader role minimum
- Mailbox Report: Mail Recipients role minimum, plus the `Reports.Read.All` Graph scope so the script can read the Microsoft 365 mailbox usage report (used to populate the Last Activity column)
- Outbound DNS lookups (DMARC and SPF checks resolve TXT records via Resolve-DnsName)
- **Note:** On some machines the WAM broker fails. Use the `-UseDeviceCode` switch if you hit authentication errors.

Install modules:
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Usage Examples

### Tenant Audit
```powershell
./scripts/Invoke-AroraMSPTenantAudit.ps1 -OutputDirectory C:\Reports
```
Connects to Microsoft Graph and Exchange Online interactively. Produces:
`C:\Reports\M365-Audit-Report-<TenantName>-<yyyy-MM-dd>.html`

Skip DNS lookups (DMARC and SPF) when running from a host without outbound DNS:
```powershell
./scripts/Invoke-AroraMSPTenantAudit.ps1 -SkipDnsChecks
```

### Mailbox Report
```powershell
./scripts/Invoke-AroraMSPMailboxReport.ps1 -OutputDirectory C:\Reports
```
Produces:
`C:\Reports\Mailbox-Report-<TenantName>-<yyyy-MM-dd>.html`

## Microsoft Graph Scopes

### Tenant Audit
- `Directory.Read.All`
- `Policy.Read.All`
- `Domain.Read.All`
- `Organization.Read.All`
- `User.Read.All`
- `Group.Read.All`
- `AuditLog.Read.All`
- `DeviceManagementApps.Read.All`
- `DeviceManagementConfiguration.Read.All`

### Mailbox Report
- `User.Read.All`
- `Organization.Read.All`
- `Directory.Read.All`
- `Reports.Read.All`

## Reporting

Both reports include:
- AroraMSP branded header with tenant name and run date
- Summary cards (counts and totals at the top)
- Print or save as PDF (browser print dialog, custom print stylesheet)
- Export findings as CSV (browser download via embedded JSON)

## Troubleshooting

### `Required module 'Microsoft.Graph' is not installed`
**Cause:** The Microsoft.Graph PowerShell module is not present on the machine.

**Fix:**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

**Note:** This module is large (~150 MB) and takes 3-5 minutes to install on a typical connection. The first `Connect-MgGraph` call after install may also take a few seconds while the SDK loads its sub-modules.

### `Required module 'ExchangeOnlineManagement' is not installed`
**Cause:** The ExchangeOnlineManagement module is not present on the machine.

**Fix:**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

### Script fails with `Object reference not set to an instance of an object` during authentication
**Cause:** WAM broker authentication fails on some Windows 11 machines with ExchangeOnlineManagement 3.x.

**Fix:** Run the script with the `-UseDeviceCode` switch:
```powershell
.\Invoke-AroraMSPMailboxReport.ps1 -UseDeviceCode
.\Invoke-AroraMSPTenantAudit.ps1 -UseDeviceCode
```
This prints a URL and code in the terminal. Open the URL in a browser, enter the code, sign in. No broker required.

### `Install-Module ExchangeOnlineManagement` fails with PackageManagement conflict
**Cause:** An older PackageManagement module is already loaded in the session.

**Fix:** Add `-AllowClobber` to the install command:
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
```

### Script fails with `Method not found: Microsoft.Identity.Client.PublicClientApplicationBuilder`
**Cause:** Running in Windows PowerShell 5.1. ExchangeOnlineManagement 3.x requires PS7.

**Fix:** Install PowerShell 7 from https://aka.ms/PSWindows, then open a new terminal, run `pwsh`, and reinstall both modules before running the script.

### `Get-MgOrganization` fails with `DeviceCodeCredential authentication failed`
**Cause:** Graph token acquired via device code does not persist correctly for certain cmdlets on some machines.

**Fix:** Already handled in the script. Tenant name falls back to `Unknown-Tenant` and the script continues. The report is still fully functional.

### Licence SKU enumeration fails with `DeviceCodeCredential authentication failed`
**Cause:** Same as above.

**Fix:** Already handled in the script. The Licence column shows blank and the script continues.

## Safety Notes

Both scripts are read-only. The HTML output may include user principal names, DNS records, and licence assignments — handle the report file accordingly.

DMARC and SPF DNS lookups query public DNS for each verified domain on the tenant. To skip the lookups, pass `-SkipDnsChecks` to the audit script.

## Disclaimer

Provided as-is for reference and learning purposes. No warranty. Test in a non-production tenant before relying on the output for compliance or audit deliverables.
