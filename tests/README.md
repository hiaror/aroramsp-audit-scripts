# AroraMSP M365 PowerShell Authentication Tests

## Overview
Two PowerShell scripts that reproduce the three authentication failures you hit when you try to run **Microsoft.Graph SDK** and **ExchangeOnlineManagement** in the same PowerShell session. Both scripts are frozen copies of the production tenant audit script taken at commit [`b77fcf2`](https://github.com/hiaror/aroramsp-audit-scripts/commit/b77fcf2), the last commit before certificate-based authentication was added.

The scripts are not part of the audit toolkit. They exist purely to demonstrate the failures, give a stable reproduction surface for documentation, and back up the test matrix in the blog post.

## Background

A full root-cause writeup of all three bugs is on the AroraMSP blog:
[**Microsoft 365 PowerShell authentication failures: three bugs, one fix**](https://aroramsp.com/blog/powershell-auth-clash.html)

Short version:
1. On PowerShell 5.1, the two modules ship incompatible MSAL assemblies. Whichever module loads first wins, and the second one fails with a missing-method error.
2. Reversing the connection order on PowerShell 5.1 changes the failure mode but not the outcome. The Graph SDK module cannot even initialise after Exchange Online has loaded its MSAL.
3. On PowerShell 7 with Microsoft.Graph v2.34 or later, both modules connect successfully, but the device-code-acquired Graph token is cached unusably and the next cmdlet fails with `DeviceCodeCredential authentication failed: Object reference not set to an instance of an object`. Confirmed SDK regression: [microsoftgraph/msgraph-sdk-powershell#3495](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3495).

## Test Matrix

All tests run on Windows 11, Microsoft.Graph PowerShell SDK v2.35.1, ExchangeOnlineManagement v3.9.2.

| # | Shell | Connection Order | Auth Method | EXO Login | Graph Login | EXO Cmdlets | Graph Cmdlets | Result |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | PS 7.6.1 | Graph first | Device code | Success | Success | Unknown | FAIL | **FAIL** |
| 2 | PS 7.6.1 | EXO first | Device code | Success | Success | Unknown | FAIL | **FAIL** |
| 3 | PS 5.1 | Graph first | Device code | FAIL | Success | N/A | N/A | **FAIL** |
| 4 | PS 5.1 | EXO first | Device code | Success | FAIL | N/A | N/A | **FAIL** |
| 5 | PS 5.1 | Graph first | Cert auth | Success | Success | N/A | N/A | **PASS** |
| 6 | PS 5.1 | EXO first | Cert auth | Success | FAIL | N/A | N/A | **FAIL** |
| 7 | PS 7.6.1 | Graph first | Cert auth | Success | Success | Success | Success | **PASS** |
| 8 | PS 7.6.1 | EXO first | Cert auth | Success | Success | Success | Success | **PASS** |

Three combinations pass end to end. All three use certificate based app-only authentication, which is what the production scripts in [`scripts/`](../scripts) use. Tests 7 and 8 (PowerShell 7 with cert auth, either order) are the recommended setup.

## Scripts

### Test-PreCertAuth.ps1
The b77fcf2 audit script, unchanged. Connects to **Microsoft Graph first**, then Exchange Online. Use this to reproduce:
- **Bug 1** (PS 5.1, device code): MSAL assembly clash on the second connect.
- **Bug 3** (PS 7, device code): Graph token cache failure on the first Graph cmdlet.

### Test-PreCertAuth-EXOFirst.ps1
Identical to `Test-PreCertAuth.ps1` except the Connect blocks are swapped: **Exchange Online connects first**, then Microsoft Graph. Use this to reproduce:
- **Bug 2** (PS 5.1, device code): Graph SDK module load failure after EXO has claimed MSAL.
- **Bug 3** (PS 7, device code): same token cache failure as above; connection order does not change the SDK regression.

Diff between the two files is exactly two `Write-Host` + `if/else` blocks reordered. Every other line is byte-identical.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+ (`pwsh`). Both shells reproduce different failures, which is the point.
- Microsoft.Graph PowerShell module **v2.34 or later** to reproduce Bug 3. Earlier versions (<= v2.33) are unaffected by the token cache regression.
- ExchangeOnlineManagement v3.x to reproduce Bug 1 and Bug 2 on PowerShell 5.1.
- A Microsoft 365 tenant the running user can authenticate against. The scripts do nothing destructive; they connect, run a single `Get-MgOrganization` query, and exit on the first cmdlet failure. Read-only.
- Outbound DNS lookups (the underlying script tries DMARC and SPF lookups; pass `-SkipDnsChecks` to skip them).

Install modules:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -RequiredVersion 2.35.1
Install-Module ExchangeOnlineManagement -Scope CurrentUser -RequiredVersion 3.9.2
```

## How to Run

The scripts inherit the parameter set of the production audit script: `-OutputDirectory`, `-TenantId`, `-SkipDnsChecks`, `-UseDeviceCode`. The b77fcf2 version pre-dates `-ClientId` and `-CertificateThumbprint`, so cert auth cannot be used with these test scripts directly. To compare cert auth against device code, run the production script in [`scripts/`](../scripts) with the `-CertificateThumbprint` parameter.

Reproduce Bug 3 (PowerShell 7, device code, Graph first):

```powershell
pwsh
.\tests\Test-PreCertAuth.ps1 `
    -TenantId "00000000-0000-0000-0000-000000000000" `
    -UseDeviceCode `
    -SkipDnsChecks
```

Expected: device code login completes in the browser; the script then fails with `DeviceCodeCredential authentication failed: Object reference not set to an instance of an object` on the first Graph cmdlet.

Reproduce Bug 1 (PowerShell 5.1, device code, Graph first):

```powershell
powershell
.\tests\Test-PreCertAuth.ps1 `
    -TenantId "00000000-0000-0000-0000-000000000000" `
    -UseDeviceCode `
    -SkipDnsChecks
```

Expected: Graph connects successfully; Exchange Online fails with `Method not found: 'Microsoft.Identity.Client.PublicClientApplicationBuilder ... WithBroker(...)'`.

Reproduce Bug 2 (PowerShell 5.1, device code, EXO first):

```powershell
powershell
.\tests\Test-PreCertAuth-EXOFirst.ps1 `
    -TenantId "00000000-0000-0000-0000-000000000000" `
    -UseDeviceCode `
    -SkipDnsChecks
```

Expected: Exchange Online connects successfully; Graph SDK fails with `The 'Connect-MgGraph' command was found in the module 'Microsoft.Graph.Authentication', but the module could not be loaded.`

## What Each Error Means

### `Method not found: 'Microsoft.Identity.Client.PublicClientApplicationBuilder ... WithBroker(...)'`
The two modules ship their own copies of MSAL (the Microsoft authentication library). When PowerShell 5.1 loads Microsoft.Graph first, Graph's MSAL gets pinned in memory. Then ExchangeOnlineManagement loads and calls a method (`WithBroker`) that exists in its own bundled MSAL but not in the version Graph put there first. Plain English: two modules brought different versions of the same toolbox; the second one cannot find the wrench it expected.

### `The 'Connect-MgGraph' command was found in the module 'Microsoft.Graph.Authentication', but the module could not be loaded.`
The reverse of the above. Exchange Online has already loaded its MSAL. When you import Microsoft.Graph.Authentication, the .NET runtime refuses to load Graph's MSAL because it conflicts with the one already in memory, and the Graph module cannot finish initialising. Plain English: the toolbox is locked by the previous tenant; the new tenant cannot even get in to set up.

### `DeviceCodeCredential authentication failed: Object reference not set to an instance of an object`
The login completes successfully (the browser flow works, the token is issued), but the SDK stores it as `null` somewhere in its cache. The next cmdlet pulls the token, finds nothing, and throws a `NullReferenceException` wrapped in a friendly DeviceCodeCredential message. Plain English: the badge was issued correctly at the front desk; the badge reader on the first internal door says the badge is empty. The bug is in the SDK, not in your tenant or in your network.

## On GitHub

These scripts and this README live at:
[**github.com/hiaror/aroramsp-audit-scripts/tree/main/tests**](https://github.com/hiaror/aroramsp-audit-scripts/tree/main/tests)

If you reproduce a different failure mode or find one of these passes on a configuration where the table predicts FAIL, open an issue on the repo with the shell version, module versions, and the exact error.
