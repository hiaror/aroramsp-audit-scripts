<#
.SYNOPSIS
  Microsoft 365 security audit. Read-only. Produces a self-contained HTML report.

.DESCRIPTION
  Connects to Microsoft Graph and Exchange Online (read-only) and runs a baseline
  set of identity, email, endpoint, and licensing checks against the current tenant.
  Emits a single self-contained HTML report with PASS / FAIL / WARNING per check,
  recommendations, AroraMSP branding, a print stylesheet, and an embedded CSV export.

.PARAMETER OutputDirectory
  Where to write the HTML report. Defaults to the current working directory.

.PARAMETER TenantId
  Optional. Tenant ID for Connect-MgGraph. If omitted, the user is prompted.

.PARAMETER SkipDnsChecks
  Skip DMARC and SPF DNS lookups. Useful from a host without outbound DNS.

.EXAMPLE
  ./Invoke-AroraMSPTenantAudit.ps1
  Runs the audit and writes the HTML report to the current directory.

.EXAMPLE
  ./Invoke-AroraMSPTenantAudit.ps1 -OutputDirectory C:\Reports
  Writes the report to C:\Reports.

.NOTES
  Author : AroraMSP - https://aroramsp.com
  Roles  : Global Reader (minimum)
  PS     : PowerShell 7.0 or later required. Use -UseDeviceCode switch if WAM broker authentication fails.
#>

[CmdletBinding()]
param(
    [string]$OutputDirectory = $PWD,
    [string]$TenantId,
    [switch]$SkipDnsChecks,
    [switch]$UseDeviceCode
)

if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "This script requires PowerShell 7 or later. Download from https://aka.ms/PSWindows, install it, then run pwsh to launch PS7 and try again."
}

$ErrorActionPreference = 'Continue'

# ----------------------------------------------------------------------------
# Module preflight
# ----------------------------------------------------------------------------
function Assert-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. Install with: Install-Module $Name -Scope CurrentUser"
    }
}
Assert-Module 'Microsoft.Graph'
Assert-Module 'ExchangeOnlineManagement'

# ----------------------------------------------------------------------------
# Connect
# ----------------------------------------------------------------------------
$graphScopes = @(
    'Directory.Read.All',
    'Policy.Read.All',
    'Domain.Read.All',
    'Organization.Read.All',
    'User.Read.All',
    'Group.Read.All',
    'AuditLog.Read.All',
    'DeviceManagementApps.Read.All',
    'DeviceManagementConfiguration.Read.All'
)

Write-Host '[+] Connecting to Microsoft Graph...' -ForegroundColor Cyan
if ($TenantId) {
    if ($UseDeviceCode) {
        Connect-MgGraph -Scopes $graphScopes -TenantId $TenantId -UseDeviceCode -NoWelcome -ErrorAction Stop
    } else {
        Connect-MgGraph -Scopes $graphScopes -TenantId $TenantId -NoWelcome -ErrorAction Stop
    }
} else {
    if ($UseDeviceCode) {
        Connect-MgGraph -Scopes $graphScopes -UseDeviceCode -NoWelcome -ErrorAction Stop
    } else {
        Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
    }
}

Write-Host '[+] Connecting to Exchange Online...' -ForegroundColor Cyan
if ($UseDeviceCode) {
    Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
} else {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}

# ----------------------------------------------------------------------------
# Tenant context
# ----------------------------------------------------------------------------
$org = Get-MgOrganization | Select-Object -First 1
$tenantName = $org.DisplayName
$tenantSafe = ($tenantName -replace '[^A-Za-z0-9\-]', '-') -replace '-+', '-'
$tenantSafe = $tenantSafe.Trim('-')
$today = Get-Date -Format 'yyyy-MM-dd'
$runYear = (Get-Date).Year
$reportPath = Join-Path $OutputDirectory ("M365-Audit-Report-{0}-{1}.html" -f $tenantSafe, $today)

Write-Host ("[+] Tenant: {0}" -f $tenantName) -ForegroundColor Green
Write-Host ("[+] Date  : {0}" -f $today) -ForegroundColor Green

# ----------------------------------------------------------------------------
# Findings collection
# ----------------------------------------------------------------------------
$findings = New-Object System.Collections.Generic.List[object]
function Add-Finding {
    param(
        [string]$Category,
        [string]$CheckName,
        [ValidateSet('PASS','FAIL','WARNING','INFO')][string]$Status,
        [string]$Detail = '',
        [string]$Recommendation = ''
    )
    $findings.Add([PSCustomObject]@{
        Category       = $Category
        CheckName      = $CheckName
        Status         = $Status
        Detail         = $Detail
        Recommendation = $Recommendation
    }) | Out-Null
}

# ============================================================================
# IDENTITY AND ACCESS
# ============================================================================
Write-Host '[*] Identity & access checks...' -ForegroundColor Yellow

$caPolicies = @()
try {
    $caPolicies = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
} catch {
    Add-Finding 'Identity & Access' 'Conditional Access policy enumeration' 'WARNING' "Failed to enumerate CA policies: $($_.Exception.Message)" 'Verify the Policy.Read.All scope was granted.'
}

# Legacy authentication blocked
try {
    $legacyBlock = $caPolicies | Where-Object {
        $_.State -eq 'enabled' -and
        $_.Conditions.ClientAppTypes -contains 'exchangeActiveSync' -and
        $_.Conditions.ClientAppTypes -contains 'other' -and
        $_.GrantControls.BuiltInControls -contains 'block'
    }
    if ($legacyBlock) {
        Add-Finding 'Identity & Access' 'Legacy authentication blocked via CA' 'PASS' "$($legacyBlock.Count) enabled policy/policies block legacy authentication."
    } else {
        Add-Finding 'Identity & Access' 'Legacy authentication blocked via CA' 'FAIL' 'No enabled Conditional Access policy blocks legacy authentication.' 'Create a CA policy targeting Exchange ActiveSync and Other clients with grant control = Block. Confirm zero successful legacy auth sign-ins for 30 days.'
    }
} catch {
    Add-Finding 'Identity & Access' 'Legacy authentication blocked via CA' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# Global Administrator count
try {
    $gaRole = Get-MgDirectoryRole -All | Where-Object DisplayName -eq 'Global Administrator' | Select-Object -First 1
    $gaCount = 0
    if ($gaRole) {
        $gaMembers = @(Get-MgDirectoryRoleMember -DirectoryRoleId $gaRole.Id -All)
        $gaCount = $gaMembers.Count
    }
    if ($gaCount -ge 2 -and $gaCount -le 4) {
        Add-Finding 'Identity & Access' 'Global Administrator count' 'PASS' "$gaCount Global Administrators (recommended range: 2-4)."
    } elseif ($gaCount -lt 2) {
        Add-Finding 'Identity & Access' 'Global Administrator count' 'WARNING' "$gaCount Global Administrator(s)." 'Maintain at least 2 for redundancy. Use PIM to keep them eligible-only.'
    } else {
        Add-Finding 'Identity & Access' 'Global Administrator count' 'WARNING' "$gaCount Global Administrators." 'Reduce to 2-4. Use PIM and lower-privilege roles for daily admin work.'
    }
} catch {
    Add-Finding 'Identity & Access' 'Global Administrator count' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# Break-glass accounts
try {
    $broadPolicies = @($caPolicies | Where-Object {
        $_.State -eq 'enabled' -and ($_.Conditions.Users.IncludeUsers -contains 'All')
    })
    if ($broadPolicies.Count -gt 0) {
        $excluded = @{}
        foreach ($p in $broadPolicies) {
            foreach ($u in $p.Conditions.Users.ExcludeUsers) {
                if ($u -and $u -ne 'None' -and $u -ne 'All') { $excluded[$u] = $true }
            }
            foreach ($g in $p.Conditions.Users.ExcludeGroups) {
                if ($g) { $excluded["group:$g"] = $true }
            }
        }
        $bgCount = $excluded.Keys.Count
        if ($bgCount -ge 2) {
            Add-Finding 'Identity & Access' 'Break-glass accounts excluded from CA' 'PASS' "$bgCount excluded principal(s) found across broad CA policies."
        } elseif ($bgCount -eq 1) {
            Add-Finding 'Identity & Access' 'Break-glass accounts excluded from CA' 'WARNING' '1 excluded principal found.' 'Maintain at least 2 break-glass accounts: cloud-only, excluded from all CA, sign-in alerts to a monitored channel.'
        } else {
            Add-Finding 'Identity & Access' 'Break-glass accounts excluded from CA' 'FAIL' 'No accounts or groups are excluded from broad CA policies.' 'Create 2 cloud-only break-glass accounts. Exclude them from all CA policies. Wire sign-in alerts to a SIEM.'
        }
    } else {
        Add-Finding 'Identity & Access' 'Break-glass accounts excluded from CA' 'WARNING' 'No broad (all users) CA policies to evaluate exclusions against.' 'Create a tenant-wide MFA CA policy first, then ensure break-glass accounts are excluded.'
    }
} catch {
    Add-Finding 'Identity & Access' 'Break-glass accounts excluded from CA' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# Tenant-wide MFA Conditional Access
try {
    $mfaPol = $caPolicies | Where-Object {
        $_.State -eq 'enabled' -and
        ($_.Conditions.Users.IncludeUsers -contains 'All') -and
        ($_.Conditions.Applications.IncludeApplications -contains 'All') -and
        ($_.GrantControls.BuiltInControls -contains 'mfa')
    }
    if ($mfaPol) {
        Add-Finding 'Identity & Access' 'Tenant-wide MFA Conditional Access policy' 'PASS' "$($mfaPol.Count) policy/policies require MFA for all users on all cloud apps."
    } else {
        Add-Finding 'Identity & Access' 'Tenant-wide MFA Conditional Access policy' 'FAIL' 'No enabled CA policy requires MFA for all users + all cloud apps.' 'Create a CA policy: include = All users + All cloud apps; grant = Require MFA. Exclude break-glass accounts only.'
    }
} catch {
    Add-Finding 'Identity & Access' 'Tenant-wide MFA Conditional Access policy' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# Self-service password reset (SSPR)
Add-Finding 'Identity & Access' 'Self-service password reset (SSPR) enabled' 'WARNING' 'SSPR enablement scope is not directly exposed via Microsoft Graph; manual verification required.' 'Verify in Entra admin center: Protection > Password reset > Properties (Self service password reset enabled = All).'

# ============================================================================
# EMAIL SECURITY
# ============================================================================
Write-Host '[*] Email security checks...' -ForegroundColor Yellow

$verifiedDomains = @()
try {
    $verifiedDomains = @((Get-MgDomain -All) | Where-Object { $_.IsVerified -and -not $_.Id.EndsWith('.onmicrosoft.com') } | ForEach-Object { $_.Id })
} catch {
    Add-Finding 'Email Security' 'Verified domain enumeration' 'WARNING' "Failed to enumerate verified domains: $($_.Exception.Message)"
}

# DMARC and SPF per verified domain
foreach ($d in $verifiedDomains) {
    if ($SkipDnsChecks) {
        Add-Finding 'Email Security' "DMARC: $d" 'INFO' 'Skipped (-SkipDnsChecks).'
        Add-Finding 'Email Security' "SPF: $d" 'INFO' 'Skipped (-SkipDnsChecks).'
        continue
    }
    # DMARC
    try {
        $rec = Resolve-DnsName -Name "_dmarc.$d" -Type TXT -ErrorAction Stop
        $dmarcTxt = $null
        foreach ($r in $rec) {
            if ($r.Strings) {
                $joined = ($r.Strings -join '')
                if ($joined -match 'v=DMARC1') { $dmarcTxt = $joined; break }
            }
        }
        if ($dmarcTxt) {
            $policy = if ($dmarcTxt -match 'p=(\w+)') { $matches[1].ToLower() } else { 'unknown' }
            switch ($policy) {
                'reject'     { Add-Finding 'Email Security' "DMARC: $d" 'PASS' 'p=reject (enforced)' }
                'quarantine' { Add-Finding 'Email Security' "DMARC: $d" 'WARNING' 'p=quarantine' 'Move to p=reject after monitoring rua aggregate reports for 30+ days.' }
                'none'       { Add-Finding 'Email Security' "DMARC: $d" 'WARNING' 'p=none (monitor-only)' 'Move to p=quarantine, then p=reject. Configure rua reporting to a mailbox or aggregator.' }
                default      { Add-Finding 'Email Security' "DMARC: $d" 'WARNING' "DMARC TXT present but policy is unclear ($policy)." 'Review and republish DMARC TXT.' }
            }
        } else {
            Add-Finding 'Email Security' "DMARC: $d" 'FAIL' 'No DMARC TXT record found.' "Publish: v=DMARC1; p=quarantine; rua=mailto:dmarc@$d"
        }
    } catch {
        Add-Finding 'Email Security' "DMARC: $d" 'FAIL' "DMARC lookup failed: $($_.Exception.Message)" "Publish: v=DMARC1; p=quarantine; rua=mailto:dmarc@$d"
    }
    # SPF
    try {
        $rec = Resolve-DnsName -Name $d -Type TXT -ErrorAction Stop
        $spfTxt = $null
        foreach ($r in $rec) {
            if ($r.Strings) {
                $joined = ($r.Strings -join '')
                if ($joined -match '^v=spf1') { $spfTxt = $joined; break }
            }
        }
        if ($spfTxt) {
            if ($spfTxt -match '-all\s*$') {
                Add-Finding 'Email Security' "SPF: $d" 'PASS' 'SPF -all (hard fail).'
            } elseif ($spfTxt -match '~all\s*$') {
                Add-Finding 'Email Security' "SPF: $d" 'WARNING' 'SPF ~all (soft fail).' 'Tighten to -all once all sending sources are confirmed in the SPF record.'
            } else {
                Add-Finding 'Email Security' "SPF: $d" 'WARNING' 'SPF present but lacks an -all or ~all qualifier.' 'Append -all to the SPF record.'
            }
        } else {
            Add-Finding 'Email Security' "SPF: $d" 'FAIL' 'No SPF TXT record found.' 'Publish: v=spf1 include:spf.protection.outlook.com -all (and any third-party sender includes).'
        }
    } catch {
        Add-Finding 'Email Security' "SPF: $d" 'FAIL' "SPF lookup failed: $($_.Exception.Message)" 'Publish: v=spf1 include:spf.protection.outlook.com -all'
    }
}

# DKIM
try {
    $dkim = Get-DkimSigningConfig -ErrorAction Stop
    foreach ($k in $dkim) {
        if ($k.Domain -and $k.Domain.EndsWith('.onmicrosoft.com')) { continue }
        if ($k.Enabled) {
            Add-Finding 'Email Security' "DKIM: $($k.Domain)" 'PASS' 'DKIM signing enabled on default selectors.'
        } else {
            Add-Finding 'Email Security' "DKIM: $($k.Domain)" 'FAIL' 'DKIM signing disabled.' "Publish CNAMEs and enable: Set-DkimSigningConfig -Identity '$($k.Domain)' -Enabled `$true"
        }
    }
} catch {
    Add-Finding 'Email Security' 'DKIM signing config' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# External auto-forwarding
try {
    $remoteDef = Get-RemoteDomain -Identity 'Default' -ErrorAction Stop
    $afRule = $null
    try { $afRule = Get-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue | Where-Object { $_.IsDefault } | Select-Object -First 1 } catch { }
    $remoteAllowed = ($remoteDef.AutoForwardEnabled -eq $true)
    $afAllowed = $false
    if ($afRule -and $afRule.AutoForwardingMode) {
        $afAllowed = ($afRule.AutoForwardingMode -ne 'Off')
    }
    if (-not $remoteAllowed -and -not $afAllowed) {
        Add-Finding 'Email Security' 'External auto-forwarding blocked' 'PASS' 'Default remote domain blocks auto-forwarding and outbound spam policy AutoForwardingMode is Off.'
    } elseif ($remoteAllowed -and $afAllowed) {
        Add-Finding 'Email Security' 'External auto-forwarding blocked' 'FAIL' 'Default remote domain AutoForwardEnabled = true and outbound spam AutoForwardingMode != Off.' "Set-RemoteDomain Default -AutoForwardEnabled `$false ; Set-HostedOutboundSpamFilterPolicy Default -AutoForwardingMode Off"
    } else {
        Add-Finding 'Email Security' 'External auto-forwarding blocked' 'WARNING' "remoteAllowed=$remoteAllowed; outboundSpamAutoForwardAllowed=$afAllowed" 'Block at both layers: remote domain AutoForwardEnabled = $false AND outbound spam policy AutoForwardingMode = Off.'
    }
} catch {
    Add-Finding 'Email Security' 'External auto-forwarding blocked' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# Mailbox audit logging (tenant)
try {
    $oc = Get-OrganizationConfig -ErrorAction Stop
    if ($oc.AuditDisabled -eq $false) {
        Add-Finding 'Email Security' 'Mailbox audit logging (tenant)' 'PASS' 'Tenant-level mailbox auditing enabled (default since 2019).'
    } else {
        Add-Finding 'Email Security' 'Mailbox audit logging (tenant)' 'FAIL' 'AuditDisabled = $true' 'Set-OrganizationConfig -AuditDisabled $false'
    }
} catch {
    Add-Finding 'Email Security' 'Mailbox audit logging (tenant)' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# Shared mailboxes - sign-in & paid licences
try {
    $sharedMbx = @(Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
    $signInEnabled = New-Object System.Collections.Generic.List[string]
    $licensed = New-Object System.Collections.Generic.List[string]
    foreach ($m in $sharedMbx) {
        if (-not $m.ExternalDirectoryObjectId) { continue }
        try {
            $u = Get-MgUser -UserId $m.ExternalDirectoryObjectId -Property 'AccountEnabled,AssignedLicenses,UserPrincipalName' -ErrorAction Stop
            if ($u.AccountEnabled) { $signInEnabled.Add($u.UserPrincipalName) | Out-Null }
            if ($u.AssignedLicenses -and $u.AssignedLicenses.Count -gt 0) { $licensed.Add($u.UserPrincipalName) | Out-Null }
        } catch { }
    }
    if ($signInEnabled.Count -eq 0) {
        Add-Finding 'Email Security' 'Shared mailboxes: sign-in disabled' 'PASS' "All $($sharedMbx.Count) shared mailbox(es) have sign-in disabled."
    } else {
        $names = $signInEnabled -join ', '
        if ($names.Length -gt 240) { $names = $names.Substring(0, 237) + '...' }
        Add-Finding 'Email Security' 'Shared mailboxes: sign-in disabled' 'FAIL' "$($signInEnabled.Count) shared mailbox(es) with sign-in enabled: $names" 'Disable interactive sign-in: Update-MgUser -UserId <id> -AccountEnabled:$false. Preserve mailbox access via Full Access permission.'
    }
    if ($licensed.Count -eq 0) {
        Add-Finding 'Licensing' 'Shared mailboxes: no paid licences' 'PASS' 'No shared mailboxes carry an assigned licence.'
    } else {
        $names = $licensed -join ', '
        if ($names.Length -gt 240) { $names = $names.Substring(0, 237) + '...' }
        Add-Finding 'Licensing' 'Shared mailboxes: no paid licences' 'WARNING' "$($licensed.Count) shared mailbox(es) carry licences: $names" 'Shared mailboxes <50 GB do not require a licence. Reclaim unless archive or litigation hold requires Exchange Online Plan 2.'
    }
} catch {
    Add-Finding 'Email Security' 'Shared mailbox checks' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# ============================================================================
# ENDPOINT
# ============================================================================
Write-Host '[*] Endpoint checks...' -ForegroundColor Yellow

# Compliance policies
try {
    $compliance = @(Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction Stop)
    if ($compliance.Count -gt 0) {
        Add-Finding 'Endpoint' 'Intune compliance policies deployed' 'PASS' "$($compliance.Count) compliance policy/policies deployed."
    } else {
        Add-Finding 'Endpoint' 'Intune compliance policies deployed' 'FAIL' 'No compliance policies found.' 'Deploy compliance policies for each managed platform; wire to Conditional Access (Require compliant device).'
    }
} catch {
    Add-Finding 'Endpoint' 'Intune compliance policies deployed' 'WARNING' "Check failed: $($_.Exception.Message)" 'Verify the DeviceManagementConfiguration.Read.All scope.'
}

# App protection policies (MAM)
try {
    $iosApp = @(Get-MgDeviceAppManagementiOSManagedAppProtection -All -ErrorAction SilentlyContinue)
    $androidApp = @(Get-MgDeviceAppManagementAndroidManagedAppProtection -All -ErrorAction SilentlyContinue)
    $count = $iosApp.Count + $androidApp.Count
    if ($count -gt 0) {
        Add-Finding 'Endpoint' 'App protection policies deployed' 'PASS' "$count MAM policy/policies deployed (iOS: $($iosApp.Count), Android: $($androidApp.Count))."
    } else {
        Add-Finding 'Endpoint' 'App protection policies deployed' 'WARNING' 'No App Protection (MAM) policies found.' 'Deploy MAM policies for managed and unmanaged iOS and Android. Block copy-out, screenshot, backup, and require PIN.'
    }
} catch {
    Add-Finding 'Endpoint' 'App protection policies deployed' 'WARNING' "Check failed: $($_.Exception.Message)" 'Verify the DeviceManagementApps.Read.All scope.'
}

# ============================================================================
# LICENSING
# ============================================================================
Write-Host '[*] Licensing checks...' -ForegroundColor Yellow

try {
    $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
    $unassigned = @($skus | Where-Object { $_.PrepaidUnits.Enabled -gt $_.ConsumedUnits } | ForEach-Object {
        $diff = $_.PrepaidUnits.Enabled - $_.ConsumedUnits
        '{0}: {1} of {2} unassigned' -f $_.SkuPartNumber, $diff, $_.PrepaidUnits.Enabled
    })
    if ($unassigned.Count -eq 0) {
        Add-Finding 'Licensing' 'Unassigned licences' 'PASS' 'All purchased licences are consumed.'
    } else {
        Add-Finding 'Licensing' 'Unassigned licences' 'WARNING' ($unassigned -join '; ') 'Reclaim or downsize unused licences at the next renewal cycle.'
    }
} catch {
    Add-Finding 'Licensing' 'Unassigned licences' 'WARNING' "Check failed: $($_.Exception.Message)"
}

# Inactive licensed users (90 days)
try {
    $cutoff = (Get-Date).AddDays(-90)
    $inactive = @(Get-MgUser -All -Property 'signInActivity,assignedLicenses,displayName,userPrincipalName,accountEnabled' -ErrorAction Stop | Where-Object {
        $_.AccountEnabled -eq $true -and
        $_.AssignedLicenses -and $_.AssignedLicenses.Count -gt 0 -and
        ($null -eq $_.SignInActivity -or $null -eq $_.SignInActivity.LastSignInDateTime -or $_.SignInActivity.LastSignInDateTime -lt $cutoff)
    })
    if ($inactive.Count -eq 0) {
        Add-Finding 'Licensing' 'Inactive licensed users (90 days)' 'PASS' 'No licensed users with no sign-in in 90 days.'
    } else {
        Add-Finding 'Licensing' 'Inactive licensed users (90 days)' 'WARNING' "$($inactive.Count) licensed user(s) with no sign-in in 90 days." 'Review and reclaim licences. Disable accounts where the user has left.'
    }
} catch {
    Add-Finding 'Licensing' 'Inactive licensed users (90 days)' 'WARNING' "Check failed: $($_.Exception.Message)" 'Sign-in activity requires AuditLog.Read.All scope and Microsoft Entra ID P1 or higher.'
}

# ============================================================================
# Build HTML report
# ============================================================================
Write-Host '[+] Building HTML report...' -ForegroundColor Cyan

$passCount  = @($findings | Where-Object Status -eq 'PASS').Count
$failCount  = @($findings | Where-Object Status -eq 'FAIL').Count
$warnCount  = @($findings | Where-Object Status -eq 'WARNING').Count
$totalCount = $findings.Count

function ConvertTo-HtmlEntity {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $s = $s -replace '&', '&amp;'
    $s = $s -replace '<', '&lt;'
    $s = $s -replace '>', '&gt;'
    $s = $s -replace '"', '&quot;'
    $s = $s -replace "'", '&#39;'
    return $s
}

# Group findings by category, preserving insertion order
$categoryOrder = New-Object System.Collections.Generic.List[string]
$categoryMap = @{}
foreach ($f in $findings) {
    if (-not $categoryMap.ContainsKey($f.Category)) {
        $categoryMap[$f.Category] = New-Object System.Collections.Generic.List[object]
        $categoryOrder.Add($f.Category) | Out-Null
    }
    $categoryMap[$f.Category].Add($f) | Out-Null
}

$findingsHtml = New-Object System.Text.StringBuilder
$catIdx = 0
foreach ($catName in $categoryOrder) {
    $catIdx++
    $catNumStr = '{0:D2}' -f $catIdx
    [void]$findingsHtml.AppendLine('<section class="category">')
    [void]$findingsHtml.AppendLine(('  <header class="cat-head"><span class="cat-num">{0}</span><h2>{1}</h2></header>' -f $catNumStr, (ConvertTo-HtmlEntity $catName)))
    foreach ($f in $categoryMap[$catName]) {
        $cls = $f.Status.ToLower()
        [void]$findingsHtml.AppendLine(('  <article class="finding {0}">' -f $cls))
        [void]$findingsHtml.AppendLine(('    <div class="status {0}">{1}</div>' -f $cls, $f.Status))
        [void]$findingsHtml.AppendLine('    <div class="body">')
        [void]$findingsHtml.AppendLine(('      <h3>{0}</h3>' -f (ConvertTo-HtmlEntity $f.CheckName)))
        if ($f.Detail) {
            [void]$findingsHtml.AppendLine(('      <p class="detail">{0}</p>' -f (ConvertTo-HtmlEntity $f.Detail)))
        }
        if ($f.Recommendation) {
            [void]$findingsHtml.AppendLine(('      <p class="rec"><span class="rec-lbl">recommendation</span> {0}</p>' -f (ConvertTo-HtmlEntity $f.Recommendation)))
        }
        [void]$findingsHtml.AppendLine('    </div>')
        [void]$findingsHtml.AppendLine('  </article>')
    }
    [void]$findingsHtml.AppendLine('</section>')
}

# Findings JSON for the CSV button
$findingsJson = ($findings | ConvertTo-Json -Compress -Depth 5)
if (-not $findingsJson) { $findingsJson = '[]' }
elseif ($findingsJson -notmatch '^\[') { $findingsJson = "[$findingsJson]" }
$findingsJson = $findingsJson -replace '</', '<\/'

# HTML template (single-quoted here-string -- no PS interpolation inside)
$template = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>M365 Audit Report - {{TENANT_NAME}} - {{DATE}}</title>
<meta name="description" content="Microsoft 365 security audit report for {{TENANT_NAME}}, generated {{DATE}} by AroraMSP." />
<link rel="preconnect" href="https://fonts.googleapis.com" />
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet" />
<style>
:root {
  --bg:#0f1117; --bg-1:#12151d; --bg-2:#171a23; --bg-3:#1d212c;
  --line:#242936; --line-2:#2e3344;
  --fg:#e7eaf0; --fg-dim:#e5e7eb; --fg-mute:#cbd5e1;
  --accent:#7c6ff7; --accent-soft:rgba(124,111,247,.18); --accent-line:rgba(124,111,247,.4);
  --cyan:#4ecdc4; --cyan-soft:rgba(78,205,196,.18); --cyan-line:rgba(78,205,196,.4);
  --good:#10b981; --good-soft:rgba(16,185,129,.16);
  --warn:#f59e0b; --warn-soft:rgba(245,158,11,.16);
  --red:#ef4444;  --red-soft:rgba(239,68,68,.16);
  --sans:'Inter',-apple-system,BlinkMacSystemFont,sans-serif;
  --mono:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,monospace;
  --maxw:1100px; --pad:clamp(20px,4vw,40px);
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:var(--bg);color:var(--fg);font-family:var(--sans);font-size:15px;line-height:1.55;-webkit-font-smoothing:antialiased;display:flex;flex-direction:column;min-height:100vh}
a{color:inherit;text-decoration:none}
.wrap{max-width:var(--maxw);margin:0 auto;padding:0 var(--pad);position:relative;z-index:1;width:100%}
.grid-bg{position:fixed;inset:0;pointer-events:none;z-index:0;background-image:linear-gradient(var(--line) 1px,transparent 1px),linear-gradient(90deg,var(--line) 1px,transparent 1px);background-size:64px 64px;mask-image:radial-gradient(ellipse at 50% 0%,black 0%,transparent 70%);opacity:.35}

.top{position:sticky;top:0;z-index:50;backdrop-filter:saturate(140%) blur(10px);background:color-mix(in oklab,var(--bg) 78%,transparent);border-bottom:1px solid var(--line)}
.top-inner{display:flex;align-items:center;justify-content:space-between;padding:14px 0;gap:16px;flex-wrap:wrap}
.brand{display:flex;align-items:center;gap:10px;color:var(--fg)}
.brand svg{display:block;height:36px;width:auto;flex-shrink:0}
.meta{font-family:var(--mono);font-size:12px;text-align:right}
.meta .dim{color:var(--fg-mute);font-size:11px}

main{flex:1;display:flex;flex-direction:column;position:relative;z-index:1}
.page{padding:clamp(28px,4vw,48px) 0 clamp(40px,5vw,64px)}
.page-head{margin-bottom:24px;max-width:760px}
.sec-label{font-family:var(--mono);font-size:12px;color:var(--cyan);display:inline-flex;align-items:center;gap:8px;margin-bottom:12px}
.sec-label::before{content:"";width:18px;height:1px;background:var(--cyan)}
.page-head h1{font-size:clamp(28px,4vw,40px);margin:0 0 12px;letter-spacing:-.02em;line-height:1.1;font-weight:600}
.page-head p{color:var(--fg-dim);margin:0;font-size:15px;max-width:64ch}

.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px}
@media (max-width:640px){.summary{grid-template-columns:repeat(2,1fr)}}
.sum-card{background:var(--bg-1);border:1px solid var(--line);border-radius:10px;padding:18px 20px}
.sum-card .num{font-family:var(--mono);font-size:28px;font-weight:600;line-height:1}
.sum-card .lbl{font-family:var(--mono);font-size:11px;color:var(--fg-mute);text-transform:uppercase;letter-spacing:.08em;margin-top:8px}
.sum-card.pass .num{color:var(--good)}
.sum-card.fail .num{color:var(--red)}
.sum-card.warn .num{color:var(--warn)}
.sum-card.total .num{color:var(--cyan)}

.actions{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:28px}
.btn{font-family:var(--mono);font-size:13px;font-weight:500;padding:10px 16px;border-radius:8px;cursor:pointer;display:inline-flex;align-items:center;gap:8px;transition:filter .15s,background .2s,border-color .2s;border:1px solid transparent}
.btn.primary{background:var(--accent);color:white;border-color:color-mix(in oklab,white 12%,var(--accent))}
.btn.primary:hover{filter:brightness(1.1)}
.btn.ghost{background:var(--bg-1);color:var(--fg);border-color:var(--line-2)}
.btn.ghost:hover{background:var(--bg-2)}
.btn svg{width:14px;height:14px}

.category{margin-bottom:22px}
.cat-head{display:flex;align-items:center;gap:14px;margin-bottom:12px}
.cat-num{font-family:var(--mono);font-size:11px;color:var(--cyan);border:1px solid var(--cyan-line);border-radius:6px;padding:3px 8px;letter-spacing:.06em}
.cat-head h2{margin:0;font-size:18px;font-weight:600;letter-spacing:-.01em}

.finding{background:var(--bg-1);border:1px solid var(--line);border-radius:10px;padding:14px 18px;margin-bottom:10px;display:grid;grid-template-columns:90px 1fr;gap:18px;align-items:start;break-inside:avoid}
.finding.fail{border-left:3px solid var(--red)}
.finding.warning{border-left:3px solid var(--warn)}
.finding.pass{border-left:3px solid var(--good)}
.finding.info{border-left:3px solid var(--cyan)}
.finding .status{font-family:var(--mono);font-size:11px;font-weight:600;letter-spacing:.06em;display:inline-flex;align-items:center;justify-content:center;padding:6px 0;border-radius:6px;height:fit-content;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.finding .status.pass{background:var(--good-soft);color:var(--good)}
.finding .status.fail{background:var(--red-soft);color:var(--red)}
.finding .status.warning{background:var(--warn-soft);color:var(--warn)}
.finding .status.info{background:var(--cyan-soft);color:var(--cyan)}
.finding .body h3{margin:0 0 6px;font-size:14.5px;font-weight:600;letter-spacing:-.01em;line-height:1.4}
.finding .body .detail{margin:0;font-family:var(--mono);font-size:12px;color:var(--fg-dim);line-height:1.55;word-break:break-word}
.finding .body .rec{margin:8px 0 0;font-size:13px;color:var(--fg-mute);line-height:1.5}
.finding .body .rec-lbl{font-family:var(--mono);font-size:10px;color:var(--accent);text-transform:uppercase;letter-spacing:.08em;margin-right:6px}

footer{border-top:1px solid var(--line);padding:22px 0 20px;font-family:var(--mono);font-size:12px;color:var(--fg-mute);margin-top:auto;position:relative;z-index:1}
.foot-top{display:flex;flex-wrap:wrap;gap:18px;justify-content:space-between;align-items:center;padding-bottom:14px;margin-bottom:14px;border-bottom:1px dashed var(--line)}
.foot-top a:hover{color:var(--fg)}
.foot-bottom{text-align:center;font-size:11px;letter-spacing:.06em;text-transform:uppercase}
.foot-bottom .dot{color:var(--good)}

@media (max-width:640px){
  .finding{grid-template-columns:1fr;gap:8px}
  .finding .status{width:fit-content;padding:4px 10px}
}

@media print{
  @page{margin:14mm 12mm}
  :root{
    --bg:#fff; --bg-1:#fff; --bg-2:#f7f8fa; --bg-3:#eef0f4;
    --line:#d4d8e0; --line-2:#b8bdc7;
    --fg:#0f1117; --fg-dim:#1f2430; --fg-mute:#4b5563;
    --accent:#5b50d6; --cyan:#0e7490;
    --good:#047857; --warn:#b45309; --red:#b91c1c;
    --good-soft:#d1fae5; --warn-soft:#fef3c7; --red-soft:#fee2e2; --cyan-soft:#cffafe;
  }
  body{background:#fff;color:#0f1117;font-size:10.5pt;line-height:1.4}
  .grid-bg,.top,.actions{display:none !important}
  main,.page{padding:0}
  .wrap{max-width:100%;padding:0}
  .page-head h1{font-size:20pt}
  .summary{gap:8px;margin-bottom:14px}
  .sum-card{border:1px solid #d4d8e0;padding:10px 12px;break-inside:avoid}
  .sum-card .num{font-size:18pt}
  .category{break-inside:avoid-page;margin-bottom:12px}
  .finding{break-inside:avoid;border:1px solid #d4d8e0;padding:10px 14px;margin-bottom:6px}
  .finding .status{padding:4px 8px}
  footer{border-top:1px solid #d4d8e0;padding:10px 0}
  .foot-top a{display:none}
  -webkit-print-color-adjust:exact;print-color-adjust:exact
}
</style>
</head>
<body>
<div class="grid-bg" aria-hidden="true"></div>

<header class="top">
  <div class="wrap top-inner">
    <a class="brand" href="https://aroramsp.com" target="_blank" rel="noopener" aria-label="AroraMSP">
      <svg viewBox="55 45 700 165" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <g transform="translate(55, 50)">
          <rect x="0" y="0" width="90" height="90" rx="18" fill="#1a1f2e"/>
          <rect x="18" y="38" width="54" height="14" rx="3" fill="none" stroke="#7c6ff7" stroke-width="1.5"/>
          <line x1="25" y1="38" x2="25" y2="28" stroke="#7c6ff7" stroke-width="1.5"/>
          <circle cx="25" cy="24" r="4" fill="none" stroke="#7c6ff7" stroke-width="1.5"/>
          <line x1="38" y1="38" x2="38" y2="22" stroke="#4ecdc4" stroke-width="1.5"/>
          <rect x="34" y="16" width="8" height="6" rx="1" fill="#4ecdc4"/>
          <line x1="52" y1="38" x2="52" y2="28" stroke="#7c6ff7" stroke-width="1.5"/>
          <circle cx="52" cy="24" r="4" fill="none" stroke="#7c6ff7" stroke-width="1.5"/>
          <line x1="65" y1="38" x2="65" y2="22" stroke="#4ecdc4" stroke-width="1.5"/>
          <rect x="61" y="16" width="8" height="6" rx="1" fill="#4ecdc4"/>
          <line x1="25" y1="52" x2="25" y2="62" stroke="#7c6ff7" stroke-width="1.5"/>
          <line x1="25" y1="62" x2="38" y2="62" stroke="#7c6ff7" stroke-width="1.5"/>
          <line x1="38" y1="62" x2="38" y2="70" stroke="#7c6ff7" stroke-width="1.5"/>
          <circle cx="38" cy="73" r="3" fill="#7c6ff7"/>
          <line x1="52" y1="52" x2="52" y2="62" stroke="#4ecdc4" stroke-width="1.5"/>
          <line x1="52" y1="62" x2="65" y2="62" stroke="#4ecdc4" stroke-width="1.5"/>
          <line x1="65" y1="62" x2="65" y2="70" stroke="#4ecdc4" stroke-width="1.5"/>
          <circle cx="65" cy="73" r="3" fill="#4ecdc4"/>
          <circle cx="25" cy="45" r="2" fill="#ffffff"/>
          <circle cx="38" cy="45" r="2" fill="#ffffff"/>
          <circle cx="52" cy="45" r="2" fill="#ffffff"/>
          <circle cx="65" cy="45" r="2" fill="#ffffff"/>
        </g>
        <text x="170" y="95" font-family="Arial,sans-serif" font-size="42" font-weight="700" fill="#ffffff" letter-spacing="-1">Arora</text>
        <text x="170" y="138" font-family="Arial,sans-serif" font-size="42" font-weight="300" fill="#7c6ff7" letter-spacing="2">MSP</text>
        <line x1="170" y1="150" x2="640" y2="150" stroke="#ffffff" stroke-width="0.5" opacity="0.2"/>
        <text x="170" y="172" font-family="Arial,sans-serif" font-size="25" fill="#ffffff" letter-spacing="3">MICROSOFT 365 CONSULTING</text>
      </svg>
    </a>
    <div class="meta">
      <div>{{TENANT_NAME}}</div>
      <div class="dim">{{DATE}}</div>
    </div>
  </div>
</header>

<main>
  <section class="page wrap">
    <div class="page-head">
      <div class="sec-label">// microsoft 365 security audit · v1.0</div>
      <h1>Tenant audit report.</h1>
      <p>Read-only baseline audit of identity, email security, endpoint, and licensing posture for the {{TENANT_NAME}} tenant. Findings below are grouped by category, with status, detail, and recommendation per check.</p>
    </div>

    <div class="summary">
      <div class="sum-card pass"><div class="num">{{SUMMARY_PASS}}</div><div class="lbl">PASS</div></div>
      <div class="sum-card fail"><div class="num">{{SUMMARY_FAIL}}</div><div class="lbl">FAIL</div></div>
      <div class="sum-card warn"><div class="num">{{SUMMARY_WARN}}</div><div class="lbl">WARNING</div></div>
      <div class="sum-card total"><div class="num">{{SUMMARY_TOTAL}}</div><div class="lbl">TOTAL CHECKS</div></div>
    </div>

    <div class="actions">
      <button class="btn primary" id="btnPrint" type="button">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
        Export to PDF
      </button>
      <button class="btn ghost" id="btnCsv" type="button">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Export to CSV
      </button>
    </div>

{{FINDINGS_HTML}}
  </section>
</main>

<footer>
  <div class="wrap">
    <div class="foot-top">
      <div>© {{YEAR}} AroraMSP · Microsoft 365 consulting</div>
      <a href="https://aroramsp.com" target="_blank" rel="noopener">aroramsp.com →</a>
    </div>
    <div class="foot-bottom"><span class="dot">●</span> M365 audit report · v1.0</div>
  </div>
</footer>

<script>
const FINDINGS = {{FINDINGS_JSON}};
const TENANT = "{{TENANT_NAME_JS}}";
const DATE = "{{DATE}}";
function csvEscape(v){ return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"'; }
function downloadCsv(){
  const headers = ['Category','CheckName','Status','Detail','Recommendation'];
  const rows = [headers.join(',')];
  for (const f of FINDINGS) rows.push(headers.map(h => csvEscape(f[h])).join(','));
  const blob = new Blob(['﻿' + rows.join('\r\n')], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const safe = (TENANT + '-' + DATE).replace(/[^A-Za-z0-9._-]/g, '_');
  a.download = 'M365-Audit-' + safe + '.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}
document.getElementById('btnPrint').addEventListener('click', () => window.print());
document.getElementById('btnCsv').addEventListener('click', downloadCsv);
</script>
</body>
</html>
'@

# JS string-safe tenant name
$tenantJs = $tenantName -replace '\\', '\\\\' -replace '"', '\"'

$html = $template
$html = $html.Replace('{{TENANT_NAME}}',    (ConvertTo-HtmlEntity $tenantName))
$html = $html.Replace('{{TENANT_NAME_JS}}', $tenantJs)
$html = $html.Replace('{{DATE}}',           $today)
$html = $html.Replace('{{YEAR}}',           "$runYear")
$html = $html.Replace('{{SUMMARY_PASS}}',   "$passCount")
$html = $html.Replace('{{SUMMARY_FAIL}}',   "$failCount")
$html = $html.Replace('{{SUMMARY_WARN}}',   "$warnCount")
$html = $html.Replace('{{SUMMARY_TOTAL}}',  "$totalCount")
$html = $html.Replace('{{FINDINGS_HTML}}',  $findingsHtml.ToString())
$html = $html.Replace('{{FINDINGS_JSON}}',  $findingsJson)

# Write file (UTF-8 without BOM)
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($reportPath, $html, $utf8NoBom)

Write-Host ''
Write-Host '[+] Audit complete.' -ForegroundColor Green
Write-Host ('     PASS    : {0}' -f $passCount) -ForegroundColor Green
Write-Host ('     FAIL    : {0}' -f $failCount) -ForegroundColor Red
Write-Host ('     WARNING : {0}' -f $warnCount) -ForegroundColor Yellow
Write-Host ('     TOTAL   : {0}' -f $totalCount) -ForegroundColor Cyan
Write-Host ''
Write-Host ("[+] Report saved to: {0}" -f $reportPath) -ForegroundColor Cyan
