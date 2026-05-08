<#
.SYNOPSIS
  Mailbox storage and archive report. Read-only. Produces a self-contained HTML report.

.DESCRIPTION
  Pulls per-mailbox primary and archive size, last logon, and licence SKU
  from Exchange Online and Microsoft Graph. Output is a single self-contained
  HTML report with summary cards, a sortable table sorted by Total GB descending,
  shared mailboxes that hold a paid licence highlighted in amber, and an
  embedded CSV export.

.PARAMETER OutputDirectory
  Where to write the HTML report. Defaults to the current working directory.

.PARAMETER TenantId
  Optional. Tenant ID for Connect-MgGraph.

.EXAMPLE
  ./Invoke-AroraMSPMailboxReport.ps1
  Connects, builds the report, and writes the HTML to the current directory.

.EXAMPLE
  ./Invoke-AroraMSPMailboxReport.ps1 -OutputDirectory C:\Reports

.NOTES
  Author : AroraMSP - https://aroramsp.com
  Roles  : Mail Recipients (minimum)
  PS     : PowerShell 7.0 or later required. Use -UseDeviceCode switch if WAM broker authentication fails.
#>

[CmdletBinding()]
param(
    [string]$OutputDirectory = $PWD,
    [string]$TenantId = "",
    [string]$ClientId = "",
    [string]$CertificateThumbprint = "",
    [switch]$UseDeviceCode
)

if ($UseDeviceCode -and $PSVersionTable.PSVersion.Major -lt 7) {
    throw "Device code authentication requires PowerShell 7 or later.
Download from https://aka.ms/PSWindows or use certificate authentication
which works on PowerShell 5 and later."
}

if (-not $UseDeviceCode) {
    if ([string]::IsNullOrWhiteSpace($TenantId) -or
        [string]::IsNullOrWhiteSpace($ClientId) -or
        [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        throw "Certificate authentication requires -TenantId, -ClientId, and -CertificateThumbprint. Alternatively use -UseDeviceCode for interactive login."
    }
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
    'User.Read.All',
    'Organization.Read.All',
    'Directory.Read.All'
)

Write-Host '[+] Connecting to Microsoft Graph...' -ForegroundColor Cyan
if ($UseDeviceCode) {
    if ($TenantId) {
        Connect-MgGraph -Scopes $graphScopes -TenantId $TenantId -UseDeviceCode -NoWelcome -ErrorAction Stop
    } else {
        Connect-MgGraph -Scopes $graphScopes -UseDeviceCode -NoWelcome -ErrorAction Stop
    }
} else {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
}

# App-only Connect-ExchangeOnline expects the primary verified domain, not the tenant GUID.
if (-not $UseDeviceCode) {
    $orgDomain = (Get-MgOrganization | Select-Object -First 1).VerifiedDomains |
        Where-Object { $_.IsDefault } | Select-Object -ExpandProperty Name
}

Write-Host '[+] Connecting to Exchange Online...' -ForegroundColor Cyan
if ($UseDeviceCode) {
    Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
} else {
    Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $orgDomain -ShowBanner:$false -ErrorAction Stop
}

# ----------------------------------------------------------------------------
# Tenant context
# ----------------------------------------------------------------------------
try {
    $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    $tenantName = $org.DisplayName
} catch {
    $tenantName = "Unknown-Tenant"
    Write-Warning "[!] Could not retrieve tenant name: $_"
}
if ($tenantName) { $tenantName = $tenantName.Trim('-') }
$tenantSafe = ($tenantName -replace '[^A-Za-z0-9\-]', '-') -replace '-+', '-'
$tenantSafe = $tenantSafe.Trim('-')
if ([string]::IsNullOrWhiteSpace($tenantSafe)) { $tenantSafe = 'Unknown-Tenant' }
$today = Get-Date -Format 'yyyy-MM-dd'
$runYear = (Get-Date).Year
$reportPath = Join-Path $OutputDirectory ("Mailbox-Report-{0}-{1}.html" -f $tenantSafe, $today)

Write-Host ("[+] Tenant: {0}" -f $tenantName) -ForegroundColor Green

# ----------------------------------------------------------------------------
# SKU lookup map (Graph SubscribedSku)
# ----------------------------------------------------------------------------
$skuMap = @{}
try {
    $skus = Get-MgSubscribedSku -All -ErrorAction Stop
    foreach ($s in $skus) { $skuMap[[string]$s.SkuId] = $s.SkuPartNumber }
} catch {
    $skuMap = @{}
    Write-Warning "[!] Failed to enumerate licence SKUs: $($_.Exception.Message)"
}

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
function Get-SizeBytes {
    param($itemSize)
    if (-not $itemSize) { return 0 }
    if ($itemSize.Value -and $itemSize.Value.PSObject.Methods['ToBytes']) {
        try { return [double]$itemSize.Value.ToBytes() } catch { }
    }
    $s = "$itemSize"
    if ($s -match '\(([\d,]+)\s+bytes\)') {
        $b = $matches[1] -replace ',', ''
        if ($b -as [double]) { return [double]$b }
    }
    return 0
}

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

# ----------------------------------------------------------------------------
# Pull mailboxes
# ----------------------------------------------------------------------------
Write-Host '[+] Enumerating mailboxes...' -ForegroundColor Cyan
$allMbx = @(Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName,UserPrincipalName,RecipientTypeDetails,ArchiveStatus,ArchiveDatabase,ExternalDirectoryObjectId,ProhibitSendReceiveQuota -ErrorAction Stop)
Write-Host ("[+] {0} mailbox(es) found." -f $allMbx.Count) -ForegroundColor Green

$rows = New-Object System.Collections.Generic.List[object]
$idx = 0
foreach ($m in $allMbx) {
    $idx++
    Write-Progress -Activity 'Collecting mailbox stats' -Status "$idx / $($allMbx.Count): $($m.UserPrincipalName)" -PercentComplete (($idx / [Math]::Max($allMbx.Count,1)) * 100)

    $primaryGB = 0.0
    $archiveGB = 0.0

    try {
        $st = Get-EXOMailboxStatistics -Identity $m.UserPrincipalName -ErrorAction Stop
        $primaryGB = [math]::Round((Get-SizeBytes $st.TotalItemSize) / 1GB, 2)
    } catch { }

    $archiveEnabled = $false
    if ($m.ArchiveStatus -eq 'Active' -or ($m.ArchiveDatabase -and "$($m.ArchiveDatabase)" -ne '')) {
        $archiveEnabled = $true
    }
    if ($archiveEnabled) {
        try {
            $aSt = Get-EXOMailboxStatistics -Identity $m.UserPrincipalName -Archive -ErrorAction Stop
            if ($aSt) { $archiveGB = [math]::Round((Get-SizeBytes $aSt.TotalItemSize) / 1GB, 2) }
        } catch { }
    }
    $totalGB = [math]::Round($primaryGB + $archiveGB, 2)

    # Quota: ProhibitSendReceiveQuota. Treat 'Unlimited' as 0 (no enforced ceiling).
    $quotaGB = 0.0
    $quotaRaw = $m.ProhibitSendReceiveQuota
    if ($quotaRaw -and "$quotaRaw" -ne 'Unlimited') {
        $quotaGB = [math]::Round((Get-SizeBytes $quotaRaw) / 1GB, 2)
    }
    if ($quotaGB -gt 0) {
        $usedPct     = [math]::Round(($primaryGB / $quotaGB) * 100, 2)
        $availableGB = [math]::Round(($quotaGB - $primaryGB), 2)
    } else {
        $usedPct     = 0.00
        $availableGB = 0.00
    }
    $overQuota = ($quotaGB -gt 0 -and $usedPct -ge 80)

    # Licence SKUs via Graph
    $skuName = ''
    if ($m.ExternalDirectoryObjectId) {
        try {
            $u = Get-MgUser -UserId $m.ExternalDirectoryObjectId -Property AssignedLicenses -ErrorAction Stop
            $names = @()
            foreach ($lic in $u.AssignedLicenses) {
                $key = [string]$lic.SkuId
                if ($skuMap.ContainsKey($key)) { $names += $skuMap[$key] } else { $names += $key }
            }
            $skuName = ($names -join ', ')
        } catch { }
    }

    $isShared = ("$($m.RecipientTypeDetails)" -eq 'SharedMailbox')
    $rows.Add([PSCustomObject]@{
        DisplayName         = $m.DisplayName
        Email               = $m.UserPrincipalName
        Type                = "$($m.RecipientTypeDetails)"
        'Used (GB)'         = $primaryGB
        'Quota (GB)'        = $quotaGB
        'Available (GB)'    = $availableGB
        'Used %'            = $usedPct
        'Archive Enabled'   = if ($archiveEnabled) { 'Yes' } else { 'No' }
        'Archive Used (GB)' = $archiveGB
        'Total Used (GB)'   = $totalGB
        'License SKU'       = $skuName
        SharedWithLicense   = ($isShared -and $skuName -ne '')
        OverQuota           = $overQuota
    }) | Out-Null
}
Write-Progress -Activity 'Collecting mailbox stats' -Completed

# Sort descending by total size
$rows = @($rows | Sort-Object -Property 'Total Used (GB)' -Descending)

# ----------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------
$totalMbx     = $rows.Count
$totalStorage = if ($rows.Count -gt 0) { [math]::Round((($rows | Measure-Object -Property 'Total Used (GB)' -Sum).Sum), 2) } else { 0 }
$archiveCount = @($rows | Where-Object { $_.'Archive Enabled' -eq 'Yes' }).Count
$largestGB    = if ($rows.Count -gt 0) { $rows[0].'Total Used (GB)' } else { 0 }
$sharedLicCount = @($rows | Where-Object { $_.SharedWithLicense }).Count
$overQuotaCount = @($rows | Where-Object { $_.OverQuota }).Count

Write-Host ''
Write-Host ('[+] Mailboxes        : {0}' -f $totalMbx) -ForegroundColor Cyan
Write-Host ('[+] Total storage GB : {0}' -f $totalStorage) -ForegroundColor Cyan
Write-Host ('[+] Archive enabled  : {0}' -f $archiveCount) -ForegroundColor Cyan
Write-Host ('[+] Largest mailbox  : {0:N2} GB' -f $largestGB) -ForegroundColor Cyan
if ($overQuotaCount -gt 0) {
    Write-Host ('[!] Over 80% quota   : {0}' -f $overQuotaCount) -ForegroundColor Yellow
}
if ($sharedLicCount -gt 0) {
    Write-Host ('[!] Shared mailboxes with licences : {0}' -f $sharedLicCount) -ForegroundColor Yellow
}

# ----------------------------------------------------------------------------
# Build HTML
# ----------------------------------------------------------------------------
Write-Host '[+] Building HTML report...' -ForegroundColor Cyan

$rowsHtml = New-Object System.Text.StringBuilder
foreach ($r in $rows) {
    $cls = if ($r.SharedWithLicense -or $r.OverQuota) { 'flag' } else { '' }
    [void]$rowsHtml.AppendLine(('<tr class="{0}">' -f $cls))
    [void]$rowsHtml.AppendLine(('  <td>{0}</td>' -f (ConvertTo-HtmlEntity $r.DisplayName)))
    [void]$rowsHtml.AppendLine(('  <td class="mono">{0}</td>' -f (ConvertTo-HtmlEntity $r.Email)))
    [void]$rowsHtml.AppendLine(('  <td>{0}</td>' -f (ConvertTo-HtmlEntity $r.Type)))
    [void]$rowsHtml.AppendLine(('  <td class="num">{0}</td>' -f ('{0:N2}' -f $r.'Used (GB)')))
    [void]$rowsHtml.AppendLine(('  <td class="num">{0}</td>' -f ('{0:N2}' -f $r.'Quota (GB)')))
    [void]$rowsHtml.AppendLine(('  <td class="num">{0}</td>' -f ('{0:N2}' -f $r.'Available (GB)')))
    [void]$rowsHtml.AppendLine(('  <td class="num">{0}</td>' -f ('{0:N2}' -f $r.'Used %')))
    [void]$rowsHtml.AppendLine(('  <td>{0}</td>' -f (ConvertTo-HtmlEntity $r.'Archive Enabled')))
    [void]$rowsHtml.AppendLine(('  <td class="num">{0}</td>' -f ('{0:N2}' -f $r.'Archive Used (GB)')))
    [void]$rowsHtml.AppendLine(('  <td class="num strong">{0}</td>' -f ('{0:N2}' -f $r.'Total Used (GB)')))
    [void]$rowsHtml.AppendLine(('  <td class="mono">{0}</td>' -f (ConvertTo-HtmlEntity $r.'License SKU')))
    [void]$rowsHtml.AppendLine('</tr>')
}

# Mailbox JSON for the CSV button
$mbxJson = ($rows | Select-Object DisplayName,Email,Type,'Used (GB)','Quota (GB)','Available (GB)','Used %','Archive Enabled','Archive Used (GB)','Total Used (GB)','License SKU' | ConvertTo-Json -Compress -Depth 5)
if (-not $mbxJson) { $mbxJson = '[]' }
elseif ($mbxJson -notmatch '^\[') { $mbxJson = "[$mbxJson]" }
$mbxJson = $mbxJson -replace '</', '<\/'

$template = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Mailbox Report - {{TENANT_NAME}} - {{DATE}}</title>
<meta name="description" content="Exchange Online mailbox storage and archive report for {{TENANT_NAME}}, generated {{DATE}} by AroraMSP." />
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
  --good:#10b981; --warn:#f59e0b; --warn-soft:rgba(245,158,11,.16); --red:#ef4444;
  --sans:'Inter',-apple-system,BlinkMacSystemFont,sans-serif;
  --mono:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,monospace;
  --maxw:1400px; --pad:clamp(20px,4vw,40px);
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:var(--bg);color:var(--fg);font-family:var(--sans);font-size:14px;line-height:1.55;-webkit-font-smoothing:antialiased;display:flex;flex-direction:column;min-height:100vh}
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

.summary{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:18px}
@media (max-width:1000px){.summary{grid-template-columns:repeat(3,1fr)}}
@media (max-width:640px){.summary{grid-template-columns:repeat(2,1fr)}}
.sum-card{background:var(--bg-1);border:1px solid var(--line);border-radius:10px;padding:18px 20px}
.sum-card .num{font-family:var(--mono);font-size:24px;font-weight:600;line-height:1;color:var(--cyan)}
.sum-card .num .unit{font-size:14px;color:var(--fg-mute);margin-left:4px}
.sum-card .lbl{font-family:var(--mono);font-size:11px;color:var(--fg-mute);text-transform:uppercase;letter-spacing:.08em;margin-top:8px}

.actions{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:24px}
.btn{font-family:var(--mono);font-size:13px;font-weight:500;padding:10px 16px;border-radius:8px;cursor:pointer;display:inline-flex;align-items:center;gap:8px;transition:filter .15s,background .2s,border-color .2s;border:1px solid transparent}
.btn.primary{background:var(--accent);color:white;border-color:color-mix(in oklab,white 12%,var(--accent))}
.btn.primary:hover{filter:brightness(1.1)}
.btn.ghost{background:var(--bg-1);color:var(--fg);border-color:var(--line-2)}
.btn.ghost:hover{background:var(--bg-2)}
.btn svg{width:14px;height:14px}

.tbl-wrap{overflow-x:auto;background:var(--bg-1);border:1px solid var(--line);border-radius:10px}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{position:sticky;top:0;background:var(--bg-2);text-align:left;font-family:var(--mono);font-size:11px;color:var(--fg-mute);text-transform:uppercase;letter-spacing:.06em;font-weight:500;padding:12px 14px;border-bottom:1px solid var(--line);cursor:pointer;user-select:none;white-space:nowrap}
thead th:hover{color:var(--fg)}
thead th[data-dir]::after{content:" ▾";color:var(--cyan);font-size:10px}
thead th[data-dir="asc"]::after{content:" ▴"}
tbody td{padding:11px 14px;border-bottom:1px solid var(--line);vertical-align:top}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover{background:var(--bg-2)}
.num{font-family:var(--mono);text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.strong{font-weight:600;color:var(--cyan)}
.mono{font-family:var(--mono);font-size:12px;color:var(--fg-dim)}
tbody tr.flag{background:var(--warn-soft)}
tbody tr.flag:hover{background:color-mix(in oklab,var(--warn-soft) 80%,white)}
tbody tr.flag td{border-bottom-color:rgba(245,158,11,.3)}

footer{border-top:1px solid var(--line);padding:22px 0 20px;font-family:var(--mono);font-size:12px;color:var(--fg-mute);margin-top:auto;position:relative;z-index:1}
.foot-top{display:flex;flex-wrap:wrap;gap:18px;justify-content:space-between;align-items:center;padding-bottom:14px;margin-bottom:14px;border-bottom:1px dashed var(--line)}
.foot-top a:hover{color:var(--fg)}
.foot-bottom{text-align:center;font-size:11px;letter-spacing:.06em;text-transform:uppercase}
.foot-bottom .dot{color:var(--good)}

.legend{font-family:var(--mono);font-size:11px;color:var(--fg-mute);margin:14px 0 0;display:flex;align-items:center;gap:10px}
.legend .dot{display:inline-block;width:10px;height:10px;border-radius:2px;background:var(--warn-soft);border:1px solid rgba(245,158,11,.4)}

@media print{
  @page{margin:12mm 10mm}
  :root{
    --bg:#fff; --bg-1:#fff; --bg-2:#f7f8fa;
    --line:#d4d8e0; --line-2:#b8bdc7;
    --fg:#0f1117; --fg-dim:#1f2430; --fg-mute:#4b5563;
    --accent:#5b50d6; --cyan:#0e7490;
    --warn:#b45309; --warn-soft:#fef3c7; --good:#047857;
  }
  body{background:#fff;color:#0f1117;font-size:9.5pt;line-height:1.35}
  .grid-bg,.top,.actions{display:none !important}
  main,.page{padding:0}
  .wrap{max-width:100%;padding:0}
  .page-head h1{font-size:18pt}
  .summary{gap:8px;margin-bottom:12px}
  .sum-card{border:1px solid #d4d8e0;padding:8px 10px;break-inside:avoid}
  .sum-card .num{font-size:14pt}
  .tbl-wrap{border:1px solid #d4d8e0;overflow:visible}
  thead th{background:#f7f8fa;color:#4b5563;font-size:8pt;padding:8px 10px;position:static}
  thead th[data-dir]::after,thead th::after{display:none}
  tbody td{padding:6px 10px;font-size:9pt}
  tbody tr.flag{background:#fef3c7}
  tbody tr{break-inside:avoid}
  footer{border-top:1px solid #d4d8e0;padding:8px 0}
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
      <div class="sec-label">// exchange online · mailbox report · v1.0</div>
      <h1>Mailbox storage report.</h1>
      <p>Per-mailbox primary and archive size, last logon, and assigned licence SKU for the {{TENANT_NAME}} tenant. Sorted descending by total size. Click a column to re-sort.</p>
    </div>

    <div class="summary">
      <div class="sum-card"><div class="num">{{TOTAL_MBX}}</div><div class="lbl">Total mailboxes</div></div>
      <div class="sum-card"><div class="num">{{TOTAL_GB}}<span class="unit">GB</span></div><div class="lbl">Total storage</div></div>
      <div class="sum-card"><div class="num">{{ARCHIVE_COUNT}}</div><div class="lbl">Archive enabled</div></div>
      <div class="sum-card"><div class="num">{{LARGEST_GB}}<span class="unit">GB</span></div><div class="lbl">Largest mailbox</div></div>
      <div class="sum-card"><div class="num">{{OVER_QUOTA_COUNT}}</div><div class="lbl">Over 80% quota</div></div>
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

    <div class="tbl-wrap">
      <table id="mbxTable">
        <thead>
          <tr>
            <th data-sort="text">DisplayName</th>
            <th data-sort="text">Email</th>
            <th data-sort="text">Type</th>
            <th data-sort="num">Used (GB)</th>
            <th data-sort="num">Quota (GB)</th>
            <th data-sort="num">Available (GB)</th>
            <th data-sort="num">Used %</th>
            <th data-sort="text">Archive Enabled</th>
            <th data-sort="num">Archive Used (GB)</th>
            <th data-sort="num" data-dir="desc">Total Used (GB)</th>
            <th data-sort="text">License SKU</th>
          </tr>
        </thead>
        <tbody>
{{ROWS_HTML}}
        </tbody>
      </table>
    </div>

    <p class="legend"><span class="dot"></span> shared mailbox with assigned licence, or mailbox over 80% of quota (review)</p>
  </section>
</main>

<footer>
  <div class="wrap">
    <div class="foot-top">
      <div>© {{YEAR}} AroraMSP · Microsoft 365 consulting</div>
      <a href="https://aroramsp.com" target="_blank" rel="noopener">aroramsp.com →</a>
    </div>
    <div class="foot-bottom"><span class="dot">●</span> Mailbox report · v1.0</div>
  </div>
</footer>

<script>
const MAILBOXES = {{MBX_JSON}};
const TENANT = "{{TENANT_NAME_JS}}";
const DATE = "{{DATE}}";

function sortTable(idx, type) {
  const table = document.getElementById('mbxTable');
  const tbody = table.tBodies[0];
  const ths = table.tHead.rows[0].cells;
  const th = ths[idx];
  const dir = th.dataset.dir === 'asc' ? 'desc' : 'asc';
  for (const t of ths) t.removeAttribute('data-dir');
  th.dataset.dir = dir;
  const rows = Array.from(tbody.rows);
  rows.sort((a, b) => {
    let va = a.cells[idx].textContent.trim();
    let vb = b.cells[idx].textContent.trim();
    if (type === 'num') {
      va = parseFloat(va.replace(/,/g, '')) || 0;
      vb = parseFloat(vb.replace(/,/g, '')) || 0;
    } else if (type === 'date') {
      va = va ? Date.parse(va) || 0 : 0;
      vb = vb ? Date.parse(vb) || 0 : 0;
    }
    return (va > vb ? 1 : va < vb ? -1 : 0) * (dir === 'asc' ? 1 : -1);
  });
  rows.forEach(r => tbody.appendChild(r));
}

document.querySelectorAll('#mbxTable thead th').forEach((th, idx) => {
  th.addEventListener('click', () => sortTable(idx, th.dataset.sort || 'text'));
});

function csvEscape(v) { return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"'; }
function downloadCsv() {
  const headers = ['DisplayName','Email','Type','Used (GB)','Quota (GB)','Available (GB)','Used %','Archive Enabled','Archive Used (GB)','Total Used (GB)','License SKU'];
  const lines = [headers.join(',')];
  for (const m of MAILBOXES) lines.push(headers.map(h => csvEscape(m[h])).join(','));
  const blob = new Blob(['﻿' + lines.join('\r\n')], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const safe = (TENANT + '-' + DATE).replace(/[^A-Za-z0-9._-]/g, '_');
  a.download = 'Mailbox-Report-' + safe + '.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}
document.getElementById('btnPrint').addEventListener('click', () => window.print());
document.getElementById('btnCsv').addEventListener('click', downloadCsv);
</script>
</body>
</html>
'@

$tenantJs = $tenantName -replace '\\', '\\\\' -replace '"', '\"'
$largestStr = if ($largestGB -gt 0) { '{0:N2}' -f $largestGB } else { '0.00' }
$totalStr   = '{0:N2}' -f $totalStorage

$html = $template
$html = $html.Replace('{{TENANT_NAME}}',     (ConvertTo-HtmlEntity $tenantName))
$html = $html.Replace('{{TENANT_NAME_JS}}',  $tenantJs)
$html = $html.Replace('{{DATE}}',            $today)
$html = $html.Replace('{{YEAR}}',            "$runYear")
$html = $html.Replace('{{TOTAL_MBX}}',       "$totalMbx")
$html = $html.Replace('{{TOTAL_GB}}',        $totalStr)
$html = $html.Replace('{{ARCHIVE_COUNT}}',   "$archiveCount")
$html = $html.Replace('{{LARGEST_GB}}',       $largestStr)
$html = $html.Replace('{{OVER_QUOTA_COUNT}}', "$overQuotaCount")
$html = $html.Replace('{{ROWS_HTML}}',        $rowsHtml.ToString())
$html = $html.Replace('{{MBX_JSON}}',         $mbxJson)

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($reportPath, $html, $utf8NoBom)

Write-Host ''
Write-Host '[+] Report complete.' -ForegroundColor Green
Write-Host ("[+] Report saved to: {0}" -f $reportPath) -ForegroundColor Cyan
