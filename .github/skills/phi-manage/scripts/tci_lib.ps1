##############################################################################
# tci_lib.ps1 - Shared helpers for TCI PHI Tracking automation
#
# Dot-source from any script:
#   . "$PSScriptRoot\tci_lib.ps1"
#
# Provides:
#   * Tci-Get / Tci-Post           - URL fallback (HTTP -> HTTPS)
#   * Tci-ParseTables              - depth-tracked <table> extractor
#   * Tci-PickDataTable            - filter by first <th>
#   * Tci-RowsFromTable            - HTML rows -> string[][]
#   * Tci-WriteCsv                 - safe comma-quoted CSV
#   * Tci-ExportXlsx               - open CSV in Excel COM, format, SaveAs
#   * Tci-SendMail                 - Outlook COM email-to-self with attachment
#   * Tci-AssertXlsx               - sanity check before sending
#   * WwToFiscalMonth              - Intel 4-4-5 fiscal month relabel
#   * Test-WwToFiscalMonth         - self-test (run once per year on new data)
#   * AddQ                         - quarter math for MOR relabel
##############################################################################

# ============================================================================
# URL fallback (port-80 probe -> HTTP primary, HTTPS fallback)
# ============================================================================
$script:TciBaseHttp  = 'http://tcitools.intel.com/TestReport.aspx'
$script:TciBaseHttps = 'https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx'
$script:TciBase      = $null
$script:TciBaseCache = Join-Path $env:TEMP 'tci_base_url.cache'

function Get-TciBase {
    # 1. In-process cache
    if ($script:TciBase) { return $script:TciBase }
    # 2. Cross-process file cache (avoids the slow Test-NetConnection probe on
    #    every fresh powershell -Command {} spawn). Trust for 12 hours.
    if (Test-Path $script:TciBaseCache) {
        try {
            $f = Get-Item $script:TciBaseCache
            if ($f.LastWriteTime -gt (Get-Date).AddHours(-12)) {
                $cached = (Get-Content $script:TciBaseCache -Raw).Trim()
                if ($cached) { $script:TciBase = $cached; return $script:TciBase }
            }
        } catch {}
    }
    # 3. Resolve via probe, then persist
    try {
        $t = Test-NetConnection tcitools.intel.com -Port 80 -InformationLevel Quiet -WarningAction SilentlyContinue
        if ($t) { $script:TciBase = $script:TciBaseHttp }
        else    { $script:TciBase = $script:TciBaseHttps }
    } catch { $script:TciBase = $script:TciBaseHttps }
    try { [IO.File]::WriteAllText($script:TciBaseCache, $script:TciBase) } catch {}
    return $script:TciBase
}

function Tci-Get {
    param([Parameter(Mandatory)][string]$ReportId, [string]$OutFile, [int]$TimeoutSec = 180, [int]$Retries = 1)
    $url = "$(Get-TciBase)?R=$ReportId"
    $resp = Invoke-TciRequest -Uri $url -TimeoutSec $TimeoutSec -Retries $Retries
    if ($OutFile) { [IO.File]::WriteAllText($OutFile, $resp.Content, [Text.UTF8Encoding]::new($false)) }
    return $resp
}

function Tci-Post {
    param(
        [Parameter(Mandatory)][string]$ReportId,
        [Parameter(Mandatory)][hashtable]$Form,
        [string]$OutFile,
        [int]$TimeoutSec = 180,
        [int]$Retries = 1
    )
    $url = "$(Get-TciBase)?R=$ReportId"
    $resp = Invoke-TciRequest -Uri $url -Method Post -Body $Form -TimeoutSec $TimeoutSec -Retries $Retries
    if ($OutFile) { [IO.File]::WriteAllText($OutFile, $resp.Content, [Text.UTF8Encoding]::new($false)) }
    return $resp
}

# Wrapper around Invoke-WebRequest that enforces a -TimeoutSec (the default is
# infinite, so a stalled TCI render hangs forever - see lesson 15). On a timeout
# the request is retried up to $Retries times before throwing; the retry usually
# succeeds because the original stall is server-side and transient.
function Invoke-TciRequest {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'Get',
        [hashtable]$Body,
        [int]$TimeoutSec = 180,
        [int]$Retries = 1
    )
    $attempt = 0
    while ($true) {
        try {
            $p = @{ Uri = $Uri; Method = $Method; UseDefaultCredentials = $true; UseBasicParsing = $true; TimeoutSec = $TimeoutSec; ErrorAction = 'Stop' }
            if ($Method -eq 'Post' -and $Body) { $p.Body = $Body }
            return Invoke-WebRequest @p
        } catch {
            $isTimeout = $_.Exception -is [System.Net.WebException] -and $_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout
            if ($attempt -lt $Retries -and $isTimeout) {
                $attempt++
                Write-Host "[http] timeout after ${TimeoutSec}s - retry $attempt/$Retries ..." -ForegroundColor Yellow
                continue
            }
            throw
        }
    }
}

# ----------------------------------------------------------------------------
# Cached init-page fetch. The report init page (filter panels + VIEWSTATE) is
# identical across runs of the same ReportId and a cached VIEWSTATE still
# validates on a later POST (verified 2026-05-29). Caches the HTML per report
# for $MaxAgeMinutes to skip the ~7.6s GET on repeat runs of the same report.
# Returns the init HTML as a string (not the response object).
# ----------------------------------------------------------------------------
function Tci-GetInit {
    param(
        [Parameter(Mandatory)][string]$ReportId,
        [int]$MaxAgeMinutes = 30,
        [switch]$Force
    )
    $cache = Join-Path $env:TEMP "tci_init_$ReportId.html"
    if (-not $Force -and (Test-Path $cache)) {
        $f = Get-Item $cache
        if ($f.LastWriteTime -gt (Get-Date).AddMinutes(-$MaxAgeMinutes)) {
            Write-Host "[init] cache hit: $ReportId (age $([int]((Get-Date)-$f.LastWriteTime).TotalMinutes)m)"
            return (Get-Content $cache -Raw)
        }
    }
    $resp = Tci-Get -ReportId $ReportId
    [IO.File]::WriteAllText($cache, $resp.Content, [Text.UTF8Encoding]::new($false))
    Write-Host "[init] fetched fresh: $ReportId ($($resp.Content.Length) chars)"
    return $resp.Content
}

# ============================================================================
# HTML table parsing (depth-tracked - handles nested tables in ASP.NET output)
# ============================================================================
function Tci-ParseTables {
    param([Parameter(Mandatory)][string]$Html)
    $tables = New-Object System.Collections.Generic.List[string]
    $idx = 0
    while ($idx -lt $Html.Length) {
        $om = [regex]::Match($Html.Substring($idx), '<table\b[^>]*>')
        if (-not $om.Success) { break }
        $sI = $idx + $om.Index + $om.Length
        $depth = 1; $pos = $sI
        while ($depth -gt 0 -and $pos -lt $Html.Length) {
            $o  = $Html.IndexOf('<table', $pos)
            $cl = $Html.IndexOf('</table>', $pos)
            if ($cl -lt 0) { $depth = 0; break }
            if ($o -ge 0 -and $o -lt $cl) { $depth++; $pos = $o + 6 }
            else                          { $depth--; $pos = $cl + 8 }
        }
        $tables.Add($Html.Substring($sI, ($pos - 8) - $sI))
        $idx = $sI
    }
    return $tables
}

function Tci-PickDataTable {
    param(
        [Parameter(Mandatory)][System.Collections.Generic.List[string]]$Tables,
        [Parameter(Mandatory)][string]$FirstTh
    )
    return $Tables | Where-Object {
        $n = [regex]::Matches($_, '(?is)<th[^>]*>\s*([^<]{1,60})') | ForEach-Object { $_.Groups[1].Value.Trim() }
        ($n.Count -gt 0) -and ($n[0] -eq $FirstTh)
    }
}

function Tci-RowsFromTable {
    param([Parameter(Mandatory)][string[]]$DataTableBodies)
    $all = New-Object System.Collections.Generic.List[object]
    $hdr = $null
    foreach ($body in $DataTableBodies) {
        $rs = [regex]::Matches($body, '(?is)<tr[^>]*>(.*?)</tr>')
        $tr = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rs) {
            $cm = [regex]::Matches($r.Groups[1].Value, '(?is)<(t[hd])\b[^>]*>(.*?)</\1>')
            if ($cm.Count -eq 0) { continue }
            $cells = foreach ($mm in $cm) {
                $t = [regex]::Replace($mm.Groups[2].Value, '(?is)<[^>]+>', ' ')
                $t = [System.Net.WebUtility]::HtmlDecode($t).Trim() -replace '\s+', ' '
                , $t
            }
            $tr.Add(@($cells))
        }
        if ($tr.Count -eq 0) { continue }
        if (-not $hdr) { $hdr = $tr[0]; $all.Add($hdr) }
        for ($i = 1; $i -lt $tr.Count; $i++) { $all.Add($tr[$i]) }
    }
    return @{ Header = $hdr; Rows = $all }
}

# ============================================================================
# CSV writer (comma-separated, quoted - safe for Excel COM)
# ============================================================================
function Tci-WriteCsv {
    param(
        [Parameter(Mandatory)][System.Collections.IEnumerable]$Rows,
        [Parameter(Mandatory)][string]$Path
    )
    $maxCols = ($Rows | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    $sb = New-Object System.Text.StringBuilder
    foreach ($row in $Rows) {
        $p = @($row) + (, '') * ($maxCols - $row.Count)
        $e = $p | ForEach-Object {
            $v = "$_"
            if ($v -match '[",\r\n]') { '"' + ($v -replace '"', '""') + '"' } else { $v }
        }
        [void]$sb.AppendLine(($e -join ','))
    }
    [IO.File]::WriteAllText($Path, $sb.ToString(), [Text.UTF8Encoding]::new($true))
}

# ============================================================================
# Excel COM: open CSV, format header, freeze, autofit, SaveAs xlsx
# Optional: custom formatter scriptblock invoked with ($ws, $lastRow, $lastCol)
# ============================================================================
function Tci-ExportXlsx {
    param(
        [Parameter(Mandatory)][string]$CsvPath,
        [Parameter(Mandatory)][string]$XlsxPath,
        [Parameter(Mandatory)][string]$SheetName,
        [int]$FreezeColumns = 0,
        [scriptblock]$CustomFormat,
        # Inline validation (replaces a separate Tci-AssertXlsx COM lifecycle).
        # Pass -Validate to enable; throws if the result fails the checks.
        [switch]$Validate,
        [int]$MinRows = 1,
        [int]$MinCols = 2
    )
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($CsvPath)
        $ws = $wb.Worksheets.Item(1); $ws.Name = $SheetName
        $lastCol = $ws.UsedRange.Columns.Count
        $lastRow = $ws.UsedRange.Rows.Count
        if ($Validate) {
            $dataRows = $lastRow - 1
            if ($lastCol -lt $MinCols) { throw "ExportXlsx: only $lastCol column(s) (expected >= $MinCols). CSV delimiter issue?" }
            if ($dataRows -lt $MinRows) { throw "ExportXlsx: only $dataRows data row(s) (expected >= $MinRows). Empty filter result?" }
        }
        $r = $ws.Range($ws.Cells(1, 1), $ws.Cells(1, $lastCol))
        $r.Font.Bold = $true; $r.Interior.Color = 0xD9D9D9
        if ($FreezeColumns -gt 0) { $ws.Application.ActiveWindow.SplitColumn = $FreezeColumns }
        $ws.Application.ActiveWindow.SplitRow = 1
        $ws.Application.ActiveWindow.FreezePanes = $true
        if ($CustomFormat) { & $CustomFormat $ws $lastRow $lastCol }
        $ws.Columns.AutoFit() | Out-Null
        if (Test-Path $XlsxPath) { Remove-Item $XlsxPath -Force }
        $wb.SaveAs($XlsxPath, 51)
        try { $wb.Close($false) } catch {}
        if ($Validate) { Write-Host "[export+assert] OK: $XlsxPath - $($lastRow - 1) rows x $lastCol cols" }
        return [PSCustomObject]@{ Path = $XlsxPath; Rows = ($lastRow - 1); Cols = $lastCol }
    } finally {
        try { $excel.Quit() } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ============================================================================
# Sanity check: open xlsx, verify col count > 1 and >= 1 data row
# Optional header substrings to assert exist
# ============================================================================
function Tci-AssertXlsx {
    param(
        [Parameter(Mandatory)][string]$XlsxPath,
        [int]$MinRows = 1,
        [int]$MinCols = 2,
        [string[]]$ExpectHeaders
    )
    if (-not (Test-Path $XlsxPath)) { throw "AssertXlsx: file not found: $XlsxPath" }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($XlsxPath)
        $ws = $wb.Worksheets.Item(1)
        $cols = $ws.UsedRange.Columns.Count
        $rows = $ws.UsedRange.Rows.Count - 1
        if ($cols -lt $MinCols) { throw "AssertXlsx: only $cols column(s) (expected >= $MinCols). CSV delimiter issue?" }
        if ($rows -lt $MinRows) { throw "AssertXlsx: only $rows data row(s) (expected >= $MinRows). Empty filter result?" }
        if ($ExpectHeaders) {
            $hdr = @()
            for ($c = 1; $c -le $cols; $c++) { $hdr += "$($ws.Cells(1, $c).Text)" }
            foreach ($eh in $ExpectHeaders) {
                if (-not ($hdr -contains $eh)) { throw "AssertXlsx: missing expected header '$eh'. Got: $($hdr -join '|')" }
            }
        }
        try { $wb.Close($false) } catch {}
        Write-Host "[assert] OK: $XlsxPath - $rows rows x $cols cols"
    } finally {
        try { $excel.Quit() } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ============================================================================
# Outlook COM email-to-self with attachment
# ============================================================================
function Tci-SendMail {
    param(
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$HtmlBody,
        [string[]]$Attachments,
        [string]$To
    )
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    if (-not $To) {
        try { $To = $ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress }
        catch { $To = $ns.CurrentUser.Address }
    }
    $m = $ol.CreateItem(0)
    $m.To = $To; $m.Subject = $Subject
    $m.HTMLBody = "<html><head><style>body,table,td,p,span,div{font-family:Calibri,Arial,sans-serif;font-size:11pt;}</style></head><body style=`"margin:0;padding:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;`">$HtmlBody</body></html>"
    foreach ($a in $Attachments) { if (Test-Path $a) { [void]$m.Attachments.Add($a) } }
    $m.Send()
    Write-Host "[mail] sent to $To"
}

# ============================================================================
# Build-PhiCard - SINGLE SOURCE OF TRUTH for the PHI email card (Calibri 11pt)
# Every sender (batch driver + standalone scripts) MUST call this so the card
# never drifts. See SKILL.md section 8 "Email card is SHARED across ALL senders".
#   $Sections = array of @{ SheetTag='SDA Weekly'; RowCount=70 }
# Returns the full HTML body string.
# ============================================================================
function Build-WifCard {
    param(
        [Parameter(Mandatory)][string]$Product,
        [string]$Group = '',
        [string]$SubGroup = '',
        [Parameter(Mandatory)][string]$Report,
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][string]$Metric,
        [Parameter(Mandatory)][string]$WifValue,
        [Parameter(Mandatory)][string]$TimeRange,
        [Parameter(Mandatory)][int]$PorRows,
        [Parameter(Mandatory)][int]$WifRows,
        [string]$Columns = '',
        [string]$GuideUrl = 'https://goto.intel.com/mpeforge'
    )
    $bd = 'border:1px solid #d4d4d4;'
    $grpRow = if ($Group -or $SubGroup) {
        "<tr><td style=`"${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;`">AT Group / SubGroup</td><td style=`"${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#202124;`">$Group / $SubGroup</td></tr>"
    } else { '' }
    $colRow = if ($Columns) {
        "<tr><td style=`"${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;`">Columns Modified</td><td style=`"${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#202124;`">$Columns</td></tr>"
    } else { '' }
    @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#202124;padding:4px;">
  <div style="background:linear-gradient(90deg,#b45309,#d97706);color:#fff;padding:14px 20px;border-radius:8px 8px 0 0;">
    <span style="font-size:16pt;font-weight:700;letter-spacing:.3px;font-family:Calibri,Arial,sans-serif;">WIF Overlay Report</span>
  </div>
  <table style="border:1px solid #d4d4d4;border-collapse:collapse;width:580px;font-family:Calibri,Arial,sans-serif;font-size:11pt;" cellspacing="0" cellpadding="0">
    <tr><td colspan="2" style="${bd}padding:10px 16px;background:#92400e;color:#fff;font-weight:600;letter-spacing:.3px;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Product &amp; Report</td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;width:200px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Product</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;"><span style="font-weight:700;color:#1e40af;">$Product</span></td></tr>
    $grpRow
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Report</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#202124;">$Report</td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Operation</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;"><span style="display:inline-block;padding:3px 12px;border-radius:10px;background:#dbeafe;color:#1e40af;font-weight:600;font-family:Calibri,Arial,sans-serif;font-size:11pt;">$Operation</span></td></tr>
    <tr><td colspan="2" style="${bd}padding:10px 16px;background:#92400e;color:#fff;font-weight:600;letter-spacing:.3px;font-family:Calibri,Arial,sans-serif;font-size:11pt;">WIF Parameters</td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Metric</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;font-weight:700;color:#202124;">$Metric</td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">WIF Value</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;"><span style="display:inline-block;padding:4px 14px;border-radius:10px;background:#fef3c7;color:#92400e;font-size:11pt;font-weight:700;font-family:Calibri,Arial,sans-serif;">$WifValue</span></td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Time Range</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;font-weight:600;color:#202124;">$TimeRange</td></tr>
    $colRow
    <tr><td colspan="2" style="${bd}padding:10px 16px;background:#92400e;color:#fff;font-weight:600;letter-spacing:.3px;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Output Summary</td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">POR Rows</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#202124;">$PorRows</td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">WIF Rows</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;"><span style="font-weight:700;color:#b45309;">$WifRows</span></td></tr>
    <tr><td style="${bd}padding:10px 16px;background:#fffbeb;font-weight:700;color:#92400e;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Total Rows</td><td style="${bd}padding:10px 16px;background:#fffbeb;text-align:right;font-weight:700;color:#92400e;font-size:11pt;font-family:Calibri,Arial,sans-serif;">$($PorRows + $WifRows)</td></tr>
  </table>
  <div style="margin-top:12px;padding:6px 20px;font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#6b7280;">
    <span style="color:#d97706;font-weight:600;">Legend:</span> Cyan rows = WIF &nbsp;|&nbsp; Blue font = modified cells
  </div>
  <div style="margin-top:6px;padding:4px 20px 12px;font-family:Calibri,Arial,sans-serif;font-size:9pt;color:#9ca3af;font-style:italic;">Generated by MPE Forge Skill 'PHI Manage' &mdash; <a href="$GuideUrl" style="color:#b45309;text-decoration:underline;">Guide</a></div>
</div>
"@
}

function Build-PhiCard {
    param(
        [Parameter(Mandatory)][string]$Product,
        [string]$Group = '',
        [string]$SubGroup = '',
        [Parameter(Mandatory)]$Sections,
        [string]$Filter = 'none',
        [string]$ProductNote = '',
        [string]$GuideUrl = 'https://goto.intel.com/mpeforge'
    )
    if ([string]::IsNullOrWhiteSpace($Filter)) { $Filter = 'none' }
    $pill = { param($t) "<span style=`"display:inline-block;padding:2px 9px;margin:1px;border-radius:10px;background:#fef3c7;color:#92400e;font-size:11pt;font-weight:600;font-family:Calibri,Arial,sans-serif;`">$t</span>" }
    $reportPills = ($Sections | ForEach-Object { & $pill $_.SheetTag }) -join ' '
    $filterPill  = "<span style=`"display:inline-block;padding:3px 12px;margin:1px;border-radius:10px;background:#fef3c7;color:#92400e;font-size:11pt;font-weight:700;text-transform:uppercase;letter-spacing:.5px;font-family:Calibri,Arial,sans-serif;`">$($Filter.ToUpper())</span>"
    $grandTotal  = ($Sections | ForEach-Object { [int]$_.RowCount } | Measure-Object -Sum).Sum
    $noteHtml = if ($ProductNote) { " <span style=`"color:#6b7280;`">$ProductNote</span>" } else { '' }
    $bd = 'border:1px solid #d4d4d4;'   # visible grey border on all cells
    $grpRow = if ($Group -or $SubGroup) {
        "<tr><td style=`"${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;`">AT Group / SubGroup</td><td style=`"${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#202124;`">$Group / $SubGroup</td></tr>"
    } else { '' }
    $bdSection = ''; $z = 0
    foreach ($s in $Sections) {
        $bg = if ($z % 2 -eq 1) { 'background:#fffbeb;' } else { 'background:#fff;' }
        $bdSection += "<tr><td style=`"${bd}padding:8px 16px;font-family:Calibri,Arial,sans-serif;font-size:11pt;$bg color:#202124;`">$($s.SheetTag)</td><td style=`"${bd}padding:8px 16px;font-family:Calibri,Arial,sans-serif;font-size:11pt;$bg text-align:right;color:#202124;font-weight:600;`">$($s.RowCount)</td></tr>"
        $z++
    }
    @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#202124;padding:4px;">
  <div style="background:linear-gradient(90deg,#b45309,#d97706);color:#fff;padding:14px 20px;border-radius:8px 8px 0 0;">
    <span style="font-size:16pt;font-weight:700;letter-spacing:.3px;font-family:Calibri,Arial,sans-serif;">PHI Report Summary</span>
  </div>
  <table style="border:1px solid #d4d4d4;border-collapse:collapse;width:580px;font-family:Calibri,Arial,sans-serif;font-size:11pt;" cellspacing="0" cellpadding="0">
    <tr><td colspan="2" style="${bd}padding:10px 16px;background:#92400e;color:#fff;font-weight:600;letter-spacing:.3px;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Summary Details</td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;width:200px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Product</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;"><span style="font-weight:700;color:#1e40af;">$Product</span>$noteHtml</td></tr>
    $grpRow
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Reports</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">$reportPills</td></tr>
    <tr><td style="${bd}padding:8px 16px;font-weight:600;color:#4a4a4a;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Filter</td><td style="${bd}padding:8px 16px;background:#fff;font-family:Calibri,Arial,sans-serif;font-size:11pt;">$filterPill</td></tr>
    <tr><td colspan="2" style="${bd}padding:10px 16px;background:#92400e;color:#fff;font-weight:600;letter-spacing:.3px;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Rows by Report</td></tr>
    $bdSection
    <tr><td style="${bd}padding:10px 16px;background:#fffbeb;font-weight:700;color:#92400e;font-family:Calibri,Arial,sans-serif;font-size:11pt;">Grand Total</td><td style="${bd}padding:10px 16px;background:#fffbeb;text-align:right;font-weight:700;color:#92400e;font-size:11pt;font-family:Calibri,Arial,sans-serif;">$grandTotal</td></tr>
  </table>
  <div style="margin-top:12px;padding:6px 20px;font-family:Calibri,Arial,sans-serif;font-size:9pt;color:#9ca3af;font-style:italic;">Generated by MPE Forge Skill 'PHI Manage' &mdash; <a href="$GuideUrl" style="color:#b45309;text-decoration:underline;">Guide</a></div>
</div>
"@
}

# ============================================================================
# Add-PhiRunHistory - append one row to tci_run_history.csv for tracking /
# regression detection. Returns the previous row count for the same
# Product+Report+Filter (or $null if none), so callers can flag big drops.
# ============================================================================
function Add-PhiRunHistory {
    param(
        [Parameter(Mandatory)][string]$Product,
        [Parameter(Mandatory)][string]$Report,
        [string]$Filter = '',
        [Parameter(Mandatory)][int]$Rows,
        [int]$Cols = 0,
        [string]$File = '',
        [string]$HistoryPath = "$env:USERPROFILE\Downloads\PHI Tracking\tci_run_history.csv"
    )
    $prev = $null
    if (Test-Path $HistoryPath) {
        $prevRow = Import-Csv $HistoryPath |
            Where-Object { $_.Product -eq $Product -and $_.Report -eq $Report -and $_.Filter -eq $Filter } |
            Select-Object -Last 1
        if ($prevRow) { $prev = [int]$prevRow.Rows }
    }
    [pscustomobject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Product   = $Product
        Report    = $Report
        Filter    = $Filter
        Rows      = $Rows
        Cols      = $Cols
        PrevRows  = if ($null -ne $prev) { $prev } else { '' }
        File      = $File
    } | Export-Csv -Path $HistoryPath -NoTypeInformation -Append
    if ($null -ne $prev -and $prev -gt 0) {
        $drop = ($prev - $Rows) / $prev
        if ($drop -gt 0.20) { Write-Host "[history] WARNING: $Product $Report $Filter rows $prev -> $Rows (down $([math]::Round($drop*100))%)" -ForegroundColor Yellow }
    }
    return $prev
}

# ============================================================================
# Intel 4-4-5 fiscal month relabel (for SDA Monthly headers)
# Anchors: Jan=01, Feb=05, Mar=09, Apr=14, May=18, Jun=22, Jul=27, Aug=31,
#          Sep=35, Oct=40, Nov=44, Dec=48
# 53-week year caveat: standard 52-wk years only. Spot-check on new year.
# ============================================================================
$script:FiscalAnchors = @(1, 5, 9, 14, 18, 22, 27, 31, 35, 40, 44, 48)
$script:FiscalMonths  = @('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')

function WwToFiscalMonth {
    param([Parameter(Mandatory)][string]$Yyyyww)
    if ($Yyyyww -notmatch '^\d{6}$') { return $Yyyyww }
    $y = [int]$Yyyyww.Substring(0, 4); $w = [int]$Yyyyww.Substring(4, 2)
    $mi = -1
    for ($i = 0; $i -lt 12; $i++) { if ($w -ge $script:FiscalAnchors[$i]) { $mi = $i } }
    if ($mi -lt 0) { return $Yyyyww }
    '{0} {1}' -f $script:FiscalMonths[$mi], $y
}

function Test-WwToFiscalMonth {
    # Self-test against verified TCI samples (PTL 4Xe, 2026-05-28).
    $cases = @(
        @{ In = '202622'; Out = 'Jun 2026' }
        @{ In = '202627'; Out = 'Jul 2026' }
        @{ In = '202631'; Out = 'Aug 2026' }
        @{ In = '202635'; Out = 'Sep 2026' }
        @{ In = '202640'; Out = 'Oct 2026' }
        @{ In = '202644'; Out = 'Nov 2026' }
        @{ In = '202648'; Out = 'Dec 2026' }
        @{ In = '202701'; Out = 'Jan 2027' }
        @{ In = '202722'; Out = 'Jun 2027' }
        @{ In = '202748'; Out = 'Dec 2027' }
    )
    $fail = 0
    foreach ($c in $cases) {
        $g = WwToFiscalMonth $c.In
        if ($g -ne $c.Out) { Write-Warning "WwToFiscalMonth($($c.In)) = '$g'  expected '$($c.Out)'"; $fail++ }
    }
    if ($fail -gt 0) { throw "WwToFiscalMonth self-test failed: $fail mismatch(es)" }
    Write-Host "[test] WwToFiscalMonth: all $($cases.Count) cases pass"
}

# ============================================================================
# Quarter math (for MOR ATRev+N -> YYYYQQ)
# ============================================================================
function AddQ {
    param([Parameter(Mandatory)][string]$Yyyyqq, [Parameter(Mandatory)][int]$N)
    $y = [int]$Yyyyqq.Substring(0, 4); $q = [int]$Yyyyqq.Substring(4, 2)
    for ($i = 0; $i -lt $N; $i++) { $q++; if ($q -gt 4) { $q = 1; $y++ } }
    '{0}{1:00}' -f $y, $q
}

# ============================================================================
# ASP.NET checkbox / hidden-field helpers (formerly inlined per-script)
# ============================================================================
function Get-PanelCheckbox {
    <#
    .SYNOPSIS Find the POST name for a single checkbox by its visible label.
    .PARAMETER Html  The full init-page HTML string.
    .PARAMETER PanelId  The DOM id of the filter panel table (e.g. 'ContentPlaceHolder1_Filters_AT_Group').
    .PARAMETER Label  The exact visible label text to match.
    #>
    param([string]$Html, [string]$PanelId, [string]$Label)
    $i = $Html.IndexOf('id="' + $PanelId + '"'); if ($i -lt 0) { return $null }
    $j = $Html.IndexOf('</table>', $i); $s = $Html.Substring($i, $j - $i)
    $m = [regex]::Match($s, '<input id="([^"]+)" type="checkbox" name="([^"]+)"[^/]*/><label for="\1">' + [regex]::Escape($Label) + '</label>')
    if ($m.Success) { $m.Groups[2].Value } else { $null }
}
Set-Alias -Name Get-CB -Value Get-PanelCheckbox -Scope Script

function Get-PanelAllCheckboxes {
    <#
    .SYNOPSIS Return all checkboxes (Name + Label) from a filter panel.
    #>
    param([string]$Html, [string]$PanelId)
    $i = $Html.IndexOf('id="' + $PanelId + '"'); if ($i -lt 0) { return @() }
    $j = $Html.IndexOf('</table>', $i); $s = $Html.Substring($i, $j - $i)
    [regex]::Matches($s, '<input id="([^"]+)" type="checkbox" name="([^"]+)"[^/]*/><label for="\1">([^<]+)</label>') |
        ForEach-Object { [PSCustomObject]@{ Name = $_.Groups[2].Value; Label = $_.Groups[3].Value.Trim() } }
}
Set-Alias -Name Get-AllCB -Value Get-PanelAllCheckboxes -Scope Script

function Get-HiddenField {
    <#
    .SYNOPSIS Extract value of a named hidden input field from ASP.NET HTML.
    #>
    param([string]$Html, [string]$FieldName)
    $m = [regex]::Match($Html, '(?is)<input[^>]*name="' + [regex]::Escape($FieldName) + '"[^>]*value="([^"]*)"')
    if ($m.Success) { return $m.Groups[1].Value }
    $m2 = [regex]::Match($Html, '(?is)<input[^>]*value="([^"]*)"[^>]*name="' + [regex]::Escape($FieldName) + '"')
    if ($m2.Success) { return $m2.Groups[1].Value }
    return ''
}
Set-Alias -Name HF -Value Get-HiddenField -Scope Script

function Build-TciPostBody {
    <#
    .SYNOPSIS Build the complete POST body hashtable for a TCI report.
    .PARAMETER InitHtml  Raw init page HTML.
    .PARAMETER Group     AT_Group label (e.g. 'Client').
    .PARAMETER SubGroup  AT_SubGroup label (e.g. 'Mobile').
    .PARAMETER Product   CommonName label (e.g. 'Wildcat Lake').
    .PARAMETER SelectAllPhi  If true, select all PHIParameters checkboxes.
    #>
    param(
        [Parameter(Mandatory)][string]$InitHtml,
        [Parameter(Mandatory)][string]$Group,
        [Parameter(Mandatory)][string]$SubGroup,
        [Parameter(Mandatory)][string]$Product,
        [switch]$SelectAllPhi
    )
    $grpCtl  = Get-PanelCheckbox $InitHtml 'ContentPlaceHolder1_Filters_AT_Group' $Group
    $subCtl  = Get-PanelCheckbox $InitHtml 'ContentPlaceHolder1_Filters_AT_SubGroup' $SubGroup
    $prodCtl = Get-PanelCheckbox $InitHtml 'ContentPlaceHolder1_Filters_CommonName' $Product
    if (-not $prodCtl) { throw "CommonName '$Product' not found in panel" }
    if (-not $grpCtl)  { throw "AT_Group '$Group' not found in panel" }
    if (-not $subCtl)  { throw "AT_SubGroup '$SubGroup' not found in panel" }

    $body = [ordered]@{
        '__EVENTTARGET'        = ''
        '__EVENTARGUMENT'      = ''
        '__VIEWSTATE'          = Get-HiddenField $InitHtml '__VIEWSTATE'
        '__VIEWSTATEGENERATOR' = Get-HiddenField $InitHtml '__VIEWSTATEGENERATOR'
        '__EVENTVALIDATION'    = Get-HiddenField $InitHtml '__EVENTVALIDATION'
        $grpCtl                = 'on'
        $subCtl                = 'on'
        $prodCtl               = 'on'
        'ctl00$ContentPlaceHolder1$btn_RunReport' = 'Run Report'
    }
    if ($SelectAllPhi) {
        $phis = Get-PanelAllCheckboxes $InitHtml 'ContentPlaceHolder1_Filters_PHIParameters'
        foreach ($p in $phis) { $body[$p.Name] = 'on' }
    }
    return $body
}
