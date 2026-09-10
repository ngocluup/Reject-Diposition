$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Raptor Lake S 8C 16A GT0 (Catlow)'
$safeName  = 'RPL_Catlow_Legacy'
$grp       = 'Client'
$sub       = 'Desktop'
$baseFile  = "${safeName}_MOR_Yield"
$xlsx      = Join-Path $env:TEMP "$baseFile.xlsx"

Write-Host "=== $product - MOR Spread + Yield ===" -ForegroundColor Cyan

# ============================================================================
# PART 1: MOR Spread
# ============================================================================
$morCache = Join-Path $env:TEMP "tci_after_${safeName}_morsp.html"
$morMaxAge = 720  # 12 hours

if ((Test-Path $morCache) -and (Get-Item $morCache).LastWriteTime -gt (Get-Date).AddMinutes(-$morMaxAge)) {
    $age = [int]((Get-Date) - (Get-Item $morCache).LastWriteTime).TotalMinutes
    Write-Host "[MOR] cache reuse: $morCache (age ${age}m)"
    $morHtml = Get-Content $morCache -Raw
} else {
    Write-Host "[MOR] fetching MORSpread..."
    $initHtml = Tci-GetInit -ReportId 'MORSpread' -MaxAgeMinutes 60

    # Parse hidden fields
    $vs  = ([regex]::Match($initHtml, 'name="__VIEWSTATE"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value
    $vsg = ([regex]::Match($initHtml, 'name="__VIEWSTATEGENERATOR"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value
    $ev  = ([regex]::Match($initHtml, 'name="__EVENTVALIDATION"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value

    # PHI Parameters panel
    $phiPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_PHIParameters"(.*?)</div>\s*</div>').Value
    $phiCtls  = [regex]::Matches($phiPanel, 'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }

    # Build form
    $form = [ordered]@{
        '__EVENTTARGET'       = ''
        '__EVENTARGUMENT'     = ''
        '__VIEWSTATE'         = $vs
        '__VIEWSTATEGENERATOR'= $vsg
        '__EVENTVALIDATION'   = $ev
        'ctl00$ContentPlaceHolder1$ctl00' = 'on'   # Client
        'ctl00$ContentPlaceHolder1$ctl10' = 'on'   # Desktop
        'ctl00$ContentPlaceHolder1$ctl274' = 'on'  # Raptor Lake S 8C 16A GT0 (Catlow)
        'ctl00$ContentPlaceHolder1$btn_RunReport' = 'Run Report'
    }
    foreach ($p in $phiCtls) { $form[$p] = 'on' }

    Write-Host "[MOR] POST phi=$($phiCtls.Count)..."
    $resp = Tci-Post -ReportId 'MORSpread' -Form $form -TimeoutSec 300
    $morHtml = $resp.Content
    [IO.File]::WriteAllText($morCache, $morHtml, [Text.UTF8Encoding]::new($false))
    Write-Host "[MOR] saved: $morCache ($($morHtml.Length) chars)"
}

# Parse MOR tables
$morTables = Tci-ParseTables $morHtml
$morDt     = Tci-PickDataTable $morTables 'MetricName'
$morParsed = Tci-RowsFromTable $morDt
$morRows   = $morParsed.Rows
Write-Host "[MOR] rows: $($morRows.Count - 1)  cols: $($morRows[0].Count)"

# ============================================================================
# PART 2: Yield from cached SDA
# ============================================================================
$sdaCache = Join-Path $env:TEMP 'tci_after_catlow_sda.html'
if (-not (Test-Path $sdaCache)) {
    throw "SDA cache not found: $sdaCache - run SDA for this product first"
}
$sdaAge = [int]((Get-Date) - (Get-Item $sdaCache).LastWriteTime).TotalMinutes
Write-Host "[YIELD] using SDA cache (age ${sdaAge}m)"
$sdaHtml   = Get-Content $sdaCache -Raw
$sdaTables = Tci-ParseTables $sdaHtml
$sdaDt     = Tci-PickDataTable $sdaTables 'CommonName'
$sdaParsed = Tci-RowsFromTable $sdaDt
$sdaAll    = $sdaParsed.Rows

$opIdx  = [array]::IndexOf($sdaAll[0], 'OperationName')
$metIdx = [array]::IndexOf($sdaAll[0], 'MetricName')

# Filter: blank OperationName + MetricName in (U/D, R/D, FINISH YIELD)
$yieldRows = New-Object System.Collections.Generic.List[object]
$yieldRows.Add($sdaAll[0])  # header
for ($i = 1; $i -lt $sdaAll.Count; $i++) {
    $op  = "$($sdaAll[$i][$opIdx])".Trim()
    $met = "$($sdaAll[$i][$metIdx])".Trim()
    if ($op -eq '' -and $met -match '(?i)^(U/D|R/D|FINISH YIELD)$') {
        $yieldRows.Add($sdaAll[$i])
    }
}
Write-Host "[YIELD] rows: $($yieldRows.Count - 1) (U/D + R/D + FINISH YIELD)"

if ($yieldRows.Count -le 1) {
    Write-Host "[YIELD] WARNING: no yield rows found!" -ForegroundColor Yellow
}

# ============================================================================
# PART 3: Build 2-tab xlsx
# ============================================================================
if (Test-Path $xlsx) { Remove-Item $xlsx -Force }

$tabs = @(
    @{ Name = "MOR Spread"; Rows = $morRows }
    @{ Name = "Yield";      Rows = $yieldRows }
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Add()
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
    $wb.Worksheets.Item(1).Name = '_placeholder_'

    foreach ($tab in $tabs) {
        $rows = $tab.Rows
        $maxCols = ($rows | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
        $csv = Join-Path $env:TEMP "${baseFile}_$($tab.Name -replace ' ','_').csv"
        $sb = New-Object System.Text.StringBuilder
        foreach ($row in $rows) {
            $p = @($row) + (, '') * ($maxCols - $row.Count)
            $e = $p | ForEach-Object { $v = "$_"; if ($v -match '[",\r\n]') { '"' + ($v -replace '"', '""') + '"' } else { $v } }
            [void]$sb.AppendLine(($e -join ','))
        }
        [IO.File]::WriteAllText($csv, $sb.ToString(), [Text.UTF8Encoding]::new($true))

        $tmp = $excel.Workbooks.Open($csv)
        $src = $tmp.Worksheets.Item(1)
        $src.Move([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Sheets.Count))
        try { $tmp.Close($false) } catch {}

        $ws = $wb.Worksheets.Item($wb.Sheets.Count)
        $ws.Name = "$($tab.Name) ($($rows.Count - 1)r)"
        $ws.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true
        $ws.Range("1:1").Font.Bold = $true
        $ws.Range("1:1").Interior.Color = 0xD9D9D9
        [void]$ws.Columns.AutoFit()
        Remove-Item $csv -Force
    }

    $wb.Worksheets.Item('_placeholder_').Delete()
    $wb.Worksheets.Item(1).Activate()
    $wb.SaveAs($xlsx, 51)
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Host "xlsx: $xlsx ($([Math]::Round((Get-Item $xlsx).Length/1KB, 1)) KB)"

# ============================================================================
# PART 4: Email with Build-PhiCard
# ============================================================================
$sections = @(
    @{ SheetTag = 'MOR Spread'; RowCount = $morRows.Count - 1 }
    @{ SheetTag = 'Yield';      RowCount = $yieldRows.Count - 1 }
)
$body = Build-PhiCard -Product $product -Group $grp -SubGroup $sub -Sections $sections -Filter 'none'
$subj = "PHI of $product - MOR Spread + Yield"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsx)
Write-Host "`nDone."
