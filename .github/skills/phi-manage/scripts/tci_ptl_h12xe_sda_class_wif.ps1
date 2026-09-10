$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Panther Lake H 4P+8E+4LP_E+12Xe'
$safeName  = 'Panther_Lake_H_4P8E4LPE12Xe'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'PORSDAForecast'
$initTag   = 'sda'
$firstTh   = 'CommonName'
$filter    = 'Class'

# WIF config
$wifSpec   = 'Classhot ETT'
$wifStart  = '202640'
$wifEnd    = '202652'
$wifValue  = '9'

Write-Host "=== $product - SDA Weekly + Class + WIF ===" -ForegroundColor Cyan

# --- Fetch / cache ---
$afterFile = Join-Path $env:TEMP "tci_after_${safeName}_${initTag}.html"
$maxAge    = 720

if ((Test-Path $afterFile) -and (Get-Item $afterFile).Length -gt 0 -and (Get-Item $afterFile).LastWriteTime -gt (Get-Date).AddMinutes(-$maxAge)) {
    $age = [int]((Get-Date) - (Get-Item $afterFile).LastWriteTime).TotalMinutes
    Write-Host "[cache] reusing $afterFile (age ${age}m)"
    $html = Get-Content $afterFile -Raw
} else {
    Write-Host "[fetch] GET init..."
    $initHtml = Tci-GetInit -ReportId $reportId -MaxAgeMinutes 60

    $vs  = ([regex]::Match($initHtml, 'name="__VIEWSTATE"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value
    $vsg = ([regex]::Match($initHtml, 'name="__VIEWSTATEGENERATOR"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value
    $ev  = ([regex]::Match($initHtml, 'name="__EVENTVALIDATION"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value

    $phiPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_PHIParameters"(.*?)</div>\s*</div>').Value
    $phiCtls  = [regex]::Matches($phiPanel, 'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }

    $grpPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_Group"(.*?)</div>\s*</div>').Value
    $grpCtl   = ([regex]::Matches($grpPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -eq $grp }).Groups[1].Value

    $subPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_SubGroup"(.*?)</div>\s*</div>').Value
    $subCtl   = ([regex]::Matches($subPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -eq $sub }).Groups[1].Value

    $cnPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_CommonName"(.*?)</div>\s*</div>').Value
    $escaped = [regex]::Escape($product)
    $cnCtl   = ([regex]::Matches($cnPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -match "^${escaped}" }).Groups[1].Value
    if (-not $cnCtl) { Write-Error "CommonName ctl not found for '$product'"; exit 1 }
    Write-Host "[ctl] CommonName = $cnCtl"

    $form = [ordered]@{
        '__EVENTTARGET'=''; '__EVENTARGUMENT'=''
        '__VIEWSTATE'=$vs; '__VIEWSTATEGENERATOR'=$vsg; '__EVENTVALIDATION'=$ev
        $grpCtl='on'; $subCtl='on'; $cnCtl='on'
        'ctl00$ContentPlaceHolder1$btn_RunReport'='Run Report'
    }
    foreach ($p in $phiCtls) { $form[$p] = 'on' }

    Write-Host "[fetch] POST phi=$($phiCtls.Count)..."
    $resp = Tci-Post -ReportId $reportId -Form $form -TimeoutSec 600 -Retries 1
    $html = $resp.Content
    if (-not $html -or $html.Length -lt 1000) {
        Set-Content $afterFile $resp
        $html = Get-Content $afterFile -Raw
    }
    if ($html.Length -lt 1000) { Write-Error "POST returned empty/short response"; exit 1 }
    [IO.File]::WriteAllText($afterFile, $html, [Text.UTF8Encoding]::new($false))
    Write-Host "[fetch] saved: $afterFile ($($html.Length) chars)"
}

# --- Parse ---
$tables = Tci-ParseTables $html
$dt     = Tci-PickDataTable $tables $firstTh
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr     = $allRows[0]
Write-Host "total rows: $($allRows.Count - 1) cols: $($hdr.Count)"

$opIdx   = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx  = [array]::IndexOf([string[]]$hdr, 'MetricName')
$subOIdx = [array]::IndexOf([string[]]$hdr, 'SubObject')
$bomIdx  = [array]::IndexOf([string[]]$hdr, 'BOM')
$porIdx  = [array]::IndexOf([string[]]$hdr, 'Proposed/POR')
Write-Host "Indexes: Op=$opIdx Met=$metIdx SubObj=$subOIdx BOM=$bomIdx POR=$porIdx"

# --- Class filter (Tab 1 baseline) ---
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
$monitors = @('MPS MONITOR','EQA MONITOR','CS MONITOR')
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim().ToUpper()
    $byClass   = $op -match '^(?i)TEST'
    $byMonitor = [string]::IsNullOrWhiteSpace($op) -and ($monitors -contains $met)
    if ($byClass -or $byMonitor) { $filtered.Add($allRows[$i]) }
}
Write-Host "Class filter: $($filtered.Count - 1) rows"
if ($filtered.Count -le 1) { Write-Error "No rows after Class filter"; exit 1 }

# --- WIF: find matching rows (Classhot ETT) ---
# OperationName=TEST_MPS AND SubObject=PBIC1 AND MetricName=TEST_TIME-SEC AND BOM=blank
$wifMatches = New-Object System.Collections.Generic.List[int]
for ($i = 1; $i -lt $filtered.Count; $i++) {
    $row = $filtered[$i]
    $op  = "$($row[$opIdx])".Trim()
    $met = "$($row[$metIdx])".Trim()
    $so  = "$($row[$subOIdx])".Trim()
    $bom = "$($row[$bomIdx])".Trim()
    if ($op -eq 'TEST_MPS' -and $so -eq 'PBIC1' -and $met -eq 'TEST_TIME-SEC' -and $bom -eq '') {
        $wifMatches.Add($i)
    }
}
Write-Host "WIF matches (Classhot ETT): $($wifMatches.Count) rows"
if ($wifMatches.Count -eq 0) { Write-Error "No rows match Classhot ETT spec"; exit 1 }

# --- Resolve WIF target columns ---
$startCol = -1; $endCol = -1
for ($c = 0; $c -lt $hdr.Count; $c++) {
    if ($hdr[$c] -eq $wifStart -and $startCol -lt 0) { $startCol = $c }
    if ($hdr[$c] -eq $wifEnd) { $endCol = $c }
}
if ($startCol -lt 0) { Write-Error "Start col $wifStart not found in headers"; exit 1 }
if ($endCol -lt 0) {
    # EOL fallback
    for ($c = $hdr.Count - 1; $c -ge 0; $c--) {
        if ($hdr[$c] -match '^\d{6}$') { $endCol = $c; break }
    }
}
Write-Host "WIF columns: $startCol ($($hdr[$startCol])) to $endCol ($($hdr[$endCol]))"

# --- Build Tab 2 (WIF tab): POR rows + WIF clones ---
$wifRows = New-Object System.Collections.Generic.List[object]
$wifRows.Add($hdr)
$rowRedCols = @{}  # key=excel row -> value=col indexes with red font

foreach ($mi in $wifMatches) {
    # Add POR row as-is
    $wifRows.Add($filtered[$mi])
    
    # Clone for WIF
    $clone = @($filtered[$mi] | ForEach-Object { $_ })
    $clone[$porIdx] = 'WIF'
    for ($c = $startCol; $c -le $endCol; $c++) {
        if ($hdr[$c] -match '^\d{6}$') { $clone[$c] = $wifValue }
    }
    $wifRows.Add($clone)
    # Track red cols for this WIF row (excel row = current count, since header is row 1)
    $excelRow = $wifRows.Count
    $redCols = @()
    for ($c = $startCol; $c -le $endCol; $c++) {
        if ($hdr[$c] -match '^\d{6}$') { $redCols += ($c + 1) }  # 1-based for Excel
    }
    $rowRedCols[$excelRow] = $redCols
}
Write-Host "WIF tab: $($wifRows.Count - 1) rows (POR + WIF)"

# --- Write Tab 1 CSV (baseline) ---
$csvBase  = Join-Path $env:TEMP "${safeName}_SDA_Class.csv"
$csvWif   = Join-Path $env:TEMP "${safeName}_SDA_Class_WIF.csv"
$xlsxFile = Join-Path $env:TEMP "${safeName}_SDA_Class_WIF.xlsx"

Tci-WriteCsv -Rows $filtered -Path $csvBase
Tci-WriteCsv -Rows $wifRows -Path $csvWif

# --- Build 2-tab xlsx ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false

try {
    # Target workbook with placeholder
    $targetWb = $excel.Workbooks.Add()
    $targetWb.Worksheets.Item(1).Name = '_placeholder_'

    # Tab 1: Class Baseline
    $wbBase = $excel.Workbooks.Open($csvBase)
    $wsBase = $wbBase.Worksheets.Item(1)
    $wsBase.Name = 'Class Baseline'
    $wsBase.Move([System.Reflection.Missing]::Value, $targetWb.Worksheets.Item($targetWb.Worksheets.Count))
    try { $wbBase.Close($false) } catch {}

    # Tab 2: WIF
    $wbWif = $excel.Workbooks.Open($csvWif)
    $wsWif = $wbWif.Worksheets.Item(1)
    $wsWif.Name = 'WIF'
    $wsWif.Move([System.Reflection.Missing]::Value, $targetWb.Worksheets.Item($targetWb.Worksheets.Count))
    try { $wbWif.Close($false) } catch {}

    # Delete placeholder
    $targetWb.Worksheets.Item('_placeholder_').Delete()

    # Format Tab 1
    $ws1 = $targetWb.Worksheets.Item('Class Baseline')
    $ws1.Activate()
    $lastCol1 = $ws1.UsedRange.Columns.Count
    $hdrRng1 = $ws1.Range($ws1.Cells(1,1), $ws1.Cells(1,$lastCol1))
    $hdrRng1.Font.Bold = $true
    $hdrRng1.Interior.Color = 14277081
    try { $aw = $excel.ActiveWindow; $aw.SplitColumn = 4; $aw.SplitRow = 1; $aw.FreezePanes = $true } catch {}
    $ws1.Columns.AutoFit() | Out-Null

    # Format Tab 2
    $ws2 = $targetWb.Worksheets.Item('WIF')
    $ws2.Activate()
    $lastCol2 = $ws2.UsedRange.Columns.Count
    $hdrRng2 = $ws2.Range($ws2.Cells(1,1), $ws2.Cells(1,$lastCol2))
    $hdrRng2.Font.Bold = $true
    $hdrRng2.Interior.Color = 14277081
    try { $aw2 = $excel.ActiveWindow; $aw2.SplitColumn = 4; $aw2.SplitRow = 1; $aw2.FreezePanes = $true } catch {}

    # Apply WIF formatting: yellow bg on WIF rows, red font on changed cells
    $porColExcel = $porIdx + 1
    $totalRows2 = $wifRows.Count
    for ($r = 2; $r -le $totalRows2; $r++) {
        $porVal = $ws2.Cells($r, $porColExcel).Text
        if ($porVal -eq 'WIF') {
            # Yellow background entire row
            $ws2.Range($ws2.Cells($r,1), $ws2.Cells($r,$lastCol2)).Interior.Color = 0x66FFFF  # yellow
            # Red font on changed columns
            if ($rowRedCols.ContainsKey($r)) {
                foreach ($col in $rowRedCols[$r]) {
                    $ws2.Cells($r, $col).Font.Color = 0x0000FF  # red (BGR)
                }
            }
        }
    }
    $ws2.Columns.AutoFit() | Out-Null

    $targetWb.SaveAs($xlsxFile, 51)
    Write-Host "xlsx: $xlsxFile ($([Math]::Round((Get-Item $xlsxFile).Length/1KB, 1)) KB)"
} finally {
    try { $targetWb.Close($false) } catch {}
    try { $excel.Quit() } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Remove-Item $csvBase -Force -ErrorAction SilentlyContinue
    Remove-Item $csvWif -Force -ErrorAction SilentlyContinue
}

# --- Print POR vs WIF comparison ---
Write-Host "`n--- POR vs WIF ---"
$siteIdx = [array]::IndexOf([string[]]$hdr, 'Site')
$dlcpIdx = [array]::IndexOf([string[]]$hdr, 'DLCP')
foreach ($mi in $wifMatches) {
    $row = $filtered[$mi]
    $site = if ($siteIdx -ge 0) { "$($row[$siteIdx])".Trim() } else { '?' }
    $dlcp = if ($dlcpIdx -ge 0) { "$($row[$dlcpIdx])".Trim() } else { '' }
    $porVals = @()
    for ($c = $startCol; $c -le [Math]::Min($startCol+2, $endCol); $c++) {
        if ($hdr[$c] -match '^\d{6}$') { $porVals += "$($hdr[$c])=$($row[$c])" }
    }
    Write-Host "  Site=$site DLCP=$dlcp POR: $($porVals -join ', ') -> WIF: $wifValue"
}

# --- Email ---
$sections = @( @{ SheetTag = 'Class Baseline'; RowCount = $filtered.Count - 1 } )
$body = Build-WifCard -Product $product -Group $grp -SubGroup $sub -Report 'SDA Weekly' -Operation 'Class' -Metric $wifSpec -WifValue $wifValue -TimeRange "$wifStart - $wifEnd" -PorRows $wifMatches.Count -WifRows $wifMatches.Count
$subj = "PHI of $product - SDA Weekly Class + WIF $wifSpec"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxFile)
Write-Host "`nDone."
