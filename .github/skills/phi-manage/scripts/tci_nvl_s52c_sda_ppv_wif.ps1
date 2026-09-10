$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Nova Lake S 52C BLLC'
$safeName  = 'Nova_Lake_S_52C_BLLC'
$grp       = 'Client'
$sub       = 'Desktop'
$reportId  = 'PORSDAForecast'
$initTag   = 'sda'
$firstTh   = 'CommonName'
$filter    = 'PPV'

# WIF config
$wifSpec   = 'PPV SS'
$wifStart  = '202640'
$wifEnd    = '202651'
$wifValue  = '10'
$wifSite   = 'SS'   # VNAT = SS

Write-Host "=== $product - SDA Weekly + PPV + WIF (SS $wifSite) ===" -ForegroundColor Cyan

# --- Use cached data ---
$afterFile = Join-Path $env:TEMP "tci_after_${safeName}_${initTag}.html"
if (-not (Test-Path $afterFile) -or (Get-Item $afterFile).Length -lt 1000) {
    Write-Error "Cache not found: $afterFile"; exit 1
}
$age = [int]((Get-Date) - (Get-Item $afterFile).LastWriteTime).TotalMinutes
Write-Host "[cache] reusing $afterFile (age ${age}m)"
$html = Get-Content $afterFile -Raw

# --- Parse ---
$tables = Tci-ParseTables $html
$dt     = Tci-PickDataTable $tables $firstTh
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr     = $allRows[0]
Write-Host "total rows: $($allRows.Count - 1) cols: $($hdr.Count)"

$opIdx   = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx  = [array]::IndexOf([string[]]$hdr, 'MetricName')
$siteIdx = [array]::IndexOf([string[]]$hdr, 'Site')
$porIdx  = [array]::IndexOf([string[]]$hdr, 'Proposed/POR')
Write-Host "Indexes: Op=$opIdx Met=$metIdx Site=$siteIdx POR=$porIdx"

# --- PPV filter (Tab 1 baseline) ---
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim()
    $isPPV = ($op -ne '' -and $op -like 'PPV*') -or
             ($op -eq '' -and $met -match '(?i)^PPV-M SAMPLE SIZE$')
    if ($isPPV) { $filtered.Add($allRows[$i]) }
}
Write-Host "PPV filter: $($filtered.Count - 1) rows"

# --- WIF: find matching rows (PPV-M SAMPLE SIZE, Site=SS) ---
$wifMatches = New-Object System.Collections.Generic.List[int]
for ($i = 1; $i -lt $filtered.Count; $i++) {
    $row  = $filtered[$i]
    $op   = "$($row[$opIdx])".Trim()
    $met  = "$($row[$metIdx])".Trim()
    $site = "$($row[$siteIdx])".Trim()
    if ($op -eq '' -and $met -eq 'PPV-M SAMPLE SIZE' -and $site -eq $wifSite) {
        $wifMatches.Add($i)
    }
}
Write-Host "WIF matches (PPV SS, Site=$wifSite): $($wifMatches.Count) rows"

# Fallback: if site not found, try blank site
if ($wifMatches.Count -eq 0) {
    Write-Host "[warn] Site=$wifSite not found, trying blank site fallback..."
    for ($i = 1; $i -lt $filtered.Count; $i++) {
        $row  = $filtered[$i]
        $op   = "$($row[$opIdx])".Trim()
        $met  = "$($row[$metIdx])".Trim()
        $site = "$($row[$siteIdx])".Trim()
        if ($op -eq '' -and $met -eq 'PPV-M SAMPLE SIZE' -and $site -eq '') {
            $wifMatches.Add($i)
        }
    }
    if ($wifMatches.Count -gt 0) { Write-Host "[warn] Using blank-site row(s) as fallback" }
}
if ($wifMatches.Count -eq 0) { Write-Error "No rows match PPV SS spec"; exit 1 }

# --- Resolve WIF target columns ---
$startCol = -1; $endCol = -1
for ($c = 0; $c -lt $hdr.Count; $c++) {
    if ($hdr[$c] -eq $wifStart -and $startCol -lt 0) { $startCol = $c }
    if ($hdr[$c] -eq $wifEnd) { $endCol = $c }
}
if ($startCol -lt 0) { Write-Error "Start col $wifStart not found"; exit 1 }
if ($endCol -lt 0) {
    for ($c = $hdr.Count - 1; $c -ge 0; $c--) {
        if ($hdr[$c] -match '^\d{6}$') { $endCol = $c; break }
    }
}
Write-Host "WIF columns: $startCol ($($hdr[$startCol])) to $endCol ($($hdr[$endCol]))"

# --- Build Tab 2 (WIF): POR rows + WIF clones ---
$wifRows = New-Object System.Collections.Generic.List[object]
$wifRows.Add($hdr)
$rowRedCols = @{}

foreach ($mi in $wifMatches) {
    $wifRows.Add($filtered[$mi])
    
    $clone = @($filtered[$mi] | ForEach-Object { $_ })
    $clone[$porIdx] = 'WIF'
    for ($c = $startCol; $c -le $endCol; $c++) {
        if ($hdr[$c] -match '^\d{6}$') { $clone[$c] = $wifValue }
    }
    $wifRows.Add($clone)
    $excelRow = $wifRows.Count
    $redCols = @()
    for ($c = $startCol; $c -le $endCol; $c++) {
        if ($hdr[$c] -match '^\d{6}$') { $redCols += ($c + 1) }
    }
    $rowRedCols[$excelRow] = $redCols
}
Write-Host "WIF tab: $($wifRows.Count - 1) rows (POR + WIF)"

# --- Write CSVs ---
$csvBase  = Join-Path $env:TEMP "${safeName}_SDA_PPV_base.csv"
$csvWif   = Join-Path $env:TEMP "${safeName}_SDA_PPV_wif.csv"
$xlsxFile = Join-Path $env:TEMP "${safeName}_SDA_PPV_WIF.xlsx"

Tci-WriteCsv -Rows $filtered -Path $csvBase
Tci-WriteCsv -Rows $wifRows -Path $csvWif

# --- Build 2-tab xlsx ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false

try {
    $targetWb = $excel.Workbooks.Add()
    $targetWb.Worksheets.Item(1).Name = '_placeholder_'

    # Tab 1: PPV Baseline
    $wbBase = $excel.Workbooks.Open($csvBase)
    $wsBase = $wbBase.Worksheets.Item(1)
    $wsBase.Name = 'PPV Baseline'
    $wsBase.Move([System.Reflection.Missing]::Value, $targetWb.Worksheets.Item($targetWb.Worksheets.Count))
    try { $wbBase.Close($false) } catch {}

    # Tab 2: WIF
    $wbWif = $excel.Workbooks.Open($csvWif)
    $wsWif = $wbWif.Worksheets.Item(1)
    $wsWif.Name = 'WIF'
    $wsWif.Move([System.Reflection.Missing]::Value, $targetWb.Worksheets.Item($targetWb.Worksheets.Count))
    try { $wbWif.Close($false) } catch {}

    $targetWb.Worksheets.Item('_placeholder_').Delete()

    # Format Tab 1
    $ws1 = $targetWb.Worksheets.Item('PPV Baseline')
    $ws1.Activate()
    $lastCol1 = $ws1.UsedRange.Columns.Count
    $ws1.Range($ws1.Cells(1,1), $ws1.Cells(1,$lastCol1)).Font.Bold = $true
    $ws1.Range($ws1.Cells(1,1), $ws1.Cells(1,$lastCol1)).Interior.Color = 14277081
    try { $aw = $excel.ActiveWindow; $aw.SplitColumn = 4; $aw.SplitRow = 1; $aw.FreezePanes = $true } catch {}
    $ws1.Columns.AutoFit() | Out-Null

    # Format Tab 2
    $ws2 = $targetWb.Worksheets.Item('WIF')
    $ws2.Activate()
    $lastCol2 = $ws2.UsedRange.Columns.Count
    $ws2.Range($ws2.Cells(1,1), $ws2.Cells(1,$lastCol2)).Font.Bold = $true
    $ws2.Range($ws2.Cells(1,1), $ws2.Cells(1,$lastCol2)).Interior.Color = 14277081
    try { $aw2 = $excel.ActiveWindow; $aw2.SplitColumn = 4; $aw2.SplitRow = 1; $aw2.FreezePanes = $true } catch {}

    # WIF formatting
    $porColExcel = $porIdx + 1
    $totalRows2 = $wifRows.Count
    for ($r = 2; $r -le $totalRows2; $r++) {
        $porVal = $ws2.Cells($r, $porColExcel).Text
        if ($porVal -eq 'WIF') {
            $ws2.Range($ws2.Cells($r,1), $ws2.Cells($r,$lastCol2)).Interior.Color = 0x66FFFF
            if ($rowRedCols.ContainsKey($r)) {
                foreach ($col in $rowRedCols[$r]) {
                    $ws2.Cells($r, $col).Font.Color = 0x0000FF
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

# --- Print POR vs WIF ---
Write-Host "`n--- POR vs WIF ---"
foreach ($mi in $wifMatches) {
    $row = $filtered[$mi]
    $site = "$($row[$siteIdx])".Trim()
    $porVals = @()
    for ($c = $startCol; $c -le [Math]::Min($startCol+2, $endCol); $c++) {
        if ($hdr[$c] -match '^\d{6}$') { $porVals += "$($hdr[$c])=$($row[$c])" }
    }
    Write-Host "  Site=$site POR: $($porVals -join ', ') -> WIF: $wifValue"
}

# --- Email ---
$body = Build-WifCard -Product $product -Group $grp -SubGroup $sub -Report 'SDA Weekly' -Operation 'PPV' -Metric $wifSpec -WifValue "$wifValue kU" -TimeRange "$wifStart - $wifEnd" -PorRows $wifMatches.Count -WifRows $wifMatches.Count
$subj = "PHI of $product - SDA Weekly PPV + WIF SS"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxFile)
Write-Host "`nDone."
