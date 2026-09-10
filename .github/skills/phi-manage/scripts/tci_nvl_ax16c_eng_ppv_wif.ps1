$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Nova Lake AX 16C'
$safeName  = 'Nova_Lake_AX_16C'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'ENGForecasts'
$initTag   = 'eng'
$firstTh   = 'ATGroup'
$filter    = 'PPV'

# WIF config
$wifSpec   = 'PPVs ETT'
$wifTarget = 'QS'     # milestone column
$wifValue  = '25'

Write-Host "=== $product - ENG + PPV + WIF ($wifSpec $wifTarget=$wifValue) ===" -ForegroundColor Cyan

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

$opIdx  = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx = [array]::IndexOf([string[]]$hdr, 'MetricName')
$porIdx = [array]::IndexOf([string[]]$hdr, 'Proposed/POR')
Write-Host "Indexes: Op=$opIdx Met=$metIdx POR=$porIdx"
Write-Host "Headers: $($hdr -join ', ')"

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
if ($filtered.Count -le 1) { Write-Error "No rows after PPV filter"; exit 1 }

# --- WIF: find matching rows (PPVs ETT = PPV_SPS + TEST TIME - MIN) ---
$wifMatches = New-Object System.Collections.Generic.List[int]
for ($i = 1; $i -lt $filtered.Count; $i++) {
    $row = $filtered[$i]
    $op  = "$($row[$opIdx])".Trim()
    $met = "$($row[$metIdx])".Trim()
    if ($op -eq 'PPV_SPS' -and $met -eq 'TEST TIME - MIN') {
        $wifMatches.Add($i)
    }
}
Write-Host "WIF matches (PPVs ETT = PPV_SPS + TEST TIME - MIN): $($wifMatches.Count) rows"
if ($wifMatches.Count -eq 0) { Write-Error "No rows match PPVs ETT spec"; exit 1 }

# --- Resolve WIF target column (QS) ---
$targetCol = -1
for ($c = 0; $c -lt $hdr.Count; $c++) {
    if ($hdr[$c] -eq $wifTarget) { $targetCol = $c; break }
}
if ($targetCol -lt 0) { Write-Error "Column '$wifTarget' not found in headers"; exit 1 }
Write-Host "WIF column: $targetCol ($($hdr[$targetCol]))"

# --- Build Tab 2 (WIF): POR rows + WIF clones ---
$wifRows = New-Object System.Collections.Generic.List[object]
$wifRows.Add($hdr)
$rowRedCols = @{}

foreach ($mi in $wifMatches) {
    $wifRows.Add($filtered[$mi])
    
    $clone = @($filtered[$mi] | ForEach-Object { $_ })
    if ($porIdx -ge 0) { $clone[$porIdx] = 'WIF' }
    $clone[$targetCol] = $wifValue
    $wifRows.Add($clone)
    $excelRow = $wifRows.Count
    $rowRedCols[$excelRow] = @($targetCol + 1)  # 1-based for Excel
}
Write-Host "WIF tab: $($wifRows.Count - 1) rows (POR + WIF)"

# --- Write CSVs ---
$csvBase  = Join-Path $env:TEMP "${safeName}_ENG_PPV_base.csv"
$csvWif   = Join-Path $env:TEMP "${safeName}_ENG_PPV_wif.csv"
$xlsxFile = Join-Path $env:TEMP "${safeName}_ENG_PPV_WIF.xlsx"

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
    $porColExcel = if ($porIdx -ge 0) { $porIdx + 1 } else { 0 }
    $totalRows2 = $wifRows.Count
    for ($r = 2; $r -le $totalRows2; $r++) {
        $isWif = $false
        if ($porColExcel -gt 0) { $isWif = ($ws2.Cells($r, $porColExcel).Text -eq 'WIF') }
        else { $isWif = ($r % 2 -eq 1) }  # fallback: odd rows are WIF clones
        if ($isWif) {
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
$siteIdx = [array]::IndexOf([string[]]$hdr, 'Site')
$dlcpIdx = [array]::IndexOf([string[]]$hdr, 'DLCP')
foreach ($mi in $wifMatches) {
    $row = $filtered[$mi]
    $site = if ($siteIdx -ge 0) { "$($row[$siteIdx])".Trim() } else { '-' }
    $dlcp = if ($dlcpIdx -ge 0) { "$($row[$dlcpIdx])".Trim() } else { '' }
    $porVal = "$($row[$targetCol])".Trim()
    Write-Host "  Site=$site DLCP=$dlcp $wifTarget POR=$porVal -> WIF=$wifValue"
}

# --- Email ---
$body = Build-WifCard -Product $product -Group $grp -SubGroup $sub -Report 'ENG' -Operation 'PPV' -Metric $wifSpec -WifValue "$wifValue mins" -TimeRange $wifTarget -PorRows $wifMatches.Count -WifRows $wifMatches.Count
$subj = "PHI of $product - ENG PPV + WIF $wifSpec $wifTarget"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxFile)
Write-Host "`nDone."
