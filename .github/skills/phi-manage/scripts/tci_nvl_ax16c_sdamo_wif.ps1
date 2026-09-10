$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Nova Lake AX 16C'
$safeName  = 'Nova_Lake_AX_16C'
$grp       = 'Client'
$sub       = 'Mobile'
$wifSpec   = 'Classhot ETT'
$wifValue  = '1001'
# WW30-WW32'26 spans Jul 2026 and Aug 2026 monthly buckets
$targetCols = @('Jul 2026','Aug 2026')

Write-Host "=== $product - WIF: $wifSpec $($targetCols -join ', ') = $wifValue ===" -ForegroundColor Cyan

# --- Load cached data ---
$afterFile = "$env:TEMP\tci_after_${safeName}_sdamo.html"
if (-not (Test-Path $afterFile)) { Write-Error "Cache not found: $afterFile"; exit 1 }
$html = Get-Content $afterFile -Raw

$tables = Tci-ParseTables $html
$dt = Tci-PickDataTable $tables 'CommonName'
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr = $allRows[0]

# Relabel headers to MMM yyyy
function WwToMonth($yyyyww) {
    if ($yyyyww -notmatch '^\d{6}$') { return $yyyyww }
    $y = [int]$yyyyww.Substring(0,4); $w = [int]$yyyyww.Substring(4,2)
    $jan1 = [datetime]::new($y,1,1)
    $ww01Sun = $jan1.AddDays(-[int]$jan1.DayOfWeek)
    $targetSun = $ww01Sun.AddDays(7*($w-1))
    return $targetSun.ToString("MMM yyyy")
}
$newHdr = @()
for ($c = 0; $c -lt $hdr.Count; $c++) { $newHdr += WwToMonth $hdr[$c] }

# Find key column indexes
$opIdx  = [array]::IndexOf([string[]]$newHdr, 'OperationName')
$soIdx  = [array]::IndexOf([string[]]$newHdr, 'SubObject')
$metIdx = [array]::IndexOf([string[]]$newHdr, 'MetricName')
$bomIdx = [array]::IndexOf([string[]]$newHdr, 'BOM')
$ppIdx  = [array]::IndexOf([string[]]$newHdr, 'Proposed/POR')

# Resolve target column indexes
$targetColIdx = @()
foreach ($tc in $targetCols) {
    $idx = [array]::IndexOf([string[]]$newHdr, $tc)
    if ($idx -lt 0) { Write-Error "Column '$tc' not found in headers"; exit 1 }
    $targetColIdx += $idx
}
Write-Host "Target columns: $($targetCols -join ', ') -> indexes $($targetColIdx -join ', ')"

# Find Classhot ETT rows: TEST_MPS + PBIC1 + TEST_TIME-SEC + BOM=blank
$porRows = @()
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $so  = "$($allRows[$i][$soIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim()
    $bom = "$($allRows[$i][$bomIdx])".Trim()
    if ($op -eq 'TEST_MPS' -and $so -eq 'PBIC1' -and $met -eq 'TEST_TIME-SEC' -and $bom -eq '') {
        $porRows += $i
    }
}
Write-Host "Matched POR rows: $($porRows.Count)"
if ($porRows.Count -eq 0) { Write-Error "No rows match Classhot ETT criteria"; exit 1 }

# Build output: header + POR rows + WIF clones
$output = New-Object System.Collections.Generic.List[object]
$output.Add($newHdr)

# Track which rows (in output) get red on which cols
$rowRedCols = @{}

foreach ($ri in $porRows) {
    # Add POR row
    $output.Add($allRows[$ri])
    # Clone for WIF
    $clone = $allRows[$ri].Clone()
    $clone[$ppIdx] = 'WIF'
    foreach ($ci in $targetColIdx) { $clone[$ci] = $wifValue }
    $output.Add($clone)
    $rowRedCols[$output.Count - 1] = $targetColIdx
}
Write-Host "Output rows: $($output.Count - 1) (POR=$($porRows.Count) + WIF=$($porRows.Count))"

# Show POR vs WIF comparison
Write-Host ""
Write-Host "POR vs WIF comparison:"
$compHdr = "  | Site |"
foreach ($tc in $targetCols) { $compHdr += " $tc POR | $tc WIF |" }
Write-Host $compHdr

$siteIdx = [array]::IndexOf([string[]]$newHdr, 'Site')
foreach ($ri in $porRows) {
    $site = if ($siteIdx -ge 0) { "$($allRows[$ri][$siteIdx])".Trim() } else { '?' }
    $line = "  | $site |"
    foreach ($ci in $targetColIdx) { $line += " $($allRows[$ri][$ci]) | $wifValue |" }
    Write-Host $line
}

# --- CSV + Excel ---
$csvFile  = Join-Path $env:TEMP "${safeName}_SDAMO_WIF.csv"
$xlsxFile = Join-Path $env:TEMP "${safeName}_SDAMO_WIF.xlsx"

Tci-WriteCsv -Rows $output -Path $csvFile
$result = Tci-ExportXlsx -CsvPath $csvFile -XlsxPath $xlsxFile -SheetName 'WIF' -FreezeColumns 4 -Validate -MinRows 1
Remove-Item $csvFile -Force -ErrorAction SilentlyContinue

# Apply red font + yellow bg to WIF target cells
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($xlsxFile)
$ws = $wb.Sheets.Item(1)
foreach ($rowKey in $rowRedCols.Keys) {
    $excelRow = $rowKey + 1  # 1-based, +1 for header
    foreach ($ci in $rowRedCols[$rowKey]) {
        $excelCol = $ci + 1  # 1-based
        $cell = $ws.Cells.Item($excelRow, $excelCol)
        $cell.Font.Color = 255  # Red
        $cell.Font.Bold = $true
        $cell.Interior.Color = 10092543  # Yellow (#FFFF99)
    }
}
$wb.Save(); $wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host "xlsx: $xlsxFile ($([Math]::Round((Get-Item $xlsxFile).Length/1KB, 1)) KB)"

# --- Email ---
$body = Build-WifCard -Product $product -Group $grp -SubGroup $sub -Report 'SDA Monthly' -Operation 'TEST_MPS' -Metric 'TEST_TIME-SEC' -WifValue $wifValue -TimeRange "$($targetCols -join ' - ')" -PorRows $porRows.Count -WifRows $porRows.Count
$subj = "PHI WIF - $product - $wifSpec"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxFile)
Write-Host "`nDone."
