$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Nova Lake AX 16C'
$safeName  = 'Nova_Lake_AX_16C'
$grp       = 'Client'
$sub       = 'Mobile'
$filter    = 'Class'
$wifSpec   = 'Classhot ETT'
$wifValue  = '1001'
$wifRange  = 'WW30-WW32 2026'  # maps to Jul 2026, Aug 2026

Write-Host "=== $product - SDA Monthly Class + WIF ===" -ForegroundColor Cyan

# --- Read cached data ---
$afterFile = "$env:TEMP\tci_after_${safeName}_sdamo.html"
if (-not (Test-Path $afterFile)) { Write-Error "Cache not found: $afterFile"; exit 1 }
$html = Get-Content $afterFile -Raw

$tables = Tci-ParseTables $html
$dt = Tci-PickDataTable $tables 'CommonName'
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr = $allRows[0]

$opIdx  = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx = [array]::IndexOf([string[]]$hdr, 'MetricName')
$ppIdx  = [array]::IndexOf([string[]]$hdr, 'Proposed/POR')
if ($ppIdx -lt 0) { for ($c = 0; $c -lt $hdr.Count; $c++) { if ($hdr[$c] -match 'Proposed') { $ppIdx = $c; break } } }

# --- Class filter ---
$classMonitors = @('MPS','EQA','CS MONITOR')
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim()
    $isClass = ($op -ne '' -and $op -like 'TEST*') -or
               ($op -eq '' -and ($classMonitors | Where-Object { $met -match "(?i)^$_" }).Count -gt 0)
    if ($isClass) { $filtered.Add($allRows[$i]) }
}
Write-Host "Class filter: $($filtered.Count - 1) rows"

# --- Find WIF target rows (Classhot ETT) ---
$wifTargetIdxs = @()
for ($i = 1; $i -lt $filtered.Count; $i++) {
    $met = "$($filtered[$i][$metIdx])".Trim()
    if ($met -match '(?i)^Classhot\s+ETT') { $wifTargetIdxs += $i }
}
Write-Host "WIF target rows (Classhot ETT): $($wifTargetIdxs.Count)"
if ($wifTargetIdxs.Count -eq 0) { Write-Error "No Classhot ETT rows found"; exit 1 }

# --- Find WIF columns (Jul 2026 = col with 202631, Aug 2026 = col with 202635) ---
$wifCols = @()
for ($c = 0; $c -lt $hdr.Count; $c++) {
    if ($hdr[$c] -eq '202631' -or $hdr[$c] -eq '202635') { $wifCols += $c }
}
Write-Host "WIF columns: $($wifCols -join ', ') (Jul-Aug 2026)"

# --- Clone WIF rows ---
$wifRows = New-Object System.Collections.Generic.List[object]
foreach ($idx in $wifTargetIdxs) {
    $srcRow = $filtered[$idx]
    $newRow = @($srcRow)  # shallow copy array
    # Mark as WIF in Proposed/POR column
    if ($ppIdx -ge 0) { $newRow[$ppIdx] = 'WIF' }
    # Set value in WIF columns
    foreach ($c in $wifCols) { $newRow[$c] = $wifValue }
    $wifRows.Add($newRow)
}
Write-Host "WIF rows created: $($wifRows.Count)"

# --- Build final output: base filtered + WIF rows appended ---
$output = New-Object System.Collections.Generic.List[object]
# Relabel monthly headers
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
$output.Add($newHdr)

# Add base rows
for ($i = 1; $i -lt $filtered.Count; $i++) { $output.Add($filtered[$i]) }
# Add WIF rows
foreach ($wr in $wifRows) { $output.Add($wr) }
Write-Host "Total output: $($output.Count - 1) rows ($($filtered.Count - 1) base + $($wifRows.Count) WIF)"

# --- CSV + Excel ---
$csvFile  = Join-Path $env:TEMP "${safeName}_SDAMO_Class_WIF.csv"
$xlsxFile = Join-Path $env:TEMP "${safeName}_SDAMO_Class_WIF.xlsx"

Tci-WriteCsv -Rows $output -Path $csvFile
$result = Tci-ExportXlsx -CsvPath $csvFile -XlsxPath $xlsxFile -SheetName 'SDA Monthly Class+WIF' -FreezeColumns 4 -Validate -MinRows 1
Remove-Item $csvFile -Force -ErrorAction SilentlyContinue
Write-Host "xlsx: $xlsxFile ($([Math]::Round((Get-Item $xlsxFile).Length/1KB, 1)) KB)"

# --- Email ---
$sections = @( @{ SheetTag = 'SDA Monthly Class'; RowCount = $filtered.Count - 1 } )
$wifSpecs = @( @{ Spec = $wifSpec; Range = $wifRange; Value = $wifValue } )
$body = Build-WifCard -Product $product -Group $grp -SubGroup $sub -Sections $sections -Filter $filter -WifSpecs $wifSpecs
$subj = "PHI WIF of $product - SDA Monthly Class - $wifSpec"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxFile)
Write-Host "`nDone."
