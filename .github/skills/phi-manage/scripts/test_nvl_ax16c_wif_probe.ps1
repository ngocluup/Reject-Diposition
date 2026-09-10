$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

$afterFile = "$env:TEMP\tci_after_Nova_Lake_AX_16C_sdamo.html"
$html = Get-Content $afterFile -Raw

$tables = Tci-ParseTables $html
$dt = Tci-PickDataTable $tables 'CommonName'
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr = $allRows[0]

# Show header columns (time columns start after fixed cols)
Write-Host "Total cols: $($hdr.Count)"
Write-Host "Time columns (raw):"
for ($c = 20; $c -lt $hdr.Count; $c++) { Write-Host "  [$c] $($hdr[$c])" }

# Relabel to MMM yyyy
function WwToMonth($yyyyww) {
    if ($yyyyww -notmatch '^\d{6}$') { return $yyyyww }
    $y = [int]$yyyyww.Substring(0,4)
    $w = [int]$yyyyww.Substring(4,2)
    $jan1 = [datetime]::new($y,1,1)
    $ww01Sun = $jan1.AddDays(-[int]$jan1.DayOfWeek)
    $targetSun = $ww01Sun.AddDays(7*($w-1))
    return $targetSun.ToString("MMM yyyy")
}
Write-Host "`nRelabeled:"
for ($c = 20; $c -lt $hdr.Count; $c++) { Write-Host "  [$c] $($hdr[$c]) -> $(WwToMonth $hdr[$c])" }

# Find SubObject column
$soIdx = [array]::IndexOf([string[]]$hdr, 'SubObject')
$opIdx = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx = [array]::IndexOf([string[]]$hdr, 'MetricName')
$bomIdx = [array]::IndexOf([string[]]$hdr, 'BOM')
Write-Host "`nKey column indexes: Op=$opIdx SubObj=$soIdx Met=$metIdx BOM=$bomIdx"

# Find Classhot ETT rows
Write-Host "`nLooking for Classhot ETT (TEST_MPS + PBIC1 + TEST_TIME-SEC + BOM=blank):"
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $so  = if ($soIdx -ge 0) { "$($allRows[$i][$soIdx])".Trim() } else { '' }
    $met = "$($allRows[$i][$metIdx])".Trim()
    $bom = if ($bomIdx -ge 0) { "$($allRows[$i][$bomIdx])".Trim() } else { '' }
    if ($op -eq 'TEST_MPS' -and $so -eq 'PBIC1' -and $met -eq 'TEST_TIME-SEC' -and $bom -eq '') {
        Write-Host "  Row ${i}: Op=$op So=$so Met=$met BOM='$bom'" -ForegroundColor Green
    }
}
