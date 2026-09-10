$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# Read cached data to find monthly column headers
$afterFile = "$env:TEMP\tci_after_Nova_Lake_AX_16C_sdamo.html"
$html = Get-Content $afterFile -Raw
$tables = Tci-ParseTables $html
$dt = Tci-PickDataTable $tables 'CommonName'
$parsed = Tci-RowsFromTable $dt
$hdr = $parsed.Rows[0]

# Show raw time columns (YYYYWW codes)
Write-Host "Time columns (raw YYYYWW -> MMM YYYY):"
for ($c = 0; $c -lt $hdr.Count; $c++) {
    if ($hdr[$c] -match '^\d{6}$') {
        $ww = $hdr[$c]
        $y = [int]$ww.Substring(0,4); $w = [int]$ww.Substring(4,2)
        $jan1 = [datetime]::new($y,1,1)
        $ww01Sun = $jan1.AddDays(-[int]$jan1.DayOfWeek)
        $targetSun = $ww01Sun.AddDays(7*($w-1))
        $month = $targetSun.ToString("MMM yyyy")
        Write-Host "  col $c : $ww -> $month"
    }
}

# Find which column(s) cover WW30-32 of 2026
Write-Host "`nLooking for columns covering WW30-32 2026..."
Write-Host "WW30 anchor Sunday:"
$jan1 = [datetime]::new(2026,1,1)
$ww01Sun = $jan1.AddDays(-[int]$jan1.DayOfWeek)
$ww30Sun = $ww01Sun.AddDays(7*29)
$ww32Sun = $ww01Sun.AddDays(7*31)
Write-Host "  WW30 = $($ww30Sun.ToString('yyyy-MM-dd')) ($($ww30Sun.ToString('MMM yyyy')))"
Write-Host "  WW32 = $($ww32Sun.ToString('yyyy-MM-dd')) ($($ww32Sun.ToString('MMM yyyy')))"
