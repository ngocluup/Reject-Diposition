$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

$afterFile = "$env:TEMP\tci_after_Nova_Lake_H_sda.html"
$html = Get-Content $afterFile -Raw

$tables = Tci-ParseTables $html
$dt = Tci-PickDataTable $tables 'CommonName'
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr = $allRows[0]

$opIdx = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx = [array]::IndexOf([string[]]$hdr, 'MetricName')

Write-Host "Total rows: $($allRows.Count - 1)"
Write-Host "`nAll unique OperationName values:"
$ops = @()
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $ops += "$($allRows[$i][$opIdx])".Trim()
}
$ops | Sort-Object -Unique | ForEach-Object { if ($_ -eq '') { Write-Host "  (blank)" } else { Write-Host "  $_" } }

Write-Host "`nTEST* rows: $(($ops | Where-Object { $_ -like 'TEST*' }).Count)"
Write-Host "PPV* rows: $(($ops | Where-Object { $_ -like 'PPV*' }).Count)"
Write-Host "BURNIN* rows: $(($ops | Where-Object { $_ -like 'BURNIN*' }).Count)"
Write-Host "Blank-op rows: $(($ops | Where-Object { $_ -eq '' }).Count)"
