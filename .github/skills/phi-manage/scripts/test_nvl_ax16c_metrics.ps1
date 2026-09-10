$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

$afterFile = "$env:TEMP\tci_after_Nova_Lake_AX_16C_sdamo.html"
$html = Get-Content $afterFile -Raw
$tables = Tci-ParseTables $html
$dt = Tci-PickDataTable $tables 'CommonName'
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr = $allRows[0]
$opIdx = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx = [array]::IndexOf([string[]]$hdr, 'MetricName')

$classMonitors = @('MPS','EQA','CS MONITOR')
Write-Host "Class rows - MetricName values:"
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim()
    $isClass = ($op -ne '' -and $op -like 'TEST*') -or
               ($op -eq '' -and ($classMonitors | Where-Object { $met -match "(?i)^$_" }).Count -gt 0)
    if ($isClass) { Write-Host "  [$op] $met" }
}
