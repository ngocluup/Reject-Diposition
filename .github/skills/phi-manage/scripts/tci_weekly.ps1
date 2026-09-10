#requires -Version 5.1
<#
Scheduled weekly PHI tracking (improvement #4).

Reads a fixed product/report list from a plain-text file (default
tci_tracking_list.txt) and runs them all through tci_batch_driver.ps1 in one
batch. The driver logs every tab to tci_run_history.csv and Add-PhiRunHistory
already warns when a report's row count drops >20% week-over-week, so the
weekly e-mails double as a regression tripwire.

List file format - one job per line:
    <Product>[ | <Report> [ | <Filter> [ | <Email> ]]]
Blank lines and lines starting with # are ignored. A blank <Report> field
(just the product) runs all 4 reports as a single bundled workbook.

Examples:
  Nova Lake S 28C
  Nova Lake U | MOR | ppv
  CML PCH | SDA Weekly

Usage:
  .\tci_weekly.ps1                       # run the list now
  .\tci_weekly.ps1 -List my_list.txt     # custom list
  .\tci_weekly.ps1 -Register             # install Monday 07:00 scheduled task
  .\tci_weekly.ps1 -Unregister           # remove the scheduled task
#>
param(
  [string]$List = '',
  [string]$Driver = '',
  [switch]$Register,
  [switch]$Unregister,
  [string]$TaskName = 'TCI_PHI_Weekly',
  [string]$At = '07:00'
)
$ErrorActionPreference = 'Stop'
if (-not $List)   { $List   = Join-Path $PSScriptRoot 'tci_tracking_list.txt' }
if (-not $Driver) { $Driver = Join-Path $PSScriptRoot 'tci_batch_driver.ps1' }
$selfPath = Join-Path $PSScriptRoot 'tci_weekly.ps1'

# --- Scheduled-task management -------------------------------------------
if ($Register) {
  $action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
              -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$selfPath`" -List `"$List`""
  $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At $At
  $set     = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
      -Settings $set -Description 'Weekly TCI PHI tracking bundle' -Force | Out-Null
  Write-Host "Registered scheduled task '$TaskName' (Mondays $At). List: $List" -ForegroundColor Green
  return
}
if ($Unregister) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "Removed scheduled task '$TaskName'." -ForegroundColor Green
  return
}

# --- Build the batch input from the list ---------------------------------
if (-not (Test-Path $List)) {
  throw "Tracking list not found: $List. Create it (one product per line) or pass -List."
}
$jobs = @()
foreach ($line in (Get-Content $List)) {
  $t = $line.Trim()
  if ([string]::IsNullOrWhiteSpace($t) -or $t.StartsWith('#')) { continue }
  $p = $t -split '\s*\|\s*'
  $jobs += [pscustomobject]@{
    Product = $p[0].Trim()
    Report  = if ($p.Count -ge 2) { $p[1].Trim() } else { '' }
    Filter  = if ($p.Count -ge 3) { $p[2].Trim() } else { '' }
    Email   = if ($p.Count -ge 4) { $p[3].Trim() } else { '' }
  }
}
if ($jobs.Count -eq 0) { throw "No jobs parsed from $List." }
Write-Host "Weekly tracking: $($jobs.Count) job(s) from $List" -ForegroundColor Cyan

$tmp = Join-Path $env:TEMP ("tci_weekly_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd'))
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = 'Batch'
  $ws.Cells.Item(1,1) = 'Product'; $ws.Cells.Item(1,2) = 'Report'
  $ws.Cells.Item(1,3) = 'Filter';  $ws.Cells.Item(1,4) = 'Email'
  $r = 2
  foreach ($j in $jobs) {
    $ws.Cells.Item($r,1) = $j.Product
    $ws.Cells.Item($r,2) = $j.Report
    $ws.Cells.Item($r,3) = $j.Filter
    $ws.Cells.Item($r,4) = $j.Email
    $r++
  }
  $wb.SaveAs($tmp, 51); $wb.Close($false)
} finally {
  try { $xl.Quit() } catch {}
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

& $Driver -InputXlsx $tmp
