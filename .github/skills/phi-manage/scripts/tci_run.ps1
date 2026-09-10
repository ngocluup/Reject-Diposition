#requires -Version 5.1
<#
Run-Phi one-shot wrapper.
Runs one or more TCI PHI report jobs (fetch + filter + xlsx + email) without
hand-building a batch xlsx. Reuses tci_batch_driver.ps1 so all fetch/filter/card
history logic stays in one place (single source of truth). Passes jobs directly
to the driver; no temporary Excel input workbook is created.

Examples:
  .\tci_run.ps1 -Product 'CML PCH' -Report 'SDA Weekly'
  .\tci_run.ps1 -Product 'CML PCH' -Report 'SDA Weekly' -Filter Class
  .\tci_run.ps1 -Product 'CML PCH','Nova Lake U' -Report 'SDA Weekly'
  .\tci_run.ps1 -Product 'CML PCH; Nova Lake U' -Report 'SDA Weekly'
  .\tci_run.ps1 -Product 'ARL Refresh S 8C+16A+GT1'        # blank Report = all 4 tabs
  .\tci_run.ps1 -Product 'PTL U404' -Report SDA -Filter yield -Email someone@intel.com
#>
param(
  [Parameter(Mandatory)][string[]]$Product,
  [string]$Report = '',
  [string]$Filter = '',
  [string]$Email  = '',
  [string]$Driver = '',
  [int]$MaxCacheHours = 12,
  [switch]$NoParallel,
  [switch]$CombineXlsx,
  [switch]$NoEmail
)
$ErrorActionPreference = 'Stop'
if (-not $Driver) { $Driver = Join-Path $PSScriptRoot 'tci_batch_driver.ps1' }

& $Driver -Product $Product -Report $Report -Filter $Filter -Email $Email -MaxCacheHours $MaxCacheHours -NoParallel:$NoParallel -CombineXlsx:$CombineXlsx -NoEmail:$NoEmail
