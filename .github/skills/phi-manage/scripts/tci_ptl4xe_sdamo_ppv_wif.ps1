##############################################################################
# tci_ptl4xe_sdamo_ppv_wif.ps1
# PTL H 4Xe - SDA Monthly - PPV filter - WIF: PPVs ETT July'26-Jan'27 100MINS
##############################################################################
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# --- Source shared lib ---
. "$PSScriptRoot\tci_lib.ps1"

# --- Inline helpers (from batch driver) ---
function Hidden($h,$n){ ([regex]::Match($h,'name="'+[regex]::Escape($n)+'"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value }
function Get-PanelCtl($html,$panelName,$labelExact){
  $pm=[regex]::Match($html,'(?is)id="ContentPlaceHolder1_Filters_'+[regex]::Escape($panelName)+'".*?</div>\s*</div>')
  if(-not $pm.Success){ throw "panel $panelName not found" }
  $items=[regex]::Matches($pm.Value,'(?is)<input[^>]*name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>\s*([^<]+?)\s*</label>')
  foreach($it in $items){ if($it.Groups[2].Value.Trim() -eq $labelExact){ return $it.Groups[1].Value } }
  throw "label '$labelExact' not found in $panelName"
}

# --- Config ---
$product = 'Panther Lake H 4P+8E+4LP_E+4Xe'
$grp = 'Client'; $sub = 'Mobile'
$reportId = 'PORMonthlyForecast'
$initTag = 'sdamo'
$tag = ($product -replace '[\\/:*?"<>|+]','_') -replace '\s+','_'
$after = "$env:TEMP\tci_after_${tag}_${initTag}.html"

Write-Host "=== PTL H 4Xe - SDA Monthly + PPV + WIF: ETT Jul'26-Jan'27 100MINS ==="

# --- Step 1: Fetch SDA Monthly (or reuse cache) ---
if ((Test-Path $after) -and (Get-Item $after).LastWriteTime -gt (Get-Date).AddHours(-12)) {
  Write-Host "[cache] reusing $after (age $([int]((Get-Date)-(Get-Item $after).LastWriteTime).TotalMinutes)m)"
} else {
  Write-Host "[fetch] GET init page..."
  $h = Tci-GetInit -ReportId $reportId -MaxAgeMinutes 720
  $vs  = Hidden $h '__VIEWSTATE'
  $vsg = Hidden $h '__VIEWSTATEGENERATOR'
  $ev  = Hidden $h '__EVENTVALIDATION'
  $cGrp = Get-PanelCtl $h 'AT_Group' $grp
  $cSub = Get-PanelCtl $h 'AT_SubGroup' $sub
  $cCn  = Get-PanelCtl $h 'CommonName' $product
  $panelPhi = [regex]::Match($h,'(?is)id="ContentPlaceHolder1_Filters_PHIParameters".*?</div>\s*</div>').Value
  $phi = [regex]::Matches($panelPhi,'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }
  Write-Host "[fetch] POST $reportId (PHI=$($phi.Count) params)..."
  $f = [ordered]@{
    '__EVENTTARGET'=''; '__EVENTARGUMENT'=''
    '__VIEWSTATE'=$vs; '__VIEWSTATEGENERATOR'=$vsg; '__EVENTVALIDATION'=$ev
    $cGrp='on'; $cSub='on'; $cCn='on'
    'ctl00$ContentPlaceHolder1$btn_RunReport'='Run Report'
  }
  foreach($p in $phi){ $f[$p]='on' }
  $rr = Tci-Post -ReportId $reportId -Form $f -OutFile $after -TimeoutSec 480
  Write-Host "[fetch] saved $([math]::Round((Get-Item $after).Length/1KB)) KB"
}

# --- Step 2: Parse HTML ---
$c = Get-Content $after -Raw
if ($c -match 'delivered no data') { throw 'TCI returned no data for this product' }
$mainStart = $c.IndexOf('id="tbl_Main"')
if ($mainStart -ge 0) { $c = $c.Substring($mainStart) }

$tables = Tci-ParseTables -Html $c
$dt = Tci-PickDataTable -Tables $tables -FirstTh 'CommonName'
$parsed = Tci-RowsFromTable -DataTableBodies $dt
$hdr = $parsed.Header; $all = $parsed.Rows
Write-Host "total rows: $($all.Count - 1) cols: $($hdr.Count)"

# ===== Column indexes =====
$iOp  = [array]::IndexOf($hdr, 'OperationName')
$iMet = [array]::IndexOf($hdr, 'MetricName')
$iPP  = [array]::IndexOf($hdr, 'Proposed/POR')
Write-Host "Indexes: Op=$iOp Met=$iMet PP=$iPP"

# ===== PPV filter =====
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
for ($k = 1; $k -lt $all.Count; $k++) {
  $row = $all[$k]
  $op  = if ($iOp -ge 0 -and $iOp -lt $row.Count) { ($row[$iOp] + '').Trim() } else { '' }
  $met = if ($iMet -ge 0 -and $iMet -lt $row.Count) { ($row[$iMet] + '').Trim().ToUpper() } else { '' }
  if ($op.ToUpper().StartsWith('PPV')) { $filtered.Add($row) }
  elseif ([string]::IsNullOrWhiteSpace($op) -and $met -eq 'PPV-M SAMPLE SIZE') { $filtered.Add($row) }
}
Write-Host "PPV filtered rows: $($filtered.Count - 1)"
if ($filtered.Count -le 1) { throw 'No PPV rows found' }

# ===== Identify ETT metric rows =====
# Probe unique MetricNames to find exact ETT label
$metrics = @{}
for ($k = 1; $k -lt $filtered.Count; $k++) {
  $met = ($filtered[$k][$iMet] + '').Trim()
  if (-not $metrics.ContainsKey($met)) { $metrics[$met] = 0 }
  $metrics[$met]++
}
Write-Host "`nUnique metrics in PPV data:"
$metrics.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host "  $($_.Key) ($($_.Value) rows)" }

# Find ETT metric = "TEST TIME - MIN" for PPV operations
$ettLabel = $null
foreach ($m in $metrics.Keys) {
  if ($m -eq 'TEST TIME - MIN') { $ettLabel = $m; break }
}
if (-not $ettLabel) {
  # Fallback: exact "ETT" or "ESTIMATED TEST TIME"
  foreach ($m in $metrics.Keys) {
    if ($m -match '(?i)^(ETT|ESTIMATED[\s_]*TEST[\s_]*TIME)$') { $ettLabel = $m; break }
  }
}
if (-not $ettLabel) { throw "Could not find ETT metric. Available: $($metrics.Keys -join ', ')" }
Write-Host "`nMatched ETT metric: '$ettLabel'"

# ===== Identify target month columns (Jul 2026 - Jan 2027) =====
# SDA Monthly columns are WW anchor codes. Map each to fiscal month and select Jul'26-Jan'27.
$targetMonths = @('Jul 2026','Aug 2026','Sep 2026','Oct 2026','Nov 2026','Dec 2026','Jan 2027')
$targetCols = New-Object System.Collections.Generic.List[int]
$colMapping = @{}  # colIndex -> monthLabel (for display)

for ($ci = 0; $ci -lt $hdr.Count; $ci++) {
  if ($hdr[$ci] -match '^\d{6}$') {
    $ml = WwToFiscalMonth $hdr[$ci]
    if ($targetMonths -contains $ml) {
      $targetCols.Add($ci)
      $colMapping[$ci] = $ml
    }
  }
}
Write-Host "Target columns ($($targetCols.Count)):"
foreach ($ci in $targetCols) { Write-Host "  col[$ci] WW=$($hdr[$ci]) -> $($colMapping[$ci])" }
if ($targetCols.Count -eq 0) { throw 'No columns found for Jul 2026 - Jan 2027' }

# ===== Apply WIF: keep only ETT rows (POR + WIF clones) =====
$ettRows = New-Object System.Collections.Generic.List[object]
$ettRows.Add($hdr)
$rowRedCols = @{}
$totalClones = 0
for ($k = 1; $k -lt $filtered.Count; $k++) {
  $row = $filtered[$k]
  $met = ($row[$iMet] + '').Trim()
  if ($met -ne $ettLabel) { continue }
  # Add original POR row
  $ettRows.Add($row)
  # Clone as WIF
  $clone = [string[]]::new($row.Length)
  for ($i = 0; $i -lt $row.Length; $i++) { $clone[$i] = $row[$i] }
  $clone[$iPP] = 'WIF'
  foreach ($ci in $targetCols) { $clone[$ci] = '100' }
  $ettRows.Add($clone)
  $rowRedCols[$ettRows.Count] = $targetCols.ToArray()
  $totalClones++
  $op = ($row[$iOp] + '').Trim()
  Write-Host "  POR+WIF: Op=$op Met=$met"
}
Write-Host "`nWIF clones: $totalClones (output: $($ettRows.Count - 1) rows = POR + WIF pairs)"
if ($totalClones -eq 0) { throw "No ETT rows found in PPV data to clone" }
$filtered = $ettRows

# ===== Relabel WW columns to month names =====
$fixedCols = @('CommonName','MetricName','ItemID','ItemSegment','ENGType','ATREVStartDate','ATGroup','ATSubGroup','OperationName','ResourceName','STRGC','DLCP','FunctionalCore','GraphicsCore','PackageSize','PackageTech','Rev','Step','TestFlow','BinConfigName','Proposed/POR','BOM','Site','SubObject')
for ($ci = 0; $ci -lt $hdr.Count; $ci++) {
  if ($hdr[$ci] -match '^\d{6}$' -and $hdr[$ci] -notin $fixedCols) { $hdr[$ci] = WwToFiscalMonth $hdr[$ci] }
}
$filtered[0] = $hdr

# ===== Write CSV =====
$outBase = 'PTL_H_4Xe_SDAMonthly_PPV_WIF'
$csv  = "$env:TEMP\$outBase.csv"
$xlsx = "$env:TEMP\$outBase.xlsx"
$maxCols = ($filtered | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
$sb = New-Object System.Text.StringBuilder
foreach ($row in $filtered) {
  $p = @($row) + (, '') * ($maxCols - $row.Count)
  $e = $p | ForEach-Object { $v = "$_"; if ($v -match '[",\r\n]') { '"' + ($v -replace '"', '""') + '"' } else { $v } }
  [void]$sb.AppendLine(($e -join ','))
}
[IO.File]::WriteAllText($csv, $sb.ToString(), [Text.UTF8Encoding]::new($true))
Write-Host "CSV: $csv"

# ===== Excel formatting =====
$excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($csv); $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'PTL 4Xe PPV WIF'
  $lastCol = $ws.UsedRange.Columns.Count
  # Header styling
  $r = $ws.Range($ws.Cells(1, 1), $ws.Cells(1, $lastCol))
  $r.Font.Bold = $true; $r.Interior.Color = 0xD9D9D9
  # WIF rows: yellow background + blue font on changed cells
  $ppCol = $iPP + 1
  for ($row = 2; $row -le $filtered.Count; $row++) {
    if ($ws.Cells($row, $ppCol).Text -eq 'WIF') {
      $ws.Range($ws.Cells($row, 1), $ws.Cells($row, $lastCol)).Interior.Color = 0x66FFFF
      if ($rowRedCols.ContainsKey($row)) {
        foreach ($ci in $rowRedCols[$row]) { $ws.Cells($row, $ci + 1).Font.Color = 0x0000FF }
      }
    }
  }
  # Freeze panes
  $ws.Application.ActiveWindow.SplitColumn = 4
  $ws.Application.ActiveWindow.SplitRow = 1
  $ws.Application.ActiveWindow.FreezePanes = $true
  $ws.Columns.AutoFit() | Out-Null
  if (Test-Path $xlsx) { Remove-Item $xlsx -Force }
  $wb.SaveAs($xlsx, 51)
  try { $wb.Close($false) } catch {}
} finally {
  try { $excel.Quit() } catch {}
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Host "xlsx: $xlsx ($([math]::Round((Get-Item $xlsx).Length/1KB,1)) KB)"

# ===== Email =====
$body = Build-WifCard -Product $product -Group $grp -SubGroup $sub `
  -Report 'SDA Monthly' -Operation 'PPV (SPS + SPM)' -Metric 'TEST TIME - MIN' `
  -WifValue '100 MINS' -TimeRange "Jul 2026 - Jan 2027" `
  -Columns "7 monthly columns" `
  -PorRows $totalClones -WifRows $totalClones
Tci-SendMail -Subject "TCI PTL H 4Xe - SDA Monthly PPV - WIF: PPVs ETT Jul'26-Jan'27 100MINS" -HtmlBody $body -Attachments @($xlsx)
Write-Host "`nDone."
