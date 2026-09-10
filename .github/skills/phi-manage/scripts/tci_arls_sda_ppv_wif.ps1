##############################################################################
# tci_arls_sda_ppv_wif.ps1
# ARL Refresh S 8C+16A+GT1 - SDA Weekly - PPV - WIF: PPVm ETT Q4'26 100MINS
# Output: only TEST TIME - MIN POR+WIF pairs for PPV_SPM
##############################################################################
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

. "$PSScriptRoot\tci_lib.ps1"

function Hidden($h,$n){ ([regex]::Match($h,'name="'+[regex]::Escape($n)+'"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value }
function Get-PanelCtl($html,$panelName,$labelExact){
  $pm=[regex]::Match($html,'(?is)id="ContentPlaceHolder1_Filters_'+[regex]::Escape($panelName)+'".*?</div>\s*</div>')
  if(-not $pm.Success){ throw "panel $panelName not found" }
  $items=[regex]::Matches($pm.Value,'(?is)<input[^>]*name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>\s*([^<]+?)\s*</label>')
  foreach($it in $items){ if($it.Groups[2].Value.Trim() -eq $labelExact){ return $it.Groups[1].Value } }
  throw "label '$labelExact' not found in $panelName"
}

# --- Config ---
$product = 'ARL Refresh S 8C+16A+GT1'
$grp = 'Client'; $sub = 'Desktop'
$reportId = 'PORSDAForecast'
$initTag = 'sda'
$tag = 'ARL_S8161LGA_Refresh'
$after = "$env:TEMP\tci_after_${tag}_${initTag}.html"

Write-Host "=== ARL-S8161LGA Refresh - SDA Weekly + PPV + WIF: PPVm ETT Q4'26 100MINS ==="

# --- Step 1: Fetch SDA Weekly (or reuse cache <12h) ---
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
  $rr = Tci-Post -ReportId $reportId -Form $f -OutFile $after -TimeoutSec 600
  Write-Host "[fetch] saved $([math]::Round((Get-Item $after).Length/1KB)) KB"
}

# --- Step 2: Parse HTML ---
$c = Get-Content $after -Raw
if ($c -match 'delivered no data') { throw 'TCI returned no data' }
$mainStart = $c.IndexOf('id="tbl_Main"')
if ($mainStart -ge 0) { $c = $c.Substring($mainStart) }

$tables = Tci-ParseTables -Html $c
$dt = Tci-PickDataTable -Tables $tables -FirstTh 'CommonName'
$parsed = Tci-RowsFromTable -DataTableBodies $dt
$hdr = $parsed.Header; $all = $parsed.Rows
Write-Host "total rows: $($all.Count - 1) cols: $($hdr.Count)"

# --- Column indexes ---
$iOp  = [array]::IndexOf($hdr, 'OperationName')
$iMet = [array]::IndexOf($hdr, 'MetricName')
$iPP  = [array]::IndexOf($hdr, 'Proposed/POR')
Write-Host "Indexes: Op=$iOp Met=$iMet PP=$iPP"

# --- PPV_SPM + TEST TIME - MIN filter ---
$ettLabel = 'TEST TIME - MIN'
$ppvmRows = New-Object System.Collections.Generic.List[object]
for ($k = 1; $k -lt $all.Count; $k++) {
  $row = $all[$k]
  $op  = ($row[$iOp] + '').Trim()
  $met = ($row[$iMet] + '').Trim()
  if ($op -eq 'PPV_SPM' -and $met -eq $ettLabel) { $ppvmRows.Add($row) }
}
Write-Host "PPV_SPM '$ettLabel' rows: $($ppvmRows.Count)"
if ($ppvmRows.Count -eq 0) { throw "No PPV_SPM TEST TIME - MIN rows found" }

# --- Identify Q4'26 WW columns (Oct-Dec 2026) ---
$targetMonths = @('Oct 2026','Nov 2026','Dec 2026')
$targetCols = New-Object System.Collections.Generic.List[int]
for ($ci = 0; $ci -lt $hdr.Count; $ci++) {
  if ($hdr[$ci] -match '^\d{6}$') {
    $ml = WwToFiscalMonth $hdr[$ci]
    if ($targetMonths -contains $ml) { $targetCols.Add($ci) }
  }
}
Write-Host "Q4'26 target columns: $($targetCols.Count) (WW$($hdr[$targetCols[0]])-WW$($hdr[$targetCols[-1]]))"
if ($targetCols.Count -eq 0) { throw 'No Q4 2026 columns found' }

# --- Build output: POR + WIF pairs ---
$output = New-Object System.Collections.Generic.List[object]
$output.Add($hdr)
$rowRedCols = @{}
$totalClones = 0
foreach ($row in $ppvmRows) {
  $output.Add($row)   # POR row
  $clone = [string[]]::new($row.Length)
  for ($i = 0; $i -lt $row.Length; $i++) { $clone[$i] = $row[$i] }
  $clone[$iPP] = 'WIF'
  foreach ($ci in $targetCols) { $clone[$ci] = '100' }
  $output.Add($clone)
  $rowRedCols[$output.Count] = $targetCols.ToArray()
  $totalClones++
}
Write-Host "Output: $($output.Count - 1) rows ($totalClones POR+WIF pairs)"

# --- Write CSV ---
$outBase = 'ARL_S_Refresh_SDA_PPVm_WIF'
$csv  = "$env:TEMP\$outBase.csv"
$xlsx = "$env:TEMP\$outBase.xlsx"
$maxCols = ($output | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
$sb = New-Object System.Text.StringBuilder
foreach ($row in $output) {
  $p = @($row) + (, '') * ($maxCols - $row.Count)
  $e = $p | ForEach-Object { $v = "$_"; if ($v -match '[",\r\n]') { '"' + ($v -replace '"', '""') + '"' } else { $v } }
  [void]$sb.AppendLine(($e -join ','))
}
[IO.File]::WriteAllText($csv, $sb.ToString(), [Text.UTF8Encoding]::new($true))

# --- Excel formatting ---
$excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($csv); $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'ARL-S PPVm WIF'
  $lastCol = $ws.UsedRange.Columns.Count
  $r = $ws.Range($ws.Cells(1, 1), $ws.Cells(1, $lastCol))
  $r.Font.Bold = $true; $r.Interior.Color = 0xD9D9D9
  $ppCol = $iPP + 1
  for ($row = 2; $row -le $output.Count; $row++) {
    if ($ws.Cells($row, $ppCol).Text -eq 'WIF') {
      $ws.Range($ws.Cells($row, 1), $ws.Cells($row, $lastCol)).Interior.Color = 0x66FFFF
      if ($rowRedCols.ContainsKey($row)) {
        foreach ($ci in $rowRedCols[$row]) { $ws.Cells($row, $ci + 1).Font.Color = 0x0000FF }
      }
    }
  }
  $ws.Application.ActiveWindow.SplitColumn = 4
  $ws.Application.ActiveWindow.SplitRow = 1
  $ws.Application.ActiveWindow.FreezePanes = $true
  $ws.Columns.AutoFit() | Out-Null
  if (Test-Path $xlsx) { Remove-Item $xlsx -Force }
  $wb.SaveAs($xlsx, 51); try { $wb.Close($false) } catch {}
} finally {
  try { $excel.Quit() } catch {}
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Host "xlsx: $xlsx ($([math]::Round((Get-Item $xlsx).Length/1KB,1)) KB)"

# --- Email ---
$body = Build-WifCard -Product $product -Group $grp -SubGroup $sub `
  -Report 'SDA Weekly' -Operation 'PPV_SPM' -Metric 'TEST TIME - MIN' `
  -WifValue '100 MINS' -TimeRange 'Q4 2026 (Oct - Dec)' `
  -Columns "13 weekly columns (WW40-WW52)" `
  -PorRows $totalClones -WifRows $totalClones
Tci-SendMail -Subject "TCI ARL-S8161LGA Refresh - SDA Weekly PPV - WIF: PPVm ETT Q4'26 100MINS" -HtmlBody $body -Attachments @($xlsx)
Write-Host "`nDone."
