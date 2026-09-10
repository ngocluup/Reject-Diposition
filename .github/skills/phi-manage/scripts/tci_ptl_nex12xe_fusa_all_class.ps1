##############################################################################
# PTL H (NEX) 4P+8E+4LP_E+12Xe FuSa - All reports (SDA Weekly, SDA Monthly,
# MOR Spread, ENG) with Class filter -> single multi-tab xlsx + polished email.
##############################################################################
. "$env:USERPROFILE\Downloads\PHI Tracking\tci_lib.ps1"
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

$ProductLabel = 'Panther Lake H (NEX) 4P+8E+4LP_E+12Xe FuSa'
$ShortName    = 'PTL H (NEX) 12Xe FuSa'
$Group        = 'Client'
$SubGroup     = 'Mobile'
$OutBase      = 'PTL_NEX_12Xe_FuSa_AllReports_Class'

# report list: Label | ReportId | FirstTh | relabel mode
$reports = @(
  @{ Label='SDA Weekly';  Id='PORSDAForecast';     Th='CommonName'; Relabel='none'  },
  @{ Label='SDA Monthly'; Id='PORMonthlyForecast'; Th='CommonName'; Relabel='month' },
  @{ Label='MOR Spread';  Id='MORSpread';          Th='MetricName'; Relabel='quarter'},
  @{ Label='ENG';         Id='ENGForecasts';       Th='ATGroup';    Relabel='none'  }
)

$anchors = @(1,5,9,14,18,22,27,31,35,40,44,48)
$months  = @('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')
function WwToFiscalMonth($yyyyww){
  if($yyyyww -notmatch '^\d{6}$'){ return $yyyyww }
  $y=[int]$yyyyww.Substring(0,4); $w=[int]$yyyyww.Substring(4,2)
  $mi=-1; for($i=0;$i -lt 12;$i++){ if($w -ge $anchors[$i]){ $mi=$i } }
  if($mi -lt 0){ return $yyyyww }
  '{0} {1}' -f $months[$mi], $y
}

function Get-CB($h, $panelId, $label) {
  $i = $h.IndexOf('id="' + $panelId + '"'); if ($i -lt 0) { return $null }
  $j = $h.IndexOf('</table>', $i); $s = $h.Substring($i, $j - $i)
  $m = [regex]::Match($s, '<input id="([^"]+)" type="checkbox" name="([^"]+)"[^/]*/><label for="\1">' + [regex]::Escape($label) + '</label>')
  if ($m.Success) { $m.Groups[2].Value }
}
function Get-AllCB($h, $panelId) {
  $i = $h.IndexOf('id="' + $panelId + '"'); if ($i -lt 0) { return @() }
  $j = $h.IndexOf('</table>', $i); $s = $h.Substring($i, $j - $i)
  [regex]::Matches($s, '<input id="([^"]+)" type="checkbox" name="([^"]+)"[^/]*/><label for="\1">([^<]+)</label>') |
    ForEach-Object { [PSCustomObject]@{ Name = $_.Groups[2].Value; Label = $_.Groups[3].Value.Trim() } }
}
function HF($h, $n) {
  $m = [regex]::Match($h, '(?is)<input[^>]*name="' + [regex]::Escape($n) + '"[^>]*value="([^"]*)"')
  if ($m.Success) { $m.Groups[1].Value } else {
    $m2 = [regex]::Match($h, '(?is)<input[^>]*value="([^"]*)"[^>]*name="' + [regex]::Escape($n) + '"')
    if ($m2.Success) { $m2.Groups[1].Value } else { "" }
  }
}

$monitors = @('MPS MONITOR','EQA MONITOR','CS MONITOR')
$results = @()  # per-report: Label, Csv, SheetName, BreakdownHash, Total, Unfiltered

foreach ($rep in $reports) {
  Write-Host "`n===== $($rep.Label) ($($rep.Id)) ====="
  $cacheFile = Join-Path $env:TEMP "tci_after_ptlnex12xe_fusa_$($rep.Id).html"

  if (Test-Path $cacheFile) {
    Write-Host "[cache] reuse $cacheFile"
    $html = Get-Content $cacheFile -Raw
  } else {
    $init = Tci-GetInit -ReportId $rep.Id
    $cbG  = Get-CB $init 'ContentPlaceHolder1_Filters_AT_Group' $Group
    $cbS  = Get-CB $init 'ContentPlaceHolder1_Filters_AT_SubGroup' $SubGroup
    $cbP  = Get-CB $init 'ContentPlaceHolder1_Filters_CommonName' $ProductLabel
    $phis = Get-AllCB $init 'ContentPlaceHolder1_Filters_PHIParameters'
    if (-not $cbP) { throw "CommonName '$ProductLabel' not found for $($rep.Id)" }
    Write-Host "  ctls: G=$cbG S=$cbS P=$cbP PHI=$($phis.Count)"
    $form = [ordered]@{
      '__EVENTTARGET'=''; '__EVENTARGUMENT'='';
      '__VIEWSTATE'          = HF $init '__VIEWSTATE'
      '__VIEWSTATEGENERATOR' = HF $init '__VIEWSTATEGENERATOR'
      '__EVENTVALIDATION'    = HF $init '__EVENTVALIDATION'
      $cbG='on'; $cbS='on'; $cbP='on'
      'ctl00$ContentPlaceHolder1$btn_RunReport'='Run Report'
    }
    foreach ($p in $phis) { $form[$p.Name] = 'on' }
    Write-Host "  POST..."
    $null = Tci-Post -ReportId $rep.Id -Form $form -OutFile $cacheFile
    $html = Get-Content $cacheFile -Raw
  }

  # parse data table
  $tables = Tci-ParseTables -Html $html
  $dt = @(Tci-PickDataTable -Tables $tables -FirstTh $rep.Th)
  if ($dt.Count -eq 0) { Write-Host "  !! no data table (first-th=$($rep.Th)); skipping"; continue }
  $parsed = Tci-RowsFromTable -DataTableBodies $dt
  $hdr = @($parsed.Header); $all = $parsed.Rows
  Write-Host "  rows=$($all.Count-1) cols=$($hdr.Count)"

  $opIdx  = [Array]::IndexOf($hdr, 'OperationName')
  $metIdx = [Array]::IndexOf($hdr, 'MetricName')

  # relabel time columns
  if ($rep.Relabel -eq 'month') {
    for ($ci=0; $ci -lt $hdr.Count; $ci++) { if ($hdr[$ci] -match '^\d{6}$') { $hdr[$ci] = WwToFiscalMonth $hdr[$ci] } }
  } elseif ($rep.Relabel -eq 'quarter') {
    $sdIdx = [Array]::IndexOf($hdr, 'ATREVStartDate')
    if ($sdIdx -ge 0) {
      $uniq = @($all[1..($all.Count-1)] | ForEach-Object { ($_[$sdIdx]+'').Trim() } | Where-Object { $_ } | Select-Object -Unique)
      if ($uniq.Count -eq 1) {
        $base = $uniq[0]
        for ($ci=0; $ci -lt $hdr.Count; $ci++) {
          if ($hdr[$ci] -eq 'ATRevStart') { $hdr[$ci] = $base }
          elseif ($hdr[$ci] -match '^ATRev\+(\d+)$') { $hdr[$ci] = AddQ $base ([int]$matches[1]) }
        }
        Write-Host "  MOR quarters relabeled from base $base"
      } else { Write-Host "  MOR mixed ATREVStartDate ($($uniq.Count)); leaving ATRev+N labels" }
    }
  }

  # Class filter (union: op startswith TEST  OR  blank-op + monitor metric)
  $kept = New-Object System.Collections.Generic.List[object]
  $kept.Add($hdr)
  $bd = [ordered]@{}
  $unfiltered = $false
  if ($opIdx -lt 0 -or $metIdx -lt 0) {
    $unfiltered = $true
    for ($k=1;$k -lt $all.Count;$k++){ $kept.Add($all[$k]) }
    Write-Host "  (no Operation/Metric column -> kept all $($all.Count-1) rows, unfiltered)"
  } else {
    for ($k=1;$k -lt $all.Count;$k++){
      $row=$all[$k]; $op=($row[$opIdx]+'').Trim(); $met=($row[$metIdx]+'').Trim().ToUpper()
      $keepIt=$false; $label=$null
      if ($op.ToUpper().StartsWith('TEST')) { $keepIt=$true; $label=$op }
      elseif ([string]::IsNullOrWhiteSpace($op) -and ($monitors -contains $met)) { $keepIt=$true; $label="(blank) $met" }
      if ($keepIt) { $kept.Add($row); if(-not $bd.Contains($label)){$bd[$label]=0}; $bd[$label]++ }
    }
    Write-Host "  Class filter: total=$($kept.Count-1)"
  }

  if (($kept.Count-1) -le 0) { Write-Host "  !! 0 rows after filter; skipping tab"; continue }

  $sheet = ($rep.Label) -replace '[\\/\?\*\[\]:]',' '
  $csv = Join-Path $env:TEMP "$OutBase`_$($rep.Id).csv"
  Tci-WriteCsv -Rows $kept -Path $csv
  $results += [PSCustomObject]@{ Label=$rep.Label; Csv=$csv; Sheet=$sheet; Breakdown=$bd; Total=($kept.Count-1); Unfiltered=$unfiltered }
}

if ($results.Count -eq 0) { throw "No report produced data." }

# ===== Build multi-tab xlsx =====
$xlsx = Join-Path $env:TEMP "$OutBase.xlsx"
$excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
try {
  $target = $excel.Workbooks.Add()
  while ($target.Worksheets.Count -gt 1) { $target.Worksheets.Item($target.Worksheets.Count).Delete() }
  $target.Worksheets.Item(1).Name = '_placeholder_'
  foreach ($res in $results) {
    $tmp = $excel.Workbooks.Open($res.Csv)
    $ws = $tmp.Worksheets.Item(1)
    $ws.Name = $res.Sheet
    $lastCol = $ws.UsedRange.Columns.Count
    $h = $ws.Range($ws.Cells(1,1), $ws.Cells(1,$lastCol)); $h.Font.Bold=$true; $h.Interior.Color=14277081
    try { $ws.Activate(); $aw=$ws.Application.ActiveWindow; $aw.SplitColumn=4; $aw.SplitRow=1; $aw.FreezePanes=$true } catch {}
    $ws.Columns.AutoFit() | Out-Null
    $ws.Move([System.Reflection.Missing]::Value, $target.Worksheets.Item($target.Worksheets.Count))
    try { $tmp.Close($false) } catch {}
  }
  $target.Worksheets.Item('_placeholder_').Delete()
  if (Test-Path $xlsx) { Remove-Item $xlsx -Force }
  $target.SaveAs($xlsx, 51); try { $target.Close($false) } catch {}
} finally { try { $excel.Quit() } catch {}; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
Write-Host "`nxlsx: $xlsx ($([math]::Round((Get-Item $xlsx).Length/1KB,1)) KB)"

# ===== Email (polished card) =====
function Pill($t,$bg,$fg){ "<span style=`"background:$bg;color:$fg;font-weight:600;padding:2px 10px;border-radius:12px;`">$t</span>" }
$grandTotal = ($results | Measure-Object -Property Total -Sum).Sum

$bdSection = ''
foreach ($res in $results) {
  $bdSection += "<tr><td colspan=`"2`" style=`"padding:6px 16px;background:#1565a6;color:#fff;font-weight:600;font-size:11pt;`">$($res.Label) - $($res.Total) rows$(if($res.Unfiltered){' (no Operation column; all rows)'})</td></tr>"
  $z=0
  if ($res.Unfiltered) {
    $bdSection += "<tr><td style=`"padding:5px 16px;border-bottom:1px solid #f0f0f0;color:#80868b;font-size:11pt;`">(unfiltered)</td><td style=`"padding:5px 16px;border-bottom:1px solid #f0f0f0;text-align:right;color:#80868b;font-size:11pt;`">$($res.Total)</td></tr>"
  } else {
    foreach ($k in $res.Breakdown.Keys) {
      $bg = if ($z % 2 -eq 1) { 'background:#fafbfc;' } else { '' }
      $dim = if ($k -like '(blank)*') { 'color:#80868b;' } else { 'color:#3c4043;' }
      $bdSection += "<tr><td style=`"padding:5px 16px;border-bottom:1px solid #f0f0f0;font-size:11pt;$bg$dim`">$k</td><td style=`"padding:5px 16px;border-bottom:1px solid #f0f0f0;font-size:11pt;$bg text-align:right;$dim`">$($res.Breakdown[$k])</td></tr>"
      $z++
    }
  }
}

$body = @"
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#202124;padding:4px;">
  <div style="background:linear-gradient(90deg,#0a3d62,#1565a6);color:#fff;padding:12px 18px;border-radius:8px 8px 0 0;">
    <span style="font-size:16pt;font-weight:600;letter-spacing:.3px;">PHI Report Summary</span>
  </div>
  <table style="border-collapse:collapse;width:560px;box-shadow:0 1px 4px rgba(0,0,0,.12);border-radius:0 0 8px 8px;overflow:hidden;font-size:11pt;">
    <tr><td colspan="2" style="padding:8px 16px;background:#0a3d62;color:#fff;font-weight:600;letter-spacing:.3px;font-size:11pt;">Summary Details</td></tr>
    <tr><td style="padding:7px 16px;border-bottom:1px solid #ececec;background:#fafbfc;font-weight:600;color:#5f6368;width:200px;">Product</td><td style="padding:7px 16px;border-bottom:1px solid #ececec;"><span style="font-weight:700;color:#b8860b;">$ShortName</span> <span style="color:#80868b;">(Panther Lake)</span></td></tr>
    <tr><td style="padding:7px 16px;border-bottom:1px solid #ececec;background:#fafbfc;font-weight:600;color:#5f6368;">AT Group / SubGroup</td><td style="padding:7px 16px;border-bottom:1px solid #ececec;">$Group&nbsp;/&nbsp;$SubGroup</td></tr>
    <tr><td style="padding:7px 16px;border-bottom:1px solid #ececec;background:#fafbfc;font-weight:600;color:#5f6368;">Reports</td><td style="padding:7px 16px;border-bottom:1px solid #ececec;">$(Pill 'SDA Weekly' '#e6f4ea' '#0a6e31') $(Pill 'SDA Monthly' '#e6f4ea' '#0a6e31') $(Pill 'MOR Spread' '#e6f4ea' '#0a6e31') $(Pill 'ENG' '#e6f4ea' '#0a6e31')</td></tr>
    <tr><td style="padding:7px 16px;border-bottom:1px solid #ececec;background:#fafbfc;font-weight:600;color:#5f6368;">Filter</td><td style="padding:7px 16px;border-bottom:1px solid #ececec;">$(Pill 'Test (Class)' '#e6f4ea' '#0a6e31')</td></tr>
    <tr><td style="padding:7px 16px;border-bottom:1px solid #ececec;background:#fafbfc;font-weight:600;color:#5f6368;">Composition</td><td style="padding:7px 16px;border-bottom:1px solid #ececec;color:#3c4043;">Union of <b>TEST_*</b> ops + blank-op monitor rollups <span style="color:#80868b;">(MPS / EQA / CS)</span>, one tab per report</td></tr>
    <tr><td colspan="2" style="padding:8px 16px;background:#0a3d62;color:#fff;font-weight:600;letter-spacing:.3px;font-size:11pt;">Class Breakdown by Report</td></tr>
    $bdSection
    <tr><td style="padding:8px 16px;background:#e8f0fe;font-weight:700;color:#0a3d62;">Grand Total</td><td style="padding:8px 16px;background:#e8f0fe;text-align:right;font-weight:700;color:#0a3d62;font-size:12pt;">$grandTotal</td></tr>
  </table>
  <div style="margin-top:10px;padding:8px 18px;font-size:11px;color:#9aa0a6;font-style:italic;">This is generated by MPE forge Skill 'PHI Manage'</div>
  <div style="padding:0 18px 10px;font-size:11px;color:#5f6368;">Want to run your own PHI reports? Learn how to use this Forge skill here: <a href="https://goto.intel.com/mpeforge" style="color:#1565a6;text-decoration:underline;">MPE Forge - PHI Manage Guide</a></div>
</div>
"@

Tci-SendMail -Subject "PHI of $ShortName - All Reports Class" -HtmlBody $body -Attachments @($xlsx)
Write-Host "Emailed."
