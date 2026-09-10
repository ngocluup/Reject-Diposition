#requires -Version 5.1
<#
TCI Batch Driver
Reads TCI_Batch_Input.xlsx, runs each row, emails each result.

Schema (sheet 'Batch'):
  Product  - exact CommonName (e.g. 'ARL Refresh S 8C+16A+GT1')
  Report   - SDA | SDA Weekly | SDA Monthly | MOR | MOR Spread | MOR Forecasts | ENG
             BLANK = run all 4 (SDA Weekly + SDA Monthly + MOR + ENG) into one 4-tab xlsx
  Filter   - blank | yield | ppv | class | + <OperationName prefix>
  Email    - optional override; defaults to current user
#>
param(
  [string]$InputXlsx = "$env:USERPROFILE\Downloads\PHI Tracking\TCI_Batch_Input.xlsx",
  [string]$Lookup = "$env:USERPROFILE\Downloads\PHI Tracking\tci_commonname_lookup.json",
  [string]$OutDir = "$env:USERPROFILE\Downloads\PHI Tracking\BatchOutput",
  [string[]]$Product = @(),
  [string]$Report = '',
  [string]$Filter = '',
  [string]$Email = '',
  [int]$MaxCacheHours = 12,
  [switch]$NoParallel,
  [switch]$CombineXlsx,
  [switch]$NoEmail
)
$ErrorActionPreference='Stop'

. "$PSScriptRoot\tci_lib.ps1"

$script:MaxCacheHours = $MaxCacheHours
$script:NoParallel = [bool]$NoParallel
$script:FreshThisRun = @{}
$script:InitCache = @{}   # in-memory init HTML cache (improvement #4)
if(-not (Test-Path $OutDir)){ New-Item -ItemType Directory -Path $OutDir | Out-Null }
# --- Load lookup ---
$lk = Get-Content $Lookup -Raw | ConvertFrom-Json
$lkMap = @{}
foreach($e in $lk.Lookup){ $lkMap[$e.CommonName] = @{ Grp=$e.AT_Group; Sub=$e.AT_SubGroup } }

# --- Report URL map ---
$reportMap = @{
  'SDA'          = @{ Url='https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=PORSDAForecast';     InitTag='sda';    FirstTh='CommonName'; RelabelWw=$false; SheetTag='SDA Weekly' }
  'SDA WEEKLY'   = @{ Url='https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=PORSDAForecast';     InitTag='sda';    FirstTh='CommonName'; RelabelWw=$false; SheetTag='SDA Weekly' }
  'SDA MONTHLY'  = @{ Url='https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=PORMonthlyForecast'; InitTag='sdamo';  FirstTh='CommonName'; RelabelWw=$true;  SheetTag='SDA Monthly' }
  'MOR'          = @{ Url='https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=MORSpread';          InitTag='morsp';  FirstTh='MetricName'; RelabelWw=$false; SheetTag='MOR' }
  'MOR SPREAD'   = @{ Url='https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=MORSpread';          InitTag='morsp';  FirstTh='MetricName'; RelabelWw=$false; SheetTag='MOR' }
  'MOR FORECASTS'= @{ Url='https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=MORForecasts';       InitTag='mor';    FirstTh='MetricName'; RelabelWw=$false; SheetTag='MOR Fcst' }
  'ENG'          = @{ Url='https://tcitools.intel.com/Web/Test/Reports/TestReport.aspx?R=ENGForecasts';       InitTag='eng';    FirstTh='ATGroup';    RelabelWw=$false; SheetTag='ENG' }
}
# blank Report -> run these 4
$blankReportSet = @('SDA Weekly','SDA Monthly','MOR','ENG')

# --- Helpers ---
function Hidden($h,$n){ ([regex]::Match($h,'name="'+[regex]::Escape($n)+'"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value }
function Get-PanelCtl($html,$panelName,$labelExact){
  $pm=[regex]::Match($html,'(?is)id="ContentPlaceHolder1_Filters_'+[regex]::Escape($panelName)+'".*?</div>\s*</div>')
  if(-not $pm.Success){ throw "panel $panelName not found" }
  $items=[regex]::Matches($pm.Value,'(?is)<input[^>]*name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>\s*([^<]+?)\s*</label>')
  foreach($it in $items){ if($it.Groups[2].Value.Trim() -eq $labelExact){ return $it.Groups[1].Value } }
  throw "label '$labelExact' not found in $panelName"
}
function WwToMonth($yyyyww){
  if([string]::IsNullOrWhiteSpace($yyyyww) -or $yyyyww.Length -ne 6){ return $yyyyww }
  $y=[int]$yyyyww.Substring(0,4); $w=[int]$yyyyww.Substring(4,2)
  $j=[datetime]::new($y,1,1)
  $sun=$j.AddDays(-[int]$j.DayOfWeek).AddDays(7*($w-1))
  $sun.ToString("MMM yyyy")
}
function Safe-Name($s){ ($s -replace '[\\/:*?"<>|+]','_') -replace '\s+','_' }
function Expand-ProductList($products){
  $out = New-Object System.Collections.Generic.List[string]
  foreach($p in @($products)){
    foreach($part in (($p+'') -split '\s*[;,]\s*')){
      $v = $part.Trim()
      if(-not [string]::IsNullOrWhiteSpace($v)){ $out.Add($v) }
    }
  }
  return @($out | Select-Object -Unique)
}
# A cache file counts as usable only if it exists AND is newer than
# $MaxCacheHours (0 = always re-fetch). Files written *during this run*
# (e.g. by the parallel prefetch) are always treated as fresh so the
# sequential pass reuses them instead of re-POSTing.
function Cache-Fresh($path){
  if($script:FreshThisRun -and $script:FreshThisRun.ContainsKey($path)){ return $true }
  if($script:MaxCacheHours -le 0){ return $false }
  if(-not (Test-Path $path)){ return $false }
  (Get-Item $path).LastWriteTime -gt (Get-Date).AddHours(-$script:MaxCacheHours)
}
# Invoke-WebRequest with a bounded timeout + one auto-retry on timeout. The
# default IWR timeout is infinite, so a stalled TCI render hangs forever
# (lesson 15). On a Timeout WebException, retry once before throwing.
function Invoke-WebRetry($Uri,$Method='Get',$Body,$TimeoutSec=300){
  $attempt=0
  while($true){
    try{
      $p=@{ Uri=$Uri; Method=$Method; UseDefaultCredentials=$true; UseBasicParsing=$true; TimeoutSec=$TimeoutSec; ErrorAction='Stop' }
      if($Method -eq 'POST' -and $Body){ $p.Body=$Body }
      return Invoke-WebRequest @p
    } catch {
      $isTimeout = $_.Exception -is [System.Net.WebException] -and $_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout
      if($attempt -lt 1 -and $isTimeout){ $attempt++; Write-Host "  [http] timeout after ${TimeoutSec}s - retry $attempt/1 ..." -ForegroundColor Yellow; continue }
      throw
    }
  }
}

# --- Parallel prefetch: warm the per-report POST caches concurrently so an
# all-reports bundle is bounded by the slowest single POST, not their sum
# (improvement #1). Only the network POST + cache write runs in runspaces;
# all HTML parsing / Excel / Outlook stays single-threaded on the main thread.
function Start-PrefetchPosts($product,$grp,$sub,$reportKeys){
  if($script:NoParallel){ return }
  $todo=@(); $seen=@{}
  foreach($rk in $reportKeys){
    $rep=$reportMap[$rk.ToUpper()]; if(-not $rep){ continue }
    $init ="$env:TEMP\tci_init_$($rep.InitTag).html"
    $after="$env:TEMP\tci_after_$(Safe-Name $product)_$($rep.InitTag).html"
    if(Cache-Fresh $after){ continue }
    if($seen.ContainsKey($rep.InitTag)){ continue }
    $seen[$rep.InitTag]=$true
    $todo += [pscustomobject]@{ Url=$rep.Url; Init=$init; After=$after; InitTag=$rep.InitTag }
  }
  if($todo.Count -lt 2){ return }   # runspace overhead not worth it for <2
  Write-Host ("  [parallel] warming {0} report POST cache(s)..." -f $todo.Count) -ForegroundColor Cyan
  $sb={
    param($Url,$Init,$After,$Product,$Grp,$Sub,$MaxCacheHours)
    $ErrorActionPreference='Stop'
    function Cache-Fresh($path,$maxH){ if($maxH -le 0){ return $false }; if(-not (Test-Path $path)){ return $false }; (Get-Item $path).LastWriteTime -gt (Get-Date).AddHours(-$maxH) }
    function Invoke-WebRetry($Uri,$Method='Get',$Body,$TimeoutSec=180){
      $attempt=0
      while($true){
        try{
          $p=@{ Uri=$Uri; Method=$Method; UseDefaultCredentials=$true; UseBasicParsing=$true; TimeoutSec=$TimeoutSec; ErrorAction='Stop' }
          if($Method -eq 'POST' -and $Body){ $p.Body=$Body }
          return Invoke-WebRequest @p
        } catch {
          $isTimeout = $_.Exception -is [System.Net.WebException] -and $_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout
          if($attempt -lt 1 -and $isTimeout){ $attempt++; continue }
          throw
        }
      }
    }
    function Hidden($h,$n){ ([regex]::Match($h,'name="'+[regex]::Escape($n)+'"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value }
    function Get-PanelCtl($html,$panelName,$labelExact){
      $pm=[regex]::Match($html,'(?is)id="ContentPlaceHolder1_Filters_'+[regex]::Escape($panelName)+'".*?</div>\s*</div>')
      if(-not $pm.Success){ throw "panel $panelName not found" }
      $items=[regex]::Matches($pm.Value,'(?is)<input[^>]*name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>\s*([^<]+?)\s*</label>')
      foreach($it in $items){ if($it.Groups[2].Value.Trim() -eq $labelExact){ return $it.Groups[1].Value } }
      throw "label '$labelExact' not found in $panelName"
    }
    if(-not (Cache-Fresh $Init $MaxCacheHours)){
      $r0=Invoke-WebRetry -Uri $Url -TimeoutSec 120
      [IO.File]::WriteAllText($Init,$r0.Content,[Text.UTF8Encoding]::new($false))
    }
    $h=Get-Content $Init -Raw
    $vs=Hidden $h '__VIEWSTATE'; $vsg=Hidden $h '__VIEWSTATEGENERATOR'; $ev=Hidden $h '__EVENTVALIDATION'
    $cGrp=Get-PanelCtl $h 'AT_Group' $Grp
    $cSub=Get-PanelCtl $h 'AT_SubGroup' $Sub
    $cCn =Get-PanelCtl $h 'CommonName' $Product
    $panelPhi=[regex]::Match($h,'(?is)id="ContentPlaceHolder1_Filters_PHIParameters".*?</div>\s*</div>').Value
    $phi=[regex]::Matches($panelPhi,'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }
    $f=[ordered]@{
      '__EVENTTARGET'=''; '__EVENTARGUMENT'=''
      '__VIEWSTATE'=$vs; '__VIEWSTATEGENERATOR'=$vsg; '__EVENTVALIDATION'=$ev
      $cGrp='on'; $cSub='on'; $cCn='on'
      'ctl00$ContentPlaceHolder1$btn_RunReport'='Run Report'
    }
    foreach($p in $phi){ $f[$p]='on' }
    $rr=Invoke-WebRetry -Uri $Url -Method POST -Body $f -TimeoutSec 240
    [IO.File]::WriteAllText($After,$rr.Content,[Text.UTF8Encoding]::new($false))
    "rows~$(([regex]::Matches($rr.Content,'(?is)<tr')).Count)"
  }
  $pool=[runspacefactory]::CreateRunspacePool(1,[Math]::Min(4,$todo.Count)); $pool.Open()
  $handles=@()
  foreach($t in $todo){
    $ps=[powershell]::Create(); $ps.RunspacePool=$pool
    [void]$ps.AddScript($sb).AddArgument($t.Url).AddArgument($t.Init).AddArgument($t.After).AddArgument($product).AddArgument($grp).AddArgument($sub).AddArgument($script:MaxCacheHours)
    $handles += [pscustomobject]@{ PS=$ps; Handle=$ps.BeginInvoke(); Tag=$t.InitTag; Init=$t.Init; After=$t.After }
  }
  foreach($hd in $handles){
    try{ $r=$hd.PS.EndInvoke($hd.Handle); Write-Host "    [parallel] $($hd.Tag): $r" -ForegroundColor DarkGray; $script:FreshThisRun[$hd.Init]=$true; $script:FreshThisRun[$hd.After]=$true }
    catch{ Write-Warning "    [parallel] $($hd.Tag) failed: $($_.Exception.Message) (sequential pass will retry)" }
    finally{ $hd.PS.Dispose() }
  }
  $pool.Close(); $pool.Dispose()
}

# --- Lookup self-heal: when a CommonName is missing from the lookup JSON,
# infer Group/SubGroup from the CommonName suffix (SKILL.md §0), verify the
# labels against the live SDA init panels when possible, persist the new
# entry to the lookup file, and continue instead of failing the run (#3).
function Resolve-LookupEntry($product){
  $u = $product.ToUpper()
  $grp='Client'; $sub='Desktop'; $how='suffix-heuristic'
  $cls=$null
  foreach($t in ($u -split '\s+')){ if($t -match '^(HX|S|H|U|P|M|Y)$'){ $cls=$t; break } }
  if($u -match 'XCC|MCC|\bAP\b|\bSP\b|XEON|GNR|SRF|CWF|DMR|CLEARWATER|GRANITE'){ $grp='Server'; $sub='Server' }
  elseif($u -match 'BMG|\bARC\b|\bGPU\b|BATTLEMAGE|CELESTIAL|DRUID'){ $grp='GPU'; $sub='Client GPU' }
  elseif($cls -in @('U','P','M','Y')){ $grp='Client'; $sub='Mobile' }
  elseif($cls -in @('S','H','HX')){ $grp='Client'; $sub='Desktop' }

  try {
    $rep=$reportMap['SDA']
    $init="$env:TEMP\tci_init_$($rep.InitTag).html"
    if(-not (Cache-Fresh $init)){
      $r0=Invoke-WebRetry -Uri $rep.Url -TimeoutSec 120
      [IO.File]::WriteAllText($init,$r0.Content,[Text.UTF8Encoding]::new($false))
    }
    $h=Get-Content $init -Raw
    [void](Get-PanelCtl $h 'CommonName' $product)   # throws if the product label is absent
    foreach($cand in @($sub,'Desktop','Mobile','Server','Client GPU')){
      try{ [void](Get-PanelCtl $h 'AT_SubGroup' $cand); $sub=$cand; $how='verified-on-init'; break }catch{}
    }
  } catch {
    Write-Warning "    self-heal: '$product' not confirmed on SDA init panel - using heuristic only"
  }

  try {
    $lkRaw = Get-Content $Lookup -Raw | ConvertFrom-Json
    $exists=$false; foreach($e in $lkRaw.Lookup){ if($e.CommonName -eq $product){ $exists=$true; break } }
    if(-not $exists){
      $lkRaw.Lookup = @($lkRaw.Lookup) + [pscustomobject]@{ CommonName=$product; AT_Group=$grp; AT_SubGroup=$sub }
      ($lkRaw | ConvertTo-Json -Depth 6) | Set-Content -Path $Lookup -Encoding UTF8
    }
  } catch { Write-Warning "    self-heal: failed to persist lookup: $($_.Exception.Message)" }

  return @{ Grp=$grp; Sub=$sub; How=$how }
}

# --- Per-report runner: returns @{ Hdr=...; Rows=...; SheetTag=... } ---
function Invoke-TciReport($product,$grp,$sub,$reportKey,$filter){
  $rep = $reportMap[$reportKey.ToUpper()]
  if(-not $rep){ throw "Unknown Report '$reportKey'" }

  $init  = "$env:TEMP\tci_init_$($rep.InitTag).html"
  $tag   = Safe-Name $product
  $after = "$env:TEMP\tci_after_$($tag)_$($rep.InitTag).html"

  # In-memory init cache: avoid re-reading same 900KB file per product (improvement #4)
  if($script:InitCache.ContainsKey($init)){
    $h = $script:InitCache[$init]
  } else {
    if(-not (Cache-Fresh $init)){
      $r0 = Invoke-WebRetry -Uri $rep.Url -TimeoutSec 120
      [IO.File]::WriteAllText($init,$r0.Content,[Text.UTF8Encoding]::new($false))
      $script:FreshThisRun[$init]=$true
    }
    $h = Get-Content $init -Raw
    $script:InitCache[$init] = $h
  }
  $vs=Hidden $h '__VIEWSTATE'; $vsg=Hidden $h '__VIEWSTATEGENERATOR'; $ev=Hidden $h '__EVENTVALIDATION'
  $cGrp = Get-PanelCtl $h 'AT_Group' $grp
  $cSub = Get-PanelCtl $h 'AT_SubGroup' $sub
  $cCn  = Get-PanelCtl $h 'CommonName' $product
  $panelPhi=[regex]::Match($h,'(?is)id="ContentPlaceHolder1_Filters_PHIParameters".*?</div>\s*</div>').Value
  $phi=[regex]::Matches($panelPhi,'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }

  if(-not (Cache-Fresh $after)){
    $f=[ordered]@{
      '__EVENTTARGET'=''; '__EVENTARGUMENT'=''
      '__VIEWSTATE'=$vs; '__VIEWSTATEGENERATOR'=$vsg; '__EVENTVALIDATION'=$ev
      $cGrp='on'; $cSub='on'; $cCn='on'
      'ctl00$ContentPlaceHolder1$btn_RunReport'='Run Report'
    }
    foreach($p in $phi){ $f[$p]='on' }
    Write-Host "  POST $($rep.Url) Grp=$grp Sub=$sub PHI=$($phi.Count)"
    $postTimeout = if($rep.InitTag -eq 'sda'){ 600 } else { 480 }
    $rr=Invoke-WebRetry -Uri $rep.Url -Method POST -Body $f -TimeoutSec $postTimeout
    [IO.File]::WriteAllText($after,$rr.Content,[Text.UTF8Encoding]::new($false))
    $script:FreshThisRun[$after]=$true
  } else {
    Write-Host "  reuse cache $after"
  }

  $c = Get-Content $after -Raw
  if($c -match 'delivered no data'){ throw 'no data' }

  # Narrow parse: extract only tbl_Main region if present (improvement #3)
  $mainStart = $c.IndexOf('id="tbl_Main"')
  if($mainStart -ge 0){ $c = $c.Substring($mainStart) }

  $tables=New-Object System.Collections.Generic.List[string]; $idx=0
  while($idx -lt $c.Length){
    $om=[regex]::Match($c.Substring($idx),'<table\b[^>]*>'); if(-not $om.Success){ break }
    $sI=$idx+$om.Index+$om.Length; $depth=1; $pos=$sI
    while($depth -gt 0 -and $pos -lt $c.Length){
      $o=$c.IndexOf('<table',$pos); $cl=$c.IndexOf('</table>',$pos)
      if($cl -lt 0){ $depth=0; break }
      if($o -ge 0 -and $o -lt $cl){ $depth++; $pos=$o+6 } else { $depth--; $pos=$cl+8 }
    }
    $tables.Add($c.Substring($sI,($pos-8)-$sI)); $idx=$sI
  }
  $firstTh = $rep.FirstTh
  $dt = $tables | Where-Object {
    $n=[regex]::Matches($_,'(?is)<th[^>]*>\s*([^<]{1,60})') | ForEach-Object { $_.Groups[1].Value.Trim() }
    ($n.Count -gt 0) -and ($n[0] -eq $firstTh)
  }
  $all=New-Object System.Collections.Generic.List[object]; $hdr=$null
  foreach($body in $dt){
    $rs=[regex]::Matches($body,'(?is)<tr[^>]*>(.*?)</tr>')
    $tr=New-Object System.Collections.Generic.List[object]
    foreach($rrr in $rs){
      $cm=[regex]::Matches($rrr.Groups[1].Value,'(?is)<(t[hd])\b[^>]*>(.*?)</\1>')
      if($cm.Count -eq 0){ continue }
      $cells=foreach($mm in $cm){ $t=[regex]::Replace($mm.Groups[2].Value,'(?is)<[^>]+>',' '); $t=[System.Net.WebUtility]::HtmlDecode($t).Trim() -replace '\s+',' '; ,$t }
      $tr.Add(@($cells))
    }
    if(-not $hdr){ $hdr=$tr[0]; $all.Add($hdr) }
    for($i=1;$i -lt $tr.Count;$i++){ $all.Add($tr[$i]) }
  }
  Write-Host "  total rows: $($all.Count - 1) cols: $($hdr.Count)"

  $iMet = [array]::IndexOf($hdr,'MetricName')
  $iOp  = [array]::IndexOf($hdr,'OperationName')

  $filtered = New-Object System.Collections.Generic.List[object]
  $filtered.Add($hdr)
  $fLc = ($filter+'').ToLower().Trim()
  if([string]::IsNullOrWhiteSpace($fLc)){
    for($k=1;$k -lt $all.Count;$k++){ $filtered.Add($all[$k]) }
  } elseif($fLc -eq 'yield'){
    $yieldMetrics = @('U/D','R/D','FINISH YIELD')
    for($k=1;$k -lt $all.Count;$k++){
      $row=$all[$k]
      $op = if($iOp -ge 0 -and $iOp -lt $row.Count){ ($row[$iOp]+'').Trim() } else { '' }
      $met= if($iMet -ge 0 -and $iMet -lt $row.Count){ ($row[$iMet]+'').Trim().ToUpper() } else { '' }
      if([string]::IsNullOrWhiteSpace($op) -and ($yieldMetrics -contains $met)){ $filtered.Add($row) }
    }
  } elseif($fLc -eq 'ppv'){
    for($k=1;$k -lt $all.Count;$k++){
      $row=$all[$k]
      $op = if($iOp -ge 0 -and $iOp -lt $row.Count){ ($row[$iOp]+'').Trim() } else { '' }
      $met= if($iMet -ge 0 -and $iMet -lt $row.Count){ ($row[$iMet]+'').Trim().ToUpper() } else { '' }
      if($op.ToUpper().StartsWith('PPV')){ $filtered.Add($row) }
      elseif([string]::IsNullOrWhiteSpace($op) -and $met -eq 'PPV-M SAMPLE SIZE'){ $filtered.Add($row) }
    }
  } elseif($fLc -eq 'class'){
    for($k=1;$k -lt $all.Count;$k++){
      $row=$all[$k]
      $op = if($iOp -ge 0 -and $iOp -lt $row.Count){ ($row[$iOp]+'').Trim() } else { '' }
      $met= if($iMet -ge 0 -and $iMet -lt $row.Count){ ($row[$iMet]+'').Trim().ToUpper() } else { '' }
      if($op.ToUpper().StartsWith('TEST')){ $filtered.Add($row) }
      elseif([string]::IsNullOrWhiteSpace($op) -and ($met -match 'MPS|EQA|CS MONITOR')){ $filtered.Add($row) }
    }
  } elseif($fLc.StartsWith('+')){
    $prefix = $fLc.TrimStart('+').Trim().ToUpper()
    for($k=1;$k -lt $all.Count;$k++){
      $row=$all[$k]
      $op = if($iOp -ge 0 -and $iOp -lt $row.Count){ ($row[$iOp]+'').Trim().ToUpper() } else { '' }
      if($op.StartsWith($prefix)){ $filtered.Add($row) }
    }
  } else {
    throw "Unknown Filter '$filter'"
  }
  Write-Host "  filtered rows: $($filtered.Count - 1)"
  if($filtered.Count -le 1){ throw 'no rows after filter' }

  if($rep.RelabelWw){
    $fixedCols = @('CommonName','MetricName','ItemID','ItemSegment','ENGType','ATREVStartDate','ATGroup','ATSubGroup','OperationName','ResourceName','STRGC','DLCP','FunctionalCore','GraphicsCore','PackageSize','PackageTech','Rev','Step','TestFlow','BinConfigName')
    for($ci=0; $ci -lt $hdr.Count; $ci++){
      if($hdr[$ci] -match '^\d{6}$' -and $hdr[$ci] -notin $fixedCols){ $hdr[$ci] = WwToMonth $hdr[$ci] }
    }
    $filtered[0] = $hdr
  }

  return @{ Hdr=$hdr; Rows=$filtered; SheetTag=$rep.SheetTag }
}

# --- Write multi-tab xlsx (reuses shared Excel COM instance for efficiency) ---
function Write-MultiTabXlsx($sections, $xlsxPath, $excelApp){
  $ownExcel = $false
  if(-not $excelApp){ $excelApp=New-Object -ComObject Excel.Application; $excelApp.Visible=$false; $excelApp.DisplayAlerts=$false; $ownExcel=$true }
  $tmpCsvs = @()
  foreach($sec in $sections){
    $csvP = [IO.Path]::Combine($env:TEMP, "tci_tab_$([Guid]::NewGuid().ToString('N')).csv")
    $maxCols=($sec.Rows | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    $sb=New-Object System.Text.StringBuilder
    foreach($row in $sec.Rows){
      $p=@($row)+(,'')*($maxCols-$row.Count)
      $e=$p | ForEach-Object { $v="$_"; if($v -match '[",\r\n]'){ '"'+($v -replace '"','""')+'"' } else { $v } }
      [void]$sb.AppendLine(($e -join ','))
    }
    [IO.File]::WriteAllText($csvP,$sb.ToString(),[Text.UTF8Encoding]::new($true))
    $tmpCsvs += [pscustomobject]@{ Csv=$csvP; SheetTag=$sec.SheetTag }
  }

  try {
    $wbT = $excelApp.Workbooks.Add()
    while($wbT.Worksheets.Count -gt 1){ $wbT.Worksheets.Item($wbT.Worksheets.Count).Delete() }
    $wbT.Worksheets.Item(1).Name = '_placeholder_'

    foreach($tc in $tmpCsvs){
      $wbC = $excelApp.Workbooks.Open($tc.Csv)
      $wsC = $wbC.Worksheets.Item(1)
      $afterSheet = $wbT.Worksheets.Item($wbT.Worksheets.Count)
      $wsC.Move([System.Reflection.Missing]::Value, $afterSheet)
      $moved = $wbT.Worksheets.Item($wbT.Worksheets.Count)
      $sn = $tc.SheetTag
      if($sn.Length -gt 31){ $sn = $sn.Substring(0,31) }
      # existing sheet names in target wb
      $existing = @(); for($si=1; $si -le $wbT.Worksheets.Count; $si++){ if($si -ne $moved.Index){ $existing += $wbT.Worksheets.Item($si).Name } }
      $base = $sn; $k = 1
      while($existing -contains $sn){
        $suffix = "_$k"; $maxLen = 31 - $suffix.Length
        $trim = if($base.Length -gt $maxLen){ $base.Substring(0,$maxLen) } else { $base }
        $sn = $trim + $suffix; $k++
      }
      Write-Host ("  rename moved sheet '{0}' -> '{1}'" -f $moved.Name, $sn)
      $moved.Name = $sn
      $lastCol = $moved.UsedRange.Columns.Count
      $hr = $moved.Range($moved.Cells(1,1),$moved.Cells(1,$lastCol)); $hr.Font.Bold=$true; $hr.Interior.Color=0xD9D9D9
      $moved.Activate() | Out-Null
      $excelApp.ActiveWindow.SplitColumn=4; $excelApp.ActiveWindow.SplitRow=1; $excelApp.ActiveWindow.FreezePanes=$true
      $moved.Columns.AutoFit() | Out-Null
      try { $wbC.Close($false) } catch {}
    }

    $ph = $wbT.Worksheets.Item('_placeholder_')
    $ph.Delete()
    if(Test-Path $xlsxPath){ Remove-Item $xlsxPath -Force }
    $wbT.SaveAs($xlsxPath, 51); $wbT.Close($false)
  } finally {
    if($ownExcel){ try { $excelApp.Quit() } catch {}; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
    foreach($tc in $tmpCsvs){ try { Remove-Item $tc.Csv -Force } catch {} }
  }
}

# --- Read input ---
$jobs = @()
$directProducts = Expand-ProductList $Product
if($directProducts.Count -gt 0){
  $rowNum = 1
  foreach($directProduct in $directProducts){
    $jobs += [pscustomobject]@{ Row=$rowNum; Product=$directProduct; Report=$Report.Trim(); Filter=$Filter.Trim(); Email=$Email.Trim() }
    $rowNum++
  }
  "Loaded $($jobs.Count) direct job(s) from parameters"
} else {
  $excel=New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
  try {
    $wb = $excel.Workbooks.Open($InputXlsx)
    $ws = $wb.Worksheets.Item(1)
    $rowCount = $ws.UsedRange.Rows.Count
    for($r=2;$r -le $rowCount;$r++){
      $prodCell = ($ws.Cells.Item($r,1).Value2 + '').Trim()
      $reportCell  = ($ws.Cells.Item($r,2).Value2 + '').Trim()
      $filterCell  = ($ws.Cells.Item($r,3).Value2 + '').Trim()
      $emailCell   = ($ws.Cells.Item($r,4).Value2 + '').Trim()
      if([string]::IsNullOrWhiteSpace($prodCell)){ continue }
      $jobs += [pscustomobject]@{ Row=$r; Product=$prodCell; Report=$reportCell; Filter=$filterCell; Email=$emailCell }
    }
    $wb.Close($false)
  } finally { try{ $excel.Quit() } catch {}; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
  "Loaded $($jobs.Count) job(s) from $InputXlsx"
}
$jobs | Format-Table Row,Product,Report,Filter,Email -AutoSize | Out-String | Write-Host

# --- Resolve lookups for ALL jobs upfront (needed for bulk prefetch) ---
$failedLookups = @()
foreach($job in $jobs){
  if(-not $lkMap.ContainsKey($job.Product)){
    $heal = Resolve-LookupEntry $job.Product
    if($heal){
      $lkMap[$job.Product] = @{ Grp=$heal.Grp; Sub=$heal.Sub }
      Write-Warning "  lookup self-heal: '$($job.Product)' -> $($heal.Grp)/$($heal.Sub) ($($heal.How)); appended to lookup"
    } else {
      $failedLookups += $job.Product
    }
  }
}

# --- Bulk prefetch: all (product x report) POSTs in parallel ---
# Optimizations (2026-05-31):
#   1. Sliding window (WaitAny) instead of fixed waves — keeps all 3 slots busy continuously
#   2. Pre-fetch 4 unique init pages once before runspaces (eliminates redundant GETs)
#   3. SDA timeout increased to 600s (was 480s — last-wave timeouts common)
#   4. Skip ENG for known-missing products (Extended Temp, lookup EngAvailable=false)
function Start-BulkPrefetch($jobs){
  if($script:NoParallel){ return }
  $todo=@()
  # Products known to be missing from ENG panel (optimization #4)
  $engSkipProducts = @($lk.Lookup | Where-Object { $_.PSObject.Properties['EngAvailable'] -and $_.EngAvailable -eq $false } | ForEach-Object { $_.CommonName })
  # Also skip any product whose name contains "Extended Temp" (always missing from ENG)
  foreach($job in $jobs){
    if($job.Product -in $failedLookups){ continue }
    $grp=$lkMap[$job.Product].Grp; $sub=$lkMap[$job.Product].Sub
    $reportsToRun = if([string]::IsNullOrWhiteSpace($job.Report)){ $blankReportSet } else { @($job.Report) }
    foreach($rk in $reportsToRun){
      $rep=$reportMap[$rk.ToUpper()]; if(-not $rep){ continue }
      # Skip ENG for known-missing products (#4)
      if($rep.InitTag -eq 'eng' -and ($job.Product -match 'Extended Temp' -or $job.Product -in $engSkipProducts)){
        Write-Host "    [skip] $($job.Product)|eng (known missing from ENG panel)" -ForegroundColor DarkYellow
        continue
      }
      $init ="$env:TEMP\tci_init_$($rep.InitTag).html"
      $after="$env:TEMP\tci_after_$(Safe-Name $job.Product)_$($rep.InitTag).html"
      if(Cache-Fresh $after){ continue }
      $key="$($job.Product)|$($rep.InitTag)"
      if($todo | Where-Object { "$($_.Product)|$($_.InitTag)" -eq $key }){ continue }
      # Weight: eng=1, morsp/mor=2, sdamo=3, sda=4 (heavier = more server load)
      $weight = switch($rep.InitTag){ 'eng'{1} 'morsp'{2} 'mor'{2} 'sdamo'{3} 'sda'{4} default{3} }
      # Timeout: SDA gets 600s (#3), others get 480s
      $postTimeout = if($rep.InitTag -eq 'sda'){ 600 } else { 480 }
      $todo += [pscustomobject]@{ Url=$rep.Url; Init=$init; After=$after; InitTag=$rep.InitTag; Product=$job.Product; Grp=$grp; Sub=$sub; Weight=$weight; PostTimeout=$postTimeout }
    }
  }
  if($todo.Count -lt 2){ return }
  # Sort by weight (lightest first) so fast reports cache quickly
  $todo = $todo | Sort-Object Weight
  $maxConcurrent = [Math]::Min(3, $todo.Count)   # TCI safe concurrency limit

  # --- Pre-fetch unique init pages (#2) — eliminates N redundant GETs ---
  $uniqueInits = @{}
  foreach($t in $todo){ if(-not $uniqueInits.ContainsKey($t.Init)){ $uniqueInits[$t.Init]=$t.Url } }
  foreach($initPath in $uniqueInits.Keys){
    if(-not (Cache-Fresh $initPath)){
      Write-Host "    [init-prefetch] GET $initPath" -ForegroundColor DarkGray
      $r0=Invoke-WebRetry -Uri $uniqueInits[$initPath] -TimeoutSec 120
      [IO.File]::WriteAllText($initPath,$r0.Content,[Text.UTF8Encoding]::new($false))
      $script:FreshThisRun[$initPath]=$true
    }
  }

  Write-Host ("`n  [bulk-prefetch] warming {0} POST cache(s) across {1} product(s) (sliding window, max {2} concurrent)..." -f $todo.Count, ($todo | ForEach-Object { $_.Product } | Select-Object -Unique).Count, $maxConcurrent) -ForegroundColor Cyan
  $sb={
    param($Url,$Init,$After,$Product,$Grp,$Sub,$MaxCacheHours,$PostTimeout)
    $ErrorActionPreference='Stop'
    function Invoke-WebRetry($Uri,$Method='Get',$Body,$TimeoutSec=300){
      $attempt=0
      while($true){
        try{
          $p=@{ Uri=$Uri; Method=$Method; UseDefaultCredentials=$true; UseBasicParsing=$true; TimeoutSec=$TimeoutSec; ErrorAction='Stop' }
          if($Method -eq 'POST' -and $Body){ $p.Body=$Body }
          return Invoke-WebRequest @p
        } catch {
          $isTimeout = $_.Exception -is [System.Net.WebException] -and $_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout
          if($attempt -lt 1 -and $isTimeout){ $attempt++; continue }
          throw
        }
      }
    }
    function Hidden($h,$n){ ([regex]::Match($h,'name="'+[regex]::Escape($n)+'"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value }
    function Get-PanelCtl($html,$panelName,$labelExact){
      $pm=[regex]::Match($html,'(?is)id="ContentPlaceHolder1_Filters_'+[regex]::Escape($panelName)+'".*?</div>\s*</div>')
      if(-not $pm.Success){ throw "panel $panelName not found" }
      $items=[regex]::Matches($pm.Value,'(?is)<input[^>]*name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>\s*([^<]+?)\s*</label>')
      foreach($it in $items){ if($it.Groups[2].Value.Trim() -eq $labelExact){ return $it.Groups[1].Value } }
      throw "label '$labelExact' not found in $panelName"
    }
    # Init is already pre-fetched by the main thread (#2) — just read it
    $h=Get-Content $Init -Raw
    $vs=Hidden $h '__VIEWSTATE'; $vsg=Hidden $h '__VIEWSTATEGENERATOR'; $ev=Hidden $h '__EVENTVALIDATION'
    $cGrp=Get-PanelCtl $h 'AT_Group' $Grp
    $cSub=Get-PanelCtl $h 'AT_SubGroup' $Sub
    $cCn =Get-PanelCtl $h 'CommonName' $Product
    $panelPhi=[regex]::Match($h,'(?is)id="ContentPlaceHolder1_Filters_PHIParameters".*?</div>\s*</div>').Value
    $phi=[regex]::Matches($panelPhi,'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }
    $f=[ordered]@{
      '__EVENTTARGET'=''; '__EVENTARGUMENT'=''
      '__VIEWSTATE'=$vs; '__VIEWSTATEGENERATOR'=$vsg; '__EVENTVALIDATION'=$ev
      $cGrp='on'; $cSub='on'; $cCn='on'
      'ctl00$ContentPlaceHolder1$btn_RunReport'='Run Report'
    }
    foreach($p in $phi){ $f[$p]='on' }
    $rr=Invoke-WebRetry -Uri $Url -Method POST -Body $f -TimeoutSec $PostTimeout
    [IO.File]::WriteAllText($After,$rr.Content,[Text.UTF8Encoding]::new($false))
    "rows~$(([regex]::Matches($rr.Content,'(?is)<tr')).Count)"
  }
  # --- Sliding window execution (#1): maintain $maxConcurrent active slots ---
  # As each finishes, immediately start the next — no idle time between waves.
  $sw=[Diagnostics.Stopwatch]::StartNew()
  $pool=[runspacefactory]::CreateRunspacePool(1,$maxConcurrent); $pool.Open()
  $active = New-Object System.Collections.Generic.List[pscustomobject]
  $queueIdx = 0; $completed = 0; $failed = 0

  # Fill initial slots
  while($active.Count -lt $maxConcurrent -and $queueIdx -lt $todo.Count){
    $t=$todo[$queueIdx]; $queueIdx++
    $ps=[powershell]::Create(); $ps.RunspacePool=$pool
    [void]$ps.AddScript($sb).AddArgument($t.Url).AddArgument($t.Init).AddArgument($t.After).AddArgument($t.Product).AddArgument($t.Grp).AddArgument($t.Sub).AddArgument($script:MaxCacheHours).AddArgument($t.PostTimeout)
    $handle=$ps.BeginInvoke()
    $active.Add([pscustomobject]@{ PS=$ps; Handle=$handle; Tag="$($t.Product)|$($t.InitTag)"; Init=$t.Init; After=$t.After })
  }

  # Process completions and refill slots
  while($active.Count -gt 0){
    # WaitAny: poll handles until one completes
    $doneIdx = -1
    while($doneIdx -lt 0){
      for($ai=0; $ai -lt $active.Count; $ai++){
        if($active[$ai].Handle.IsCompleted){ $doneIdx=$ai; break }
      }
      if($doneIdx -lt 0){ Start-Sleep -Milliseconds 200 }
    }
    $hd=$active[$doneIdx]; $active.RemoveAt($doneIdx)
    try{
      $r=$hd.PS.EndInvoke($hd.Handle)
      $completed++
      Write-Host "    [$completed/$($todo.Count)] $($hd.Tag): $r" -ForegroundColor DarkGray
      $script:FreshThisRun[$hd.Init]=$true; $script:FreshThisRun[$hd.After]=$true
    } catch {
      $failed++
      Write-Warning "    [$($completed+$failed)/$($todo.Count)] $($hd.Tag) failed: $($_.Exception.Message)"
    } finally { $hd.PS.Dispose() }

    # Refill: start next item immediately (no inter-wave pause needed — slot just freed)
    if($queueIdx -lt $todo.Count){
      $t=$todo[$queueIdx]; $queueIdx++
      $ps=[powershell]::Create(); $ps.RunspacePool=$pool
      [void]$ps.AddScript($sb).AddArgument($t.Url).AddArgument($t.Init).AddArgument($t.After).AddArgument($t.Product).AddArgument($t.Grp).AddArgument($t.Sub).AddArgument($script:MaxCacheHours).AddArgument($t.PostTimeout)
      $handle=$ps.BeginInvoke()
      $active.Add([pscustomobject]@{ PS=$ps; Handle=$handle; Tag="$($t.Product)|$($t.InitTag)"; Init=$t.Init; After=$t.After })
    }
  }
  $pool.Close(); $pool.Dispose()
  Write-Host ("  [bulk-prefetch] done in {0:N1}s ({1} OK, {2} failed, {3} total)" -f $sw.Elapsed.TotalSeconds, $completed, $failed, $todo.Count) -ForegroundColor Cyan
}
Start-BulkPrefetch $jobs

# --- Shared Excel COM instance (improvement #2: reuse across jobs) ---
$script:ExcelApp = New-Object -ComObject Excel.Application; $script:ExcelApp.Visible=$false; $script:ExcelApp.DisplayAlerts=$false

# --- Outlook (skip if -NoEmail) ---
$ol=$null; $selfMail='(no-email)'
if(-not $NoEmail){
  $ol=New-Object -ComObject Outlook.Application; $ns=$ol.GetNamespace("MAPI")
  try{ $selfMail=$ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch { $selfMail=$ns.CurrentUser.Address }
}

$summary = @()
$combinedSections = @()   # for -CombineXlsx mode (improvement #5)
foreach($job in $jobs){
  $label = "[Row $($job.Row)] $($job.Product) | $(if($job.Report){$job.Report}else{'(ALL 4)'}) | $($job.Filter)"
  "`n=== $label ==="
  try {
    if($job.Product -in $failedLookups){
      throw "Product '$($job.Product)' not in lookup and could not be inferred"
    }
    $grp = $lkMap[$job.Product].Grp; $sub = $lkMap[$job.Product].Sub

    $reportsToRun = if([string]::IsNullOrWhiteSpace($job.Report)){ $blankReportSet } else { @($job.Report) }

    Start-PrefetchPosts -product $job.Product -grp $grp -sub $sub -reportKeys $reportsToRun

    $sections = @()
    foreach($rk in $reportsToRun){
      # Skip ENG for known-missing products (#4)
      if($rk.ToUpper() -eq 'ENG' -and ($job.Product -match 'Extended Temp')){
        Write-Warning "  skip ENG : known missing for Extended Temp variant"
        continue
      }
      "-- $rk --"
      try {
        $sec = Invoke-TciReport -product $job.Product -grp $grp -sub $sub -reportKey $rk -filter $job.Filter
        $sections += $sec
      } catch {
        Write-Warning "  skip $rk : $($_.Exception.Message)"
      }
    }
    if($sections.Count -eq 0){ throw 'no sections produced' }

    $tagF = if([string]::IsNullOrWhiteSpace($job.Filter)){ '' } else { '_' + (Safe-Name $job.Filter) }
    $repTag = if([string]::IsNullOrWhiteSpace($job.Report)){ 'AllReports' } else { Safe-Name $job.Report }
    $base = "$(Safe-Name $job.Product)_$repTag$tagF"
    $xlsx = Join-Path $OutDir "$base.xlsx"
    if($CombineXlsx -and $jobs.Count -gt 1){
      # Prefix sheet tags with product short name for combined mode
      $shortProd = (Safe-Name $job.Product).Substring(0,[Math]::Min(12,(Safe-Name $job.Product).Length))
      foreach($s in $sections){ $s.SheetTag = "$shortProd $($s.SheetTag)" }
      $combinedSections += $sections
      $xlsx = '(combined)'
    } else {
      Write-MultiTabXlsx -sections $sections -xlsxPath $xlsx -excelApp $script:ExcelApp
    }
    "xlsx: $xlsx ($( if($xlsx -ne '(combined)'){ [math]::Round((Get-Item $xlsx).Length/1KB,1).ToString()+' KB' } else { 'deferred' } ), $($sections.Count) tab(s))"

    $to = if([string]::IsNullOrWhiteSpace($job.Email)){ $selfMail } else { $job.Email }
    $rowsSum = ($sections | ForEach-Object { $_.Rows.Count - 1 } | Measure-Object -Sum).Sum
    $filterDisp = if([string]::IsNullOrWhiteSpace($job.Filter)){ 'none' } else { $job.Filter }
    $reportDisp = if([string]::IsNullOrWhiteSpace($job.Report)){ 'All Reports' } else { $job.Report }

    # log run history (per report tab) for tracking / regression detection
    $colsMax = ($sections | ForEach-Object { if($_.Hdr){ $_.Hdr.Count } else { 0 } } | Measure-Object -Maximum).Maximum
    foreach($s in $sections){
        Add-PhiRunHistory -Product $job.Product -Report $s.SheetTag -Filter $filterDisp `
            -Rows ($s.Rows.Count - 1) -Cols $(if($s.Hdr){ $s.Hdr.Count } else { 0 }) -File $xlsx | Out-Null
    }

    # shared email card (single source of truth in tci_lib.ps1)
    $cardSections = $sections | ForEach-Object { @{ SheetTag = $_.SheetTag; RowCount = $_.Rows.Count - 1 } }
    $body = Build-PhiCard -Product $job.Product -Group $grp -SubGroup $sub -Sections $cardSections -Filter $filterDisp

    if(-not $NoEmail -and (-not $CombineXlsx -or $jobs.Count -eq 1)){
      $mi=$ol.CreateItem(0); $mi.To=$to
      $mi.Subject="PHI of $($job.Product) - $reportDisp$(if($job.Filter){' ' + $job.Filter})"
      $mi.HTMLBody=$body
      $mi.Attachments.Add($xlsx) | Out-Null
      $mi.Send()
      "sent to $to"
    } elseif($CombineXlsx -and $jobs.Count -gt 1){
      "(email deferred to combined send)"
    } else {
      "(email skipped - NoEmail)"
    }

    $summary += [pscustomobject]@{ Row=$job.Row; Product=$job.Product; Report=$job.Report; Filter=$job.Filter; Tabs=$sections.Count; Rows=$rowsSum; Status='OK' }
  } catch {
    Write-Warning "FAILED $label : $($_.Exception.Message)"
    Write-Warning "  at: $($_.InvocationInfo.PositionMessage)"
    Write-Warning "  stack: $($_.ScriptStackTrace)"
    $summary += [pscustomobject]@{ Row=$job.Row; Product=$job.Product; Report=$job.Report; Filter=$job.Filter; Tabs=0; Rows=0; Status="ERROR: $($_.Exception.Message)" }
  }
}

# --- Combined xlsx mode: write one workbook + one email for all products ---
if($CombineXlsx -and $combinedSections.Count -gt 0){
  $tagF = if([string]::IsNullOrWhiteSpace($Filter)){ '' } else { '_' + (Safe-Name $Filter) }
  $repTag = if([string]::IsNullOrWhiteSpace($Report)){ 'AllReports' } else { Safe-Name $Report }
  $comboBase = "Combined_$repTag$tagF"
  $comboXlsx = Join-Path $OutDir "$comboBase.xlsx"
  Write-MultiTabXlsx -sections $combinedSections -xlsxPath $comboXlsx -excelApp $script:ExcelApp
  "combined xlsx: $comboXlsx ($([math]::Round((Get-Item $comboXlsx).Length/1KB,1)) KB, $($combinedSections.Count) tab(s))"
  if(-not $NoEmail){
    $to = if([string]::IsNullOrWhiteSpace($Email)){ $selfMail } else { $Email }
    $prodList = ($jobs | ForEach-Object { $_.Product }) -join ', '
    $mi=$ol.CreateItem(0); $mi.To=$to
    $mi.Subject="PHI Combined: $prodList"
    $mi.HTMLBody="<p>Combined PHI report for: <b>$prodList</b></p><p>$($combinedSections.Count) tabs, filter: $(if($Filter){$Filter}else{'none'})</p>"
    $mi.Attachments.Add($comboXlsx) | Out-Null
    $mi.Send()
    "combined email sent to $to"
  } else {
    "(combined email skipped - NoEmail)"
  }
}

# --- Cleanup shared Excel COM ---
try { $script:ExcelApp.Quit() } catch {}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

"`n=== Batch Summary ==="
$summary | Format-Table Row,Product,Report,Filter,Tabs,Rows,Status -AutoSize
