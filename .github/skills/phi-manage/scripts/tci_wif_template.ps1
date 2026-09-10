##############################################################################
# tci_wif_template.ps1 — Reusable WIF generator for TCI SDA Weekly + Class
#
# Usage: Define $config hashtable before dot-sourcing, or set variables then
#        call the main logic. See examples at bottom.
#
# Required variables (set before running):
#   $ProductCache  - path to cached HTML (e.g. "$env:TEMP\tci_after_ptlu404_sda.html")
#   $ShortName     - display name (e.g. 'PTL U404')
#   $ProductLabel  - full CommonName (e.g. 'Panther Lake U 4P+0E+4LP_E')
#   $SheetName     - Excel sheet name (e.g. 'PTL U404 Weekly Class WIF')
#   $OutputBase    - output filename base without extension (e.g. 'PTL_U404_WIF_custom')
#   $WifSpecs      - array of hashtables, each with:
#       Label    - display name for the spec
#       Op       - OperationName to match ('' = blank)
#       Sub      - SubObject to match ('' = blank)
#       Met      - MetricName to match (case-insensitive)
#       ColStart - first WW column code (e.g. '202630')
#       ColEnd   - last WW column code (e.g. '202643') or 'EOL'
#       Value    - value to write into WIF cells (e.g. '700', '12%')
#       Site     - [optional] Site column value to match (e.g. 'SS', 'PG8')
##############################################################################

$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

$csv  = "$env:TEMP\$OutputBase.csv"
$xlsx = "$env:TEMP\$OutputBase.xlsx"

if (-not (Test-Path $ProductCache)) { throw "No cached HTML at $ProductCache" }
Write-Host "[cache] reusing $ProductCache"

# ===== Parse HTML tables =====
$c = Get-Content $ProductCache -Raw
$tables = New-Object System.Collections.Generic.List[string]; $idx = 0
while ($idx -lt $c.Length) {
  $om = [regex]::Match($c.Substring($idx), '<table\b[^>]*>'); if (-not $om.Success) { break }
  $sI = $idx + $om.Index + $om.Length; $depth = 1; $pos = $sI
  while ($depth -gt 0 -and $pos -lt $c.Length) {
    $o = $c.IndexOf('<table', $pos); $cl = $c.IndexOf('</table>', $pos)
    if ($cl -lt 0) { $depth = 0; break }
    if ($o -ge 0 -and $o -lt $cl) { $depth++; $pos = $o + 6 } else { $depth--; $pos = $cl + 8 }
  }
  $tables.Add($c.Substring($sI, ($pos - 8) - $sI)); $idx = $sI
}
$dt = $tables | Where-Object {
  $n = [regex]::Matches($_, '(?is)<th[^>]*>\s*([^<]{1,60})') | ForEach-Object { $_.Groups[1].Value.Trim() }
  ($n.Count -gt 0) -and ($n[0] -eq 'CommonName')
}
$all = New-Object System.Collections.Generic.List[object]; $hdr = $null
foreach ($body2 in $dt) {
  $rs = [regex]::Matches($body2, '(?is)<tr[^>]*>(.*?)</tr>')
  $tr = New-Object System.Collections.Generic.List[object]
  foreach ($r in $rs) {
    $cm = [regex]::Matches($r.Groups[1].Value, '(?is)<(t[hd])\b[^>]*>(.*?)</\1>')
    if ($cm.Count -eq 0) { continue }
    $cells = foreach ($mm in $cm) { $t = [regex]::Replace($mm.Groups[2].Value, '(?is)<[^>]+>', ' '); $t = [System.Net.WebUtility]::HtmlDecode($t).Trim() -replace '\s+', ' '; , $t }
    $tr.Add(@($cells))
  }
  if (-not $hdr) { $hdr = $tr[0]; $all.Add($hdr) }
  for ($i = 1; $i -lt $tr.Count; $i++) { $all.Add($tr[$i]) }
}
Write-Host "total rows: $($all.Count - 1) cols: $($hdr.Count)"

# ===== Column indexes =====
$iOp   = [array]::IndexOf($hdr, 'OperationName')
$iSub  = [array]::IndexOf($hdr, 'SubObject')
$iMet  = [array]::IndexOf($hdr, 'MetricName')
$iPP   = [array]::IndexOf($hdr, 'Proposed/POR')
$iBOM  = [array]::IndexOf($hdr, 'BOM')
$iSite = [array]::IndexOf($hdr, 'Site')
Write-Host "Indexes: Op=$iOp Sub=$iSub Met=$iMet PP=$iPP BOM=$iBOM Site=$iSite"

# ===== Class filter =====
$monitorMetrics = @('MPS MONITOR', 'EQA MONITOR', 'CS MONITOR')
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
$nTest = 0; $nMon = 0
for ($k = 1; $k -lt $all.Count; $k++) {
  $row = $all[$k]
  $op  = ($row[$iOp] + '').Trim()
  $met = ($row[$iMet] + '').Trim().ToUpper()
  if ($op.ToUpper().StartsWith('TEST')) { $filtered.Add($row); $nTest++ }
  elseif ([string]::IsNullOrWhiteSpace($op) -and ($monitorMetrics -contains $met)) { $filtered.Add($row); $nMon++ }
}
Write-Host "Class filter: TEST_*=$nTest monitor=$nMon total=$($filtered.Count - 1)"

# ===== Process WIF specs =====
$rowRedCols = @{}
$totalClones = 0

foreach ($spec in $WifSpecs) {
  Write-Host ""
  Write-Host "--- $($spec.Label): Op='$($spec.Op)' Sub='$($spec.Sub)' Met='$($spec.Met)' Site='$($spec.Site)' range=$($spec.ColStart)-$($spec.ColEnd) value=$($spec.Value) ---"

  # Resolve target columns
  $targetCols = New-Object System.Collections.Generic.List[int]
  $startWw = [int]$spec.ColStart
  $endWw = if ($spec.ColEnd -eq 'EOL') { [int]::MaxValue } else { [int]$spec.ColEnd }
  for ($ci = 0; $ci -lt $hdr.Count; $ci++) {
    if ($hdr[$ci] -match '^\d{6}$') {
      $ww = [int]$hdr[$ci]
      if ($ww -ge $startWw -and $ww -le $endWw) { $targetCols.Add($ci) }
    }
  }
  Write-Host "  Target cols: $($targetCols.Count)"
  if ($targetCols.Count -eq 0) { Write-Host "  WARNING: No columns in range!"; continue }

  # Match rows
  $matches = New-Object System.Collections.Generic.List[object]
  for ($k = 1; $k -lt $filtered.Count; $k++) {
    $r = $filtered[$k]
    $rowOp   = ($r[$iOp] + '').Trim()
    $rowSub  = ($r[$iSub] + '').Trim()
    $rowMet  = ($r[$iMet] + '').Trim().ToUpper()
    $rowSite = if ($iSite -ge 0) { ($r[$iSite] + '').Trim() } else { '' }
    $rowBom  = if ($iBOM -ge 0 -and $iBOM -lt $r.Count) { ($r[$iBOM] + '').Trim() } else { '' }

    $opMatch   = ($spec.Op -eq '' -and [string]::IsNullOrWhiteSpace($rowOp)) -or ($spec.Op -ne '' -and $rowOp -eq $spec.Op)
    $subMatch  = ($spec.Sub -eq '' -and [string]::IsNullOrWhiteSpace($rowSub)) -or ($spec.Sub -ne '' -and $rowSub -eq $spec.Sub)
    $metMatch  = $rowMet -eq $spec.Met.ToUpper()
    $siteMatch = if ($spec.ContainsKey('Site') -and $spec.Site -ne '') { $rowSite -eq $spec.Site } else { $true }
    $bomMatch  = [string]::IsNullOrWhiteSpace($rowBom)

    if ($opMatch -and $subMatch -and $metMatch -and $siteMatch -and $bomMatch) {
      $matches.Add($r)
    }
  }
  Write-Host "  Matched: $($matches.Count)"
  foreach ($m in $matches) { Write-Host "    Res=$($m[5]) Site=$($m[$iSite]) PP=$($m[$iPP]) $($spec.ColStart)(POR)=$($m[$targetCols[0]])" }
  if ($matches.Count -eq 0) { Write-Host "  WARNING: No rows found for $($spec.Label)!"; continue }

  # Clone matched rows as WIF
  foreach ($m in $matches) {
    $clone = [string[]]::new($m.Length)
    for ($i = 0; $i -lt $m.Length; $i++) { $clone[$i] = $m[$i] }
    $clone[$iPP] = 'WIF'
    foreach ($ci in $targetCols) { $clone[$ci] = $spec.Value }
    $filtered.Add($clone)
    $rowRedCols[$filtered.Count] = $targetCols.ToArray()
  }
  $totalClones += $matches.Count
}

Write-Host ""
Write-Host "Final rows: $($filtered.Count - 1) (with $totalClones WIF clones)"
$filtered[0] = $hdr

# ===== Write CSV =====
$maxCols = ($filtered | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
$sb = New-Object System.Text.StringBuilder
foreach ($row in $filtered) {
  $p = @($row) + (, '') * ($maxCols - $row.Count)
  $e = $p | ForEach-Object { $v = "$_"; if ($v -match '[",\r\n]') { '"' + ($v -replace '"', '""') + '"' } else { $v } }
  [void]$sb.AppendLine(($e -join ','))
}
[IO.File]::WriteAllText($csv, $sb.ToString(), [Text.UTF8Encoding]::new($true))

# ===== Excel formatting =====
$excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($csv); $ws = $wb.Worksheets.Item(1); $ws.Name = $SheetName
  $lastCol = $ws.UsedRange.Columns.Count
  $r = $ws.Range($ws.Cells(1, 1), $ws.Cells(1, $lastCol)); $r.Font.Bold = $true; $r.Interior.Color = 0xD9D9D9
  $ppCol = $iPP + 1
  for ($row = 2; $row -le $filtered.Count; $row++) {
    if ($ws.Cells($row, $ppCol).Text -eq 'WIF') {
      $ws.Range($ws.Cells($row, 1), $ws.Cells($row, $lastCol)).Interior.Color = 0x66FFFF
      if ($rowRedCols.ContainsKey($row)) {
        foreach ($ci in $rowRedCols[$row]) { $ws.Cells($row, $ci + 1).Font.Color = 0x0000FF }
      }
    }
  }
  $ws.Application.ActiveWindow.SplitColumn = 4; $ws.Application.ActiveWindow.SplitRow = 1; $ws.Application.ActiveWindow.FreezePanes = $true
  $ws.Columns.AutoFit() | Out-Null
  if (Test-Path $xlsx) { Remove-Item $xlsx -Force }
  $wb.SaveAs($xlsx, 51); try { $wb.Close($false) } catch {}
} finally { try { $excel.Quit() } catch {}; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
Write-Host "xlsx: $xlsx ($([math]::Round((Get-Item $xlsx).Length/1KB,1)) KB)"

# ===== Email =====
$ol = New-Object -ComObject Outlook.Application; $ns = $ol.GetNamespace("MAPI")
try { $me = $ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress } catch { $me = $ns.CurrentUser.Address }
$specSummary = ($WifSpecs | ForEach-Object { "$($_.Label) $($_.ColStart)-$($_.ColEnd)=$($_.Value)" }) -join '; '
$mi = $ol.CreateItem(0); $mi.To = $me
$mi.Subject = "TCI $ShortName - WIF: $specSummary"
$mi.HTMLBody = "<p>$ShortName ($ProductLabel) SDA Weekly Class with <b>$totalClones</b> WIF clone(s). WIF rows yellow; changed cells red font.</p><ul>$(foreach($s in $WifSpecs){"<li><b>$($s.Label)</b>: $($s.Op)/$($s.Sub)/$($s.Met) $(if($s.Site){"Site=$($s.Site) "})$($s.ColStart)-$($s.ColEnd) = $($s.Value)</li>"})</ul>"
$mi.Attachments.Add($xlsx) | Out-Null
$mi.Send()
Write-Host "sent to $me"
