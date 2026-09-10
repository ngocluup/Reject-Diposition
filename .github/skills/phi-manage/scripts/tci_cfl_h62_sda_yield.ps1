$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Inline helpers (not in lib) ===
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
        if ($m2.Success) { $m2.Groups[1].Value } else { '' }
    }
}

# === Config ===
$product   = 'CFL H62'
$safeName  = 'CFL_H62'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'PORSDAForecast'
$initTag   = 'sda'
$firstTh   = 'CommonName'

Write-Host "=== $product - SDA Weekly Yield ===" -ForegroundColor Cyan

# --- Fetch or reuse cache ---
$afterFile = Join-Path $env:TEMP "tci_after_${safeName}_${initTag}.html"
$maxAge    = 720  # 12h

if ((Test-Path $afterFile) -and (Get-Item $afterFile).Length -gt 1000 -and
    (Get-Item $afterFile).LastWriteTime -gt (Get-Date).AddMinutes(-$maxAge)) {
    $age = [int]((Get-Date) - (Get-Item $afterFile).LastWriteTime).TotalMinutes
    Write-Host "[cache] reusing $afterFile (age ${age}m)"
} else {
    Write-Host "[fetch] GET init..."
    $initHtml = Tci-GetInit -ReportId $reportId
    
    $grpCtl  = Get-CB $initHtml 'ContentPlaceHolder1_Filters_AT_Group' $grp
    $subCtl  = Get-CB $initHtml 'ContentPlaceHolder1_Filters_AT_SubGroup' $sub
    $prodCtl = Get-CB $initHtml 'ContentPlaceHolder1_Filters_CommonName' $product
    $phis    = Get-AllCB $initHtml 'ContentPlaceHolder1_Filters_PHIParameters'
    
    if (-not $prodCtl) { Write-Error "CommonName '$product' not found"; exit 1 }
    Write-Host "[fetch] grp=$grpCtl sub=$subCtl prod=$prodCtl phis=$($phis.Count)"
    
    $body = [ordered]@{
        '__EVENTTARGET'=''
        '__EVENTARGUMENT'=''
        '__VIEWSTATE'         = HF $initHtml '__VIEWSTATE'
        '__VIEWSTATEGENERATOR'= HF $initHtml '__VIEWSTATEGENERATOR'
        '__EVENTVALIDATION'   = HF $initHtml '__EVENTVALIDATION'
        $grpCtl  = 'on'
        $subCtl  = 'on'
        $prodCtl = 'on'
        'ctl00$ContentPlaceHolder1$btn_RunReport' = 'Run Report'
    }
    foreach ($p in $phis) { $body[$p.Name] = 'on' }
    
    Write-Host "[fetch] POST..."
    $resp = Tci-Post -ReportId $reportId -Form $body
    Set-Content $afterFile $resp
    Write-Host "[fetch] saved $afterFile ($((Get-Item $afterFile).Length) chars)"
}

$html = Get-Content $afterFile -Raw
Write-Host "html length: $($html.Length)"

# --- Parse ---
$tables = Tci-ParseTables $html
$dt     = Tci-PickDataTable $tables $firstTh
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr     = $allRows[0]
Write-Host "total rows: $($allRows.Count - 1) cols: $($hdr.Count)"

$opIdx  = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx = [array]::IndexOf([string[]]$hdr, 'MetricName')
Write-Host "Indexes: Op=$opIdx Met=$metIdx"

# --- Yield filter: blank op + MetricName IN (U/D, R/D, FINISH YIELD) ---
$yieldMetrics = @('U/D','R/D','FINISH YIELD')
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim().ToUpper()
    if ([string]::IsNullOrWhiteSpace($op) -and $yieldMetrics -contains $met) {
        $filtered.Add($allRows[$i])
    }
}
$rowCount = $filtered.Count - 1
Write-Host "Yield filter: $rowCount rows"

if ($rowCount -eq 0) {
    Write-Host "No yield rows found. Checking blank-op metrics..."
    for ($i = 1; $i -lt $allRows.Count; $i++) {
        $op = "$($allRows[$i][$opIdx])".Trim()
        if ([string]::IsNullOrWhiteSpace($op)) {
            Write-Host "  blank-op metric: $($allRows[$i][$metIdx])"
        }
    }
    Write-Error "No yield rows (U/D, R/D, FINISH YIELD) for $product"
    exit 1
}

# --- Build xlsx ---
$csv  = Join-Path $env:TEMP "${safeName}_SDA_Yield.csv"
$xlsx = Join-Path $env:TEMP "${safeName}_SDA_Yield.xlsx"

Tci-WriteCsv -Rows $filtered -Path $csv
$result = Tci-ExportXlsx -CsvPath $csv -XlsxPath $xlsx -SheetName 'CFL H62 Yield' -Validate -MinRows $rowCount -MinCols 10
Write-Host "xlsx: $($result.Path) ($([Math]::Round((Get-Item $xlsx).Length/1KB, 1)) KB) rows=$($result.Rows) cols=$($result.Cols)"

# --- Email ---
$sections = @(@{ SheetTag = 'SDA Weekly'; RowCount = $rowCount })
$body = Build-PhiCard -Product $product -Group $grp -SubGroup $sub -Sections $sections -Filter 'Yield'
$subj = "PHI of $product - SDA Weekly Yield"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsx)
Write-Host "`nDone."

Remove-Item $csv -Force -ErrorAction SilentlyContinue
