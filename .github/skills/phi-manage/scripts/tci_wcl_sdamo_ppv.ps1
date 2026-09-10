$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Inline helpers ===
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
$product   = 'Wildcat Lake'
$safeName  = 'Wildcat_Lake'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'PORMonthlyForecast'
$initTag   = 'sdamo'
$firstTh   = 'CommonName'

Write-Host "=== $product - SDA Monthly PPV ===" -ForegroundColor Cyan

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

# --- Relabel WW to MMM YYYY (fiscal 4-4-5) ---
$anchors = @(1,5,9,14,18,22,27,31,35,40,44,48)
$months  = @('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')
for ($ci = 0; $ci -lt $hdr.Count; $ci++) {
    if ($hdr[$ci] -match '^\d{6}$') {
        $y = [int]$hdr[$ci].Substring(0,4); $w = [int]$hdr[$ci].Substring(4,2)
        $mi = -1; for ($k = 0; $k -lt 12; $k++) { if ($w -ge $anchors[$k]) { $mi = $k } }
        if ($mi -ge 0) { $hdr[$ci] = '{0} {1}' -f $months[$mi], $y }
    }
}

# --- PPV filter: op startswith PPV OR (blank op AND met = PPV-M SAMPLE SIZE) ---
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim()
    $isPPV = ($op -ne '' -and $op -like 'PPV*') -or
             ($op -eq '' -and $met -eq 'PPV-M SAMPLE SIZE')
    if ($isPPV) { $filtered.Add($allRows[$i]) }
}
$rowCount = $filtered.Count - 1
Write-Host "PPV filter: $rowCount rows"

if ($rowCount -eq 0) { Write-Error "No PPV rows for $product"; exit 1 }

# --- Build xlsx ---
$csv  = Join-Path $env:TEMP "${safeName}_SDAMO_PPV.csv"
$xlsx = Join-Path $env:TEMP "${safeName}_SDAMO_PPV.xlsx"

Tci-WriteCsv -Rows $filtered -Path $csv
$result = Tci-ExportXlsx -CsvPath $csv -XlsxPath $xlsx -SheetName 'WCL SDA Monthly PPV' -Validate -MinRows $rowCount -MinCols 10
Write-Host "xlsx: $($result.Path) ($([Math]::Round((Get-Item $xlsx).Length/1KB, 1)) KB) rows=$($result.Rows) cols=$($result.Cols)"

# --- Email ---
$sections = @(@{ SheetTag = 'SDA Monthly'; RowCount = $rowCount })
$body = Build-PhiCard -Product $product -Group $grp -SubGroup $sub -Sections $sections -Filter 'PPV'
$subj = "PHI of $product - SDA Monthly PPV"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsx)
Write-Host "`nDone."

Remove-Item $csv -Force -ErrorAction SilentlyContinue
