$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Nova Lake H'
$safeName  = 'Nova_Lake_H'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'PORSDAForecast'
$initTag   = 'sda'
$firstTh   = 'CommonName'

Write-Host "=== $product - SDA Weekly + Class ===" -ForegroundColor Cyan

# --- Fetch / cache ---
$afterFile = Join-Path $env:TEMP "tci_after_${safeName}_${initTag}.html"
$maxAge    = 720

if ((Test-Path $afterFile) -and (Get-Item $afterFile).Length -gt 0 -and (Get-Item $afterFile).LastWriteTime -gt (Get-Date).AddMinutes(-$maxAge)) {
    $age = [int]((Get-Date) - (Get-Item $afterFile).LastWriteTime).TotalMinutes
    Write-Host "[cache] reusing $afterFile (age ${age}m)"
    $html = Get-Content $afterFile -Raw
} else {
    Write-Host "[fetch] GET init + POST..."
    $initHtml = Tci-GetInit -ReportId $reportId -MaxAgeMinutes 60

    $vs  = ([regex]::Match($initHtml, 'name="__VIEWSTATE"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value
    $vsg = ([regex]::Match($initHtml, 'name="__VIEWSTATEGENERATOR"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value
    $ev  = ([regex]::Match($initHtml, 'name="__EVENTVALIDATION"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value

    $phiPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_PHIParameters"(.*?)</div>\s*</div>').Value
    $phiCtls  = [regex]::Matches($phiPanel, 'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }

    $grpPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_Group"(.*?)</div>\s*</div>').Value
    $grpCtl   = ([regex]::Matches($grpPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -eq $grp }).Groups[1].Value

    $subPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_SubGroup"(.*?)</div>\s*</div>').Value
    $subCtl   = ([regex]::Matches($subPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -eq $sub }).Groups[1].Value

    $cnPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_CommonName"(.*?)</div>\s*</div>').Value
    $escaped = [regex]::Escape($product)
    $cnCtl   = ([regex]::Matches($cnPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -match "^${escaped}$" }).Groups[1].Value
    if (-not $cnCtl) {
        $cnCtl = ([regex]::Matches($cnPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -match "^${escaped}" }).Groups[1].Value
    }
    if (-not $cnCtl) { Write-Error "CommonName ctl not found for '$product'"; exit 1 }
    Write-Host "[ctl] CommonName = $cnCtl"

    $form = [ordered]@{
        '__EVENTTARGET'=''; '__EVENTARGUMENT'=''
        '__VIEWSTATE'=$vs; '__VIEWSTATEGENERATOR'=$vsg; '__EVENTVALIDATION'=$ev
        $grpCtl='on'; $subCtl='on'; $cnCtl='on'
        'ctl00$ContentPlaceHolder1$btn_RunReport'='Run Report'
    }
    foreach ($p in $phiCtls) { $form[$p] = 'on' }

    Write-Host "[fetch] POST phi=$($phiCtls.Count)..."
    $resp = Tci-Post -ReportId $reportId -Form $form -TimeoutSec 600 -Retries 1
    $html = $resp.Content
    if (-not $html -or $html.Length -lt 1000) { Write-Error "POST returned empty/short response ($($html.Length) chars)"; exit 1 }
    [IO.File]::WriteAllText($afterFile, $html, [Text.UTF8Encoding]::new($false))
    Write-Host "[fetch] saved: $afterFile ($($html.Length) chars)"
}

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

# --- Class filter ---
$classMonitors = @('MPS','EQA','CS MONITOR')
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim()
    $isClass = ($op -ne '' -and $op -like 'TEST*') -or
               ($op -eq '' -and ($classMonitors | Where-Object { $met -match "(?i)^$_" }).Count -gt 0)
    if ($isClass) { $filtered.Add($allRows[$i]) }
}
Write-Host "Class filter: $($filtered.Count - 1) rows"
if ($filtered.Count -le 1) { Write-Error "No rows after Class filter"; exit 1 }

# --- Output unique metrics and time columns for WIF planning ---
$metrics = @{}
for ($i = 1; $i -lt $filtered.Count; $i++) {
    $k = "$($filtered[$i][$metIdx])".Trim()
    if ($k -and -not $metrics.ContainsKey($k)) { $metrics[$k] = $true }
}
Write-Host "`nAvailable Class metrics ($($metrics.Count)):"
$metrics.Keys | Sort-Object | ForEach-Object { Write-Host "  $_" }

# Time columns (after fixed columns)
$fixedEnd = 20  # typical: columns 0-19 are fixed, 20+ are WW codes
$timeCols = @()
for ($c = $fixedEnd; $c -lt $hdr.Count; $c++) {
    if ($hdr[$c] -match '^\d{6}$') { $timeCols += $hdr[$c] }
}
Write-Host "`nTime columns: $($timeCols[0]) to $($timeCols[-1]) ($($timeCols.Count) weeks)"
