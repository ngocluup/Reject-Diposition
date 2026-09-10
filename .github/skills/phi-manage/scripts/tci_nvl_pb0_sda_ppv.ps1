$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Nova Lake P (B0)'
$safeName  = 'Nova_Lake_P_B0'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'PORSDAForecast'
$initTag   = 'sda'
$firstTh   = 'CommonName'

Write-Host "=== $product - SDA Weekly + PPV ===" -ForegroundColor Cyan

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

    # PHI Parameters
    $phiPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_PHIParameters"(.*?)</div>\s*</div>').Value
    $phiCtls  = [regex]::Matches($phiPanel, 'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }

    # AT_Group/SubGroup ctls
    $grpPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_Group"(.*?)</div>\s*</div>').Value
    $grpCtl   = ([regex]::Matches($grpPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -eq $grp }).Groups[1].Value

    $subPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_SubGroup"(.*?)</div>\s*</div>').Value
    $subCtl   = ([regex]::Matches($subPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -eq $sub }).Groups[1].Value

    # CommonName ctl - escape parenthesis in product name
    $cnPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_CommonName"(.*?)</div>\s*</div>').Value
    $escaped = [regex]::Escape($product)
    $cnCtl   = ([regex]::Matches($cnPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -match "^${escaped}" }).Groups[1].Value
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
$ppIdx  = [array]::IndexOf([string[]]$hdr, 'Proposed/POR')
if ($ppIdx -lt 0) { for ($c = 0; $c -lt $hdr.Count; $c++) { if ($hdr[$c] -match 'Proposed') { $ppIdx = $c; break } } }
Write-Host "Indexes: Op=$opIdx Met=$metIdx PP=$ppIdx"

# --- PPV filter ---
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim()
    $isPPV = ($op -ne '' -and $op -like 'PPV*') -or
             ($op -eq '' -and $met -match '(?i)^PPV-M SAMPLE SIZE')
    if ($isPPV) { $filtered.Add($allRows[$i]) }
}
Write-Host "PPV filter: $($filtered.Count - 1) rows"
if ($filtered.Count -le 1) { Write-Error "No rows after PPV filter"; exit 1 }

# --- Show available metrics and time columns for WIF ---
Write-Host "`n--- PPV Metrics found ---"
$metrics = @{}
for ($i = 1; $i -lt $filtered.Count; $i++) {
    $m = "$($filtered[$i][$metIdx])".Trim()
    if (-not $metrics.ContainsKey($m)) { $metrics[$m] = 0 }
    $metrics[$m]++
}
$metrics.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host "  $($_.Key) ($($_.Value) rows)" }

Write-Host "`n--- Time columns (WW codes) ---"
# Time columns start after the fixed metadata columns (usually around col 20+)
$timeCols = @()
for ($c = $ppIdx + 1; $c -lt $hdr.Count; $c++) {
    if ($hdr[$c] -match '^\d{6}$') { $timeCols += $hdr[$c] }
}
Write-Host "  First: $($timeCols[0])  Last: $($timeCols[-1])  Count: $($timeCols.Count)"
Write-Host "  All: $($timeCols -join ', ')"
