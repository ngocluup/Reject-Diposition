$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

$product   = 'Nova Lake AX 28C'
$safeName  = 'Nova_Lake_AX_28C'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'PORSDAForecast'
$initTag   = 'sda'
$firstTh   = 'CommonName'

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
    $cnCtl   = ([regex]::Matches($cnPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -match "^${escaped}" }).Groups[1].Value
    if (-not $cnCtl) { Write-Error "CommonName ctl not found for '$product'"; exit 1 }

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
    if (-not $html -or $html.Length -lt 1000) { Write-Error "POST returned empty/short"; exit 1 }
    [IO.File]::WriteAllText($afterFile, $html, [Text.UTF8Encoding]::new($false))
    Write-Host "[fetch] saved: $afterFile ($($html.Length) chars)"
}

$tables = Tci-ParseTables $html
$dt = Tci-PickDataTable $tables $firstTh
$parsed = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr = $allRows[0]

$opIdx = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx = [array]::IndexOf([string[]]$hdr, 'MetricName')

Write-Host "Total rows: $($allRows.Count - 1)"
Write-Host ""

$ops = @()
for ($i = 1; $i -lt $allRows.Count; $i++) { $ops += "$($allRows[$i][$opIdx])".Trim() }

$testCount = ($ops | Where-Object { $_ -like 'TEST*' }).Count
$ppvCount  = ($ops | Where-Object { $_ -like 'PPV*' }).Count
$biCount   = ($ops | Where-Object { $_ -like 'BURNIN*' }).Count
$blankOps  = ($ops | Where-Object { $_ -eq '' }).Count

# Count monitors in blank-op
$classMonitors = @('MPS','EQA','CS MONITOR')
$monCount = 0
for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op = "$($allRows[$i][$opIdx])".Trim()
    $met = "$($allRows[$i][$metIdx])".Trim()
    if ($op -eq '' -and ($classMonitors | Where-Object { $met -match "(?i)^$_" }).Count -gt 0) { $monCount++ }
}

$classTotal = $testCount + $monCount

Write-Host "Available filters:"
Write-Host "  Class (TEST + monitors): $classTotal rows (TEST=$testCount, monitors=$monCount)"
Write-Host "  PPV: $ppvCount rows"
Write-Host "  BI: $biCount rows"
Write-Host "  All: $($allRows.Count - 1) rows"

Write-Host ""
Write-Host "Unique operations:"
$ops | Sort-Object -Unique | ForEach-Object { if ($_ -eq '') { Write-Host "  (blank)" } else { Write-Host "  $_" } }
