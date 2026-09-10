$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Nova Lake AX 16C'
$safeName  = 'Nova_Lake_AX_16C'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'PORMonthlyForecast'
$initTag   = 'sdamo'
$firstTh   = 'CommonName'
$filter    = 'PPV'

Write-Host "=== $product - SDA Monthly + PPV ===" -ForegroundColor Cyan

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
    if (-not $html -or $html.Length -lt 1000) { Write-Error "POST returned empty/short response"; exit 1 }
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

# --- Relabel monthly WW columns to MMM YYYY ---
function WwToMonth($yyyyww) {
    if ($yyyyww -notmatch '^\d{6}$') { return $yyyyww }
    $y = [int]$yyyyww.Substring(0,4); $w = [int]$yyyyww.Substring(4,2)
    $jan1 = [datetime]::new($y,1,1)
    $ww01Sun = $jan1.AddDays(-[int]$jan1.DayOfWeek)
    $targetSun = $ww01Sun.AddDays(7*($w-1))
    return $targetSun.ToString("MMM yyyy")
}
$newHdr = @()
for ($c = 0; $c -lt $hdr.Count; $c++) { $newHdr += WwToMonth $hdr[$c] }
$filtered[0] = $newHdr

# --- CSV + Excel ---
$csvFile  = Join-Path $env:TEMP "${safeName}_SDAMO_PPV.csv"
$xlsxFile = Join-Path $env:TEMP "${safeName}_SDAMO_PPV.xlsx"

Tci-WriteCsv -Rows $filtered -Path $csvFile
$result = Tci-ExportXlsx -CsvPath $csvFile -XlsxPath $xlsxFile -SheetName 'SDA Monthly PPV' -FreezeColumns 4 -Validate -MinRows 1
Remove-Item $csvFile -Force -ErrorAction SilentlyContinue
Write-Host "xlsx: $xlsxFile ($([Math]::Round((Get-Item $xlsxFile).Length/1KB, 1)) KB)"

# --- Email ---
$sections = @( @{ SheetTag = 'SDA Monthly PPV'; RowCount = $filtered.Count - 1 } )
$body = Build-PhiCard -Product $product -Group $grp -SubGroup $sub -Sections $sections -Filter $filter
$subj = "PHI of $product - SDA Monthly PPV"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxFile)
Write-Host "`nDone."
