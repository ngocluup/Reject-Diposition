$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'Nova Lake AX 16C'
$safeName  = 'Nova_Lake_AX_16C'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'ENGForecasts'
$initTag   = 'eng'
$firstTh   = 'ATGroup'
$filter    = 'PPV'

Write-Host "=== $product - ENG + PPV ===" -ForegroundColor Cyan

# --- Fetch / cache ---
$afterFile = Join-Path $env:TEMP "tci_after_${safeName}_${initTag}.html"
$maxAge    = 720

if ((Test-Path $afterFile) -and (Get-Item $afterFile).Length -gt 0 -and (Get-Item $afterFile).LastWriteTime -gt (Get-Date).AddMinutes(-$maxAge)) {
    $age = [int]((Get-Date) - (Get-Item $afterFile).LastWriteTime).TotalMinutes
    Write-Host "[cache] reusing $afterFile (age ${age}m)"
    $html = Get-Content $afterFile -Raw
} else {
    Write-Host "[fetch] GET init..."
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
    if (-not $html -or $html.Length -lt 1000) {
        Set-Content $afterFile $resp
        $html = Get-Content $afterFile -Raw
    }
    if ($html.Length -lt 1000) { Write-Error "POST returned empty/short response"; exit 1 }
    [IO.File]::WriteAllText($afterFile, $html, [Text.UTF8Encoding]::new($false))
    Write-Host "[fetch] saved: $afterFile ($($html.Length) chars)"
}

# --- Parse ---
$tables = Tci-ParseTables $html
$dt     = Tci-PickDataTable $tables $firstTh
if (-not $dt) {
    # Check for no-data
    if ($html -match '(?i)no data|no records') {
        Write-Host "NO DATA returned for $product on ENG report."
        Write-Host "This product may not be in engineering phase (ENG only has NPI products)."
        exit 1
    }
    Write-Error "Data table not found (first th=$firstTh)"; exit 1
}
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
             ($op -eq '' -and $met -match '(?i)^PPV-M SAMPLE SIZE$')
    if ($isPPV) { $filtered.Add($allRows[$i]) }
}
Write-Host "PPV filter: $($filtered.Count - 1) rows"
if ($filtered.Count -le 1) { Write-Error "No rows after PPV filter"; exit 1 }

# --- CSV + Excel ---
$csvFile  = Join-Path $env:TEMP "${safeName}_ENG_PPV.csv"
$xlsxFile = Join-Path $env:TEMP "${safeName}_ENG_PPV.xlsx"

Tci-WriteCsv -Rows $filtered -Path $csvFile
$result = Tci-ExportXlsx -CsvPath $csvFile -XlsxPath $xlsxFile -SheetName 'ENG PPV' -FreezeColumns 4 -Validate -MinRows 1
Remove-Item $csvFile -Force -ErrorAction SilentlyContinue
Write-Host "xlsx: $xlsxFile ($([Math]::Round((Get-Item $xlsxFile).Length/1KB, 1)) KB)"

# --- Email ---
$sections = @( @{ SheetTag = 'ENG PPV'; RowCount = $filtered.Count - 1 } )
$body = Build-PhiCard -Product $product -Group $grp -SubGroup $sub -Sections $sections -Filter $filter
$subj = "PHI of $product - ENG PPV"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxFile)
Write-Host "`nDone."
