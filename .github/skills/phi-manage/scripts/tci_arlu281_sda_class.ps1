$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# === Config ===
$product   = 'ARL U 2C+8A+GT1'
$safeName  = 'ARL_U_2C_8A_GT1'
$grp       = 'Client'
$sub       = 'Mobile'
$reportId  = 'PORSDAForecast'
$initTag   = 'sda'
$firstTh   = 'CommonName'
$filter    = 'Class'

Write-Host "=== $product - SDA Weekly + Class ===" -ForegroundColor Cyan

# --- Fetch / cache ---
$afterFile = Join-Path $env:TEMP "tci_after_${safeName}_${initTag}.html"
$maxAge    = 720  # 12 hours

if ((Test-Path $afterFile) -and (Get-Item $afterFile).LastWriteTime -gt (Get-Date).AddMinutes(-$maxAge)) {
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

    # AT_Group/SubGroup ctls (from SDA init)
    $grpPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_Group"(.*?)</div>\s*</div>').Value
    $grpCtl   = ([regex]::Matches($grpPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -eq $grp }).Groups[1].Value

    $subPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_SubGroup"(.*?)</div>\s*</div>').Value
    $subCtl   = ([regex]::Matches($subPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -eq $sub }).Groups[1].Value

    $form = [ordered]@{
        '__EVENTTARGET'        = ''
        '__EVENTARGUMENT'      = ''
        '__VIEWSTATE'          = $vs
        '__VIEWSTATEGENERATOR' = $vsg
        '__EVENTVALIDATION'    = $ev
        $grpCtl                = 'on'
        $subCtl                = 'on'
        'ctl00$ContentPlaceHolder1$ctl43' = 'on'  # ARL U 2C+8A+GT1
        'ctl00$ContentPlaceHolder1$btn_RunReport' = 'Run Report'
    }
    foreach ($p in $phiCtls) { $form[$p] = 'on' }

    Write-Host "[fetch] POST phi=$($phiCtls.Count)..."
    $resp = Tci-Post -ReportId $reportId -Form $form -TimeoutSec 600 -Retries 1
    $html = $resp.Content
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

# --- Class filter ---
# Class = op startswith TEST OR (blank op AND metric IN MPS/EQA/CS MONITOR)
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

if ($filtered.Count -le 1) {
    Write-Host "WARNING: no Class rows found!" -ForegroundColor Yellow
    exit 1
}

# --- CSV + Excel ---
$csvFile  = Join-Path $env:TEMP "${safeName}_SDA_Class.csv"
$xlsxFile = Join-Path $env:TEMP "${safeName}_SDA_Class.xlsx"

Tci-WriteCsv -Rows $filtered -Path $csvFile
$result = Tci-ExportXlsx -CsvPath $csvFile -XlsxPath $xlsxFile -SheetName 'SDA Class' -FreezeColumns 4 -Validate -MinRows 1
Remove-Item $csvFile -Force -ErrorAction SilentlyContinue
Write-Host "xlsx: $xlsxFile ($([Math]::Round((Get-Item $xlsxFile).Length/1KB, 1)) KB)"

# --- Email ---
$sections = @( @{ SheetTag = 'SDA Weekly Class'; RowCount = $filtered.Count - 1 } )
$body = Build-PhiCard -Product $product -Group $grp -SubGroup $sub -Sections $sections -Filter $filter
$subj = "PHI of $product - SDA Weekly Class"
Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxFile)
Write-Host "`nDone."
