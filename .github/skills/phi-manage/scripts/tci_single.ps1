<#
.SYNOPSIS Universal TCI PHI report runner.
.DESCRIPTION Fetches a single TCI report for a given product, applies an operation
filter, optionally applies WIF, builds xlsx, and emails. Replaces per-product scripts.

.PARAMETER Product   Exact CommonName (e.g. 'Wildcat Lake', 'CFL H62').
.PARAMETER Report    SDA | SDAMonthly | MOR | ENG  (shortcuts resolved internally).
.PARAMETER Filter    Class | PPV | BI | Yield | All  (default: All = no filter).
.PARAMETER Wif       Comma-separated WIF specs: 'PPVs ETT QS 25, Classhot ETT ww40-52 9'.
                     If omitted, no WIF applied.
.PARAMETER NoEmail   Skip emailing (just produce xlsx).
.PARAMETER ForceFetch  Ignore cache, re-POST from TCI.

.EXAMPLE
  .\tci_single.ps1 -Product 'Wildcat Lake' -Report MOR -Filter PPV
  .\tci_single.ps1 -Product 'CFL H62' -Report SDA -Filter Yield
  .\tci_single.ps1 -Product 'Nova Lake AX 16C' -Report ENG -Filter PPV -Wif 'PPVs ETT QS 25'
#>
param(
    [Parameter(Mandatory)][string]$Product,
    [Parameter(Mandatory)][ValidateSet('SDA','SDAMonthly','MOR','ENG')][string]$Report,
    [ValidateSet('Class','PPV','BI','Yield','All')][string]$Filter = 'All',
    [string]$Wif,
    [switch]$NoEmail,
    [switch]$ForceFetch
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# ============================================================================
# Resolve report config
# ============================================================================
$reportMap = @{
    'SDA'        = @{ Id='PORSDAForecast';    Tag='sda';       FirstTh='CommonName';  Relabel='none' }
    'SDAMonthly' = @{ Id='PORMonthlyForecast'; Tag='sdamo';     FirstTh='CommonName';  Relabel='ww2month' }
    'MOR'        = @{ Id='MORSpread';          Tag='morspread'; FirstTh='MetricName';  Relabel='none' }
    'ENG'        = @{ Id='ENGForecasts';       Tag='eng';       FirstTh='ATGroup';     Relabel='none' }
}
$rc = $reportMap[$Report]

# ============================================================================
# Resolve product group/subgroup from lookup JSON
# ============================================================================
$lookupPath = Join-Path $PSScriptRoot 'tci_commonname_lookup.json'
if (-not (Test-Path $lookupPath)) { throw "Lookup file not found: $lookupPath" }
$lookup = (Get-Content $lookupPath -Raw | ConvertFrom-Json).Lookup
$hit = $lookup | Where-Object { $_.CommonName -eq $Product }
if (-not $hit) { throw "Product '$Product' not found in lookup JSON. Run lookup refresh." }
$grp = $hit.AT_Group
$sub = $hit.AT_SubGroup

$safeName = ($Product -replace '[^A-Za-z0-9]','_')
$reportLabel = switch($Report) { 'SDA'{'SDA Weekly'} 'SDAMonthly'{'SDA Monthly'} 'MOR'{'MOR Spread'} 'ENG'{'ENG'} }
$filterLabel = if ($Filter -eq 'All') { '' } else { $Filter }
$sheetName = "$Product $reportLabel$(if($filterLabel){' '+$filterLabel})"
if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }

Write-Host "=== $Product - $reportLabel $filterLabel ===" -ForegroundColor Cyan
Write-Host "  Group=$grp SubGroup=$sub Report=$($rc.Id)"

# ============================================================================
# Fetch or reuse cache
# ============================================================================
$afterFile = Join-Path $env:TEMP "tci_after_${safeName}_$($rc.Tag).html"
$maxAge = 720  # 12h in minutes

$useCache = (-not $ForceFetch) -and (Test-Path $afterFile) -and
            (Get-Item $afterFile).Length -gt 1000 -and
            (Get-Item $afterFile).LastWriteTime -gt (Get-Date).AddMinutes(-$maxAge)

if ($useCache) {
    $age = [int]((Get-Date) - (Get-Item $afterFile).LastWriteTime).TotalMinutes
    Write-Host "[cache] reusing (age ${age}m)"
} else {
    Write-Host "[fetch] GET init..."
    $initHtml = Tci-GetInit -ReportId $rc.Id

    $body = Build-TciPostBody -InitHtml $initHtml -Group $grp -SubGroup $sub -Product $Product -SelectAllPhi
    Write-Host "[fetch] POST ($($rc.Id))..."
    $resp = Tci-Post -ReportId $rc.Id -Form $body
    Set-Content $afterFile $resp
    Write-Host "[fetch] saved ($((Get-Item $afterFile).Length) chars)"
}

$html = Get-Content $afterFile -Raw

# ============================================================================
# Parse data table
# ============================================================================
$tables  = Tci-ParseTables $html
$dt      = Tci-PickDataTable $tables $rc.FirstTh
$parsed  = Tci-RowsFromTable $dt
$allRows = $parsed.Rows
$hdr     = $allRows[0]

$opIdx  = [array]::IndexOf([string[]]$hdr, 'OperationName')
$metIdx = [array]::IndexOf([string[]]$hdr, 'MetricName')
Write-Host "total rows: $($allRows.Count - 1) cols: $($hdr.Count) (Op=$opIdx Met=$metIdx)"

# ============================================================================
# Relabel time columns (SDA Monthly only)
# ============================================================================
if ($rc.Relabel -eq 'ww2month') {
    $anchors = @(1,5,9,14,18,22,27,31,35,40,44,48)
    $months  = @('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')
    for ($ci = 0; $ci -lt $hdr.Count; $ci++) {
        if ($hdr[$ci] -match '^\d{6}$') {
            $y = [int]$hdr[$ci].Substring(0,4); $w = [int]$hdr[$ci].Substring(4,2)
            $mi = -1; for ($k = 0; $k -lt 12; $k++) { if ($w -ge $anchors[$k]) { $mi = $k } }
            if ($mi -ge 0) { $hdr[$ci] = '{0} {1}' -f $months[$mi], $y }
        }
    }
}

# ============================================================================
# Apply operation filter
# ============================================================================
$filtered = New-Object System.Collections.Generic.List[object]
$filtered.Add($hdr)

for ($i = 1; $i -lt $allRows.Count; $i++) {
    $op  = "$($allRows[$i][$opIdx])".Trim()
    $met = if ($metIdx -ge 0) { "$($allRows[$i][$metIdx])".Trim() } else { '' }

    $keep = switch ($Filter) {
        'Class' {
            ($op -ne '' -and $op -like 'TEST*') -or
            ($op -eq '' -and $met.ToUpper() -in @('MPS MONITOR','EQA MONITOR','CS MONITOR'))
        }
        'PPV' {
            ($op -ne '' -and $op -like 'PPV*') -or
            ($op -eq '' -and $met -eq 'PPV-M SAMPLE SIZE')
        }
        'BI' {
            $op -like 'BURNIN*'
        }
        'Yield' {
            [string]::IsNullOrWhiteSpace($op) -and $met.ToUpper() -in @('U/D','R/D','FINISH YIELD')
        }
        'All' { $true }
    }
    if ($keep) { $filtered.Add($allRows[$i]) }
}

$rowCount = $filtered.Count - 1
Write-Host "$Filter filter: $rowCount rows"

if ($rowCount -eq 0) {
    Write-Host "No rows after filter. Available blank-op metrics:" -ForegroundColor Yellow
    for ($i = 1; $i -lt $allRows.Count; $i++) {
        $op = "$($allRows[$i][$opIdx])".Trim()
        if ([string]::IsNullOrWhiteSpace($op) -and $metIdx -ge 0) {
            Write-Host "  $($allRows[$i][$metIdx])"
        }
    }
    Write-Host "Available operations:" -ForegroundColor Yellow
    $ops = @{}
    for ($i = 1; $i -lt $allRows.Count; $i++) {
        $op = "$($allRows[$i][$opIdx])".Trim()
        if ($op -ne '') { if (-not $ops.ContainsKey($op)) { $ops[$op] = 0 }; $ops[$op]++ }
    }
    $ops.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host "  $($_.Name): $($_.Value) rows" }
    Write-Error "No rows match filter '$Filter' for $Product on $reportLabel"
    exit 1
}

# ============================================================================
# WIF processing (optional)
# ============================================================================
$wifApplied = $false
$wifRows = $null

if ($Wif) {
    $wifApplied = $true
    $porIdx = [array]::IndexOf([string[]]$hdr, 'Proposed/POR')

    # Parse WIF specs: "PPVs ETT QS 25" or "Classhot ETT ww40-52 9"
    # For ENG: target is a milestone column (ES0/ES1/ES2/QS/PO)
    # For SDA: target is a time range (ww codes or month labels)
    $wifSpecs = $Wif -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

    $wifFiltered = New-Object System.Collections.Generic.List[object]
    $wifFiltered.Add($hdr)
    $rowRedCols = @{}

    # Add all POR rows first
    for ($i = 1; $i -lt $filtered.Count; $i++) { $wifFiltered.Add($filtered[$i]) }
    $baseCount = $wifFiltered.Count

    foreach ($spec in $wifSpecs) {
        # Parse: "<metric-shorthand> <target-col-or-range> <value>"
        $parts = $spec -split '\s+', 3
        if ($parts.Count -lt 3) { Write-Warning "Skipping malformed WIF spec: $spec"; continue }
        $specName = $parts[0] + ' ' + $parts[1]  # e.g. "PPVs ETT"
        # For ENG the 3rd token might be "QS 25" -> split further
        $remainder = $parts[2]
        $remParts = $remainder -split '\s+', 2
        $targetName = $remParts[0]
        $wifValue = if ($remParts.Count -gt 1) { $remParts[1] } else { $remParts[0] }

        # If specName has 2 words and remainder has 2 words, we have: spec(2) target(1) value(1)
        # e.g. "PPVs ETT QS 25" -> specName="PPVs ETT", targetName="QS", wifValue="25"

        # Resolve target column
        $targetCol = [array]::IndexOf([string[]]$hdr, $targetName)
        if ($targetCol -lt 0) {
            Write-Warning "WIF target column '$targetName' not found in headers"
            continue
        }

        # Resolve matching rows based on spec shorthand
        $specUpper = $specName.ToUpper()
        $matchIdx = New-Object System.Collections.Generic.List[int]

        for ($i = 1; $i -lt $filtered.Count; $i++) {
            $row = $filtered[$i]
            $op2  = "$($row[$opIdx])".Trim()
            $met2 = if ($metIdx -ge 0) { "$($row[$metIdx])".Trim() } else { '' }

            $matched = switch -Wildcard ($specUpper) {
                'PPVS ETT'     { $op2 -eq 'PPV_SPS' -and $met2 -eq 'TEST TIME - MIN' }
                'PPVM ETT'     { $op2 -eq 'PPV_SPM' -and $met2 -eq 'TEST TIME - MIN' }
                'PPVS RCS'     { $op2 -eq 'PPV_SPS' -and $met2 -like 'RETEST RATE*' }
                'PPVM RCS'     { $op2 -eq 'PPV_SPM' -and $met2 -like 'RETEST RATE*' }
                'CLASSHOT ETT' { $op2 -eq 'TEST_MPS' -and $met2 -eq 'TEST_TIME-SEC' }
                'CLASSHOT RCS' { $op2 -eq 'TEST_MPS' -and $met2 -like 'RETEST RATE*' }
                default        { $false }
            }
            if ($matched) { $matchIdx.Add($i) }
        }

        Write-Host "WIF spec '$specName' target='$targetName' value='$wifValue': $($matchIdx.Count) matches"

        foreach ($mi in $matchIdx) {
            # Add POR row
            $wifFiltered.Add($filtered[$mi])
            # Add WIF clone
            $clone = @($filtered[$mi] | ForEach-Object { $_ })
            if ($porIdx -ge 0) { $clone[$porIdx] = 'WIF' }
            $clone[$targetCol] = $wifValue
            $wifFiltered.Add($clone)
            $excelRow = $wifFiltered.Count
            $rowRedCols[$excelRow] = @($targetCol + 1)
        }
    }

    $wifRows = $wifFiltered
    Write-Host "WIF tab: $($wifFiltered.Count - 1) rows (baseline + POR + WIF clones)"
}

# ============================================================================
# Build xlsx
# ============================================================================
$fileSuffix = if ($filterLabel) { "_$filterLabel" } else { '' }
$wifSuffix  = if ($wifApplied) { '_WIF' } else { '' }
$csvPath    = Join-Path $env:TEMP "${safeName}_${Report}${fileSuffix}${wifSuffix}.csv"
$xlsxPath   = Join-Path $env:TEMP "${safeName}_${Report}${fileSuffix}${wifSuffix}.xlsx"

if ($wifApplied) {
    # 2-tab workbook: Tab1=Baseline, Tab2=WIF
    $csvBase = Join-Path $env:TEMP "${safeName}_${Report}${fileSuffix}_base.csv"
    $csvWif  = Join-Path $env:TEMP "${safeName}_${Report}${fileSuffix}_wif.csv"

    Tci-WriteCsv -Rows $filtered -Path $csvBase
    Tci-WriteCsv -Rows $wifRows -Path $csvWif

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    try {
        $targetWb = $excel.Workbooks.Add()
        $targetWb.Worksheets.Item(1).Name = '_placeholder_'

        # Tab 1
        $wbB = $excel.Workbooks.Open($csvBase)
        $wsB = $wbB.Worksheets.Item(1)
        $tab1Name = "$filterLabel Baseline"
        if ($tab1Name.Length -gt 31) { $tab1Name = $tab1Name.Substring(0,31) }
        $wsB.Name = $tab1Name
        $wsB.Move([System.Reflection.Missing]::Value, $targetWb.Worksheets.Item($targetWb.Worksheets.Count))
        try { $wbB.Close($false) } catch {}

        # Tab 2
        $wbW = $excel.Workbooks.Open($csvWif)
        $wsW = $wbW.Worksheets.Item(1)
        $wsW.Name = 'WIF'
        $wsW.Move([System.Reflection.Missing]::Value, $targetWb.Worksheets.Item($targetWb.Worksheets.Count))
        try { $wbW.Close($false) } catch {}

        $targetWb.Worksheets.Item('_placeholder_').Delete()

        # Format Tab 1
        $ws1 = $targetWb.Worksheets.Item($tab1Name)
        $ws1.Activate()
        $lc1 = $ws1.UsedRange.Columns.Count
        $ws1.Range($ws1.Cells(1,1), $ws1.Cells(1,$lc1)).Font.Bold = $true
        $ws1.Range($ws1.Cells(1,1), $ws1.Cells(1,$lc1)).Interior.Color = 14277081
        try { $aw=$excel.ActiveWindow; $aw.SplitColumn=4; $aw.SplitRow=1; $aw.FreezePanes=$true } catch {}
        $ws1.Columns.AutoFit() | Out-Null

        # Format Tab 2
        $ws2 = $targetWb.Worksheets.Item('WIF')
        $ws2.Activate()
        $lc2 = $ws2.UsedRange.Columns.Count
        $ws2.Range($ws2.Cells(1,1), $ws2.Cells(1,$lc2)).Font.Bold = $true
        $ws2.Range($ws2.Cells(1,1), $ws2.Cells(1,$lc2)).Interior.Color = 14277081
        try { $aw2=$excel.ActiveWindow; $aw2.SplitColumn=4; $aw2.SplitRow=1; $aw2.FreezePanes=$true } catch {}

        # WIF color coding
        $porColExcel = if ($porIdx -ge 0) { $porIdx + 1 } else { 0 }
        $totalR = $wifRows.Count
        for ($r = 2; $r -le $totalR; $r++) {
            $isWif = $false
            if ($porColExcel -gt 0) { $isWif = ($ws2.Cells($r, $porColExcel).Text -eq 'WIF') }
            if ($isWif) {
                $ws2.Range($ws2.Cells($r,1), $ws2.Cells($r,$lc2)).Interior.Color = 0x66FFFF
                if ($rowRedCols.ContainsKey($r)) {
                    foreach ($col in $rowRedCols[$r]) { $ws2.Cells($r, $col).Font.Color = 0x0000FF }
                }
            }
        }
        $ws2.Columns.AutoFit() | Out-Null

        $targetWb.SaveAs($xlsxPath, 51)
        Write-Host "xlsx: $xlsxPath ($([Math]::Round((Get-Item $xlsxPath).Length/1KB, 1)) KB)"
    } finally {
        try { $targetWb.Close($false) } catch {}
        try { $excel.Quit() } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Remove-Item $csvBase -Force -ErrorAction SilentlyContinue
        Remove-Item $csvWif -Force -ErrorAction SilentlyContinue
    }
} else {
    # Single-tab workbook
    Tci-WriteCsv -Rows $filtered -Path $csvPath
    $result = Tci-ExportXlsx -CsvPath $csvPath -XlsxPath $xlsxPath -SheetName $sheetName -Validate -MinRows $rowCount -MinCols 10
    Write-Host "xlsx: $($result.Path) ($([Math]::Round((Get-Item $xlsxPath).Length/1KB, 1)) KB) rows=$($result.Rows) cols=$($result.Cols)"
    Remove-Item $csvPath -Force -ErrorAction SilentlyContinue
}

# ============================================================================
# Email
# ============================================================================
if (-not $NoEmail) {
    $sections = @(@{ SheetTag = $reportLabel; RowCount = $rowCount })
    $wifLabel = if ($wifApplied) { " + WIF" } else { '' }
    if ($wifApplied) {
        $body = Build-WifCard -Product $Product -Group $grp -SubGroup $sub -Report $reportLabel -Operation $filterLabel -Metric $Wif -WifValue $Wif -TimeRange '' -PorRows $rowCount -WifRows ($wifRows.Count - 1 - $rowCount)
    } else {
        $body = Build-PhiCard -Product $Product -Group $grp -SubGroup $sub -Sections $sections -Filter $filterLabel
    }
    $subj = "PHI of $Product - $reportLabel $filterLabel$wifLabel"
    Tci-SendMail -Subject $subj -HtmlBody $body -Attachments @($xlsxPath)
}

Write-Host "`nDone." -ForegroundColor Green
