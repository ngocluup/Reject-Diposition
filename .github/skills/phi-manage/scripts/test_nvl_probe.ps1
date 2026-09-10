$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

$product  = 'Nova Lake AX 28C'
$reportId = 'PORMonthlyForecast'
$initHtml = Tci-GetInit -ReportId $reportId -MaxAgeMinutes 60

$vs  = ([regex]::Match($initHtml, 'name="__VIEWSTATE"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value
$vsg = ([regex]::Match($initHtml, 'name="__VIEWSTATEGENERATOR"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value
$ev  = ([regex]::Match($initHtml, 'name="__EVENTVALIDATION"\s+id="[^"]+"\s+value="([^"]*)"')).Groups[1].Value

# PHI Parameters
$phiPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_PHIParameters"(.*?)</div>\s*</div>').Value
$phiCtls  = [regex]::Matches($phiPanel, 'name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"') | ForEach-Object { $_.Groups[1].Value }

# CommonName ctl
$cnPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_CommonName"(.*?)</div>\s*</div>').Value
$escaped = [regex]::Escape($product)
$cnCtl   = ([regex]::Matches($cnPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>') | Where-Object { $_.Groups[2].Value.Trim() -match "^${escaped}" }).Groups[1].Value
Write-Host "CN ctl: $cnCtl"

# Get all groups and subgroups
$grpPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_Group"(.*?)</div>\s*</div>').Value
$grpItems = [regex]::Matches($grpPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>')

$subPanel = [regex]::Match($initHtml, '(?is)id="ContentPlaceHolder1_Filters_AT_SubGroup"(.*?)</div>\s*</div>').Value
$subItems = [regex]::Matches($subPanel, '(?is)<input[^>]+name="(ctl00\$ContentPlaceHolder1\$ctl\d+)"[^>]*/>\s*<label[^>]*>([^<]+)</label>')

Write-Host "Groups: $($grpItems.Count)  SubGroups: $($subItems.Count)"

# Try each group+subgroup combo
foreach ($g in $grpItems) {
    $gCtl = $g.Groups[1].Value; $gName = $g.Groups[2].Value.Trim()
    foreach ($s in $subItems) {
        $sCtl = $s.Groups[1].Value; $sName = $s.Groups[2].Value.Trim()
        $form = [ordered]@{
            '__EVENTTARGET'=''; '__EVENTARGUMENT'=''
            '__VIEWSTATE'=$vs; '__VIEWSTATEGENERATOR'=$vsg; '__EVENTVALIDATION'=$ev
            $gCtl='on'; $sCtl='on'; $cnCtl='on'
            'ctl00$ContentPlaceHolder1$btn_RunReport'='Run Report'
        }
        foreach ($p in $phiCtls) { $form[$p] = 'on' }
        Write-Host -NoNewline "  $gName / $sName ... "
        try {
            $resp = Tci-Post -ReportId $reportId -Form $form -TimeoutSec 120 -Retries 0
            $content = $resp.Content
            if ($content -match 'No Data|noData') { Write-Host "NoData" }
            else {
                $tables = Tci-ParseTables $content
                $dt = Tci-PickDataTable $tables 'CommonName'
                if ($dt) {
                    $parsed = Tci-RowsFromTable $dt
                    Write-Host "OK - $($parsed.Rows.Count - 1) rows" -ForegroundColor Green
                } else { Write-Host "NoTable" }
            }
        } catch { Write-Host "Error: $($_.Exception.Message)" }
    }
}
