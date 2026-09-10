$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

$init = Tci-GetInit -ReportId 'PORSDAForecast' -MaxAgeMinutes 60
# Try CommonName panel first
$cnPanel = [regex]::Match($init, '(?is)id="ContentPlaceHolder1_Filters_CommonName"(.*?)</div>\s*</div>').Value
$labels = [regex]::Matches($cnPanel, '(?is)<label[^>]*>([^<]+)</label>') | ForEach-Object { $_.Groups[1].Value.Trim() }
$ttl = $labels | Where-Object { $_ -match '(?i)^NVL|Nova\s*Lake|Novalake' } | Sort-Object
if ($ttl) { $ttl }
else {
    # Broad search across all labels
    Write-Host "Not in CommonName panel. Broad search:"
    $all = [regex]::Matches($init, '<label[^>]*>([^<]+)</label>') | ForEach-Object { $_.Groups[1].Value.Trim() }
    $hits = $all | Where-Object { $_ -match '(?i)^NVL|Nova\s*Lake|Novalake' } | Sort-Object -Unique
    if ($hits) { $hits } else { Write-Host "No NVL/Nova Lake found anywhere in init HTML" }
}
