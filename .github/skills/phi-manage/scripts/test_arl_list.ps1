. "$PSScriptRoot\tci_lib.ps1"
$init = Tci-GetInit -ReportId 'MORSpread'
# Extract all label text from CommonName filter panel
$panelMatch = [regex]::Match($init, '(?s)Filters_CommonName.*?</fieldset>')
if ($panelMatch.Success) {
    $labels = [regex]::Matches($panelMatch.Value, '<label[^>]*>([^<]+)</label>') | ForEach-Object { $_.Groups[1].Value }
    $arl = $labels | Where-Object { $_ -match 'ARL|Arrow Lake' } | Sort-Object
    Write-Host "=== ARL products in TCI ($($arl.Count) total) ==="
    $i = 1
    foreach ($p in $arl) { Write-Host "  $i. $p"; $i++ }
} else {
    Write-Host "Panel not found, trying broad search..."
    $labels = [regex]::Matches($init, '<label[^>]*>([^<]*(?:ARL|Arrow)[^<]*)</label>') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
    foreach ($p in $labels) { Write-Host "  $p" }
}
