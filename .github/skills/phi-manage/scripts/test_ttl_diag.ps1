$h = Get-Content "$env:TEMP\tci_after_Nova_Lake_AX_28C_sdamo.html" -Raw
if ($h -match 'No Data|noData|no records') { Write-Host "RESPONSE SAYS: No Data" }
else { Write-Host "Has content: $($h.Length) chars" }
$ths = [regex]::Matches($h, '<th[^>]*>([^<]+)</th>') | Select-Object -First 20 | ForEach-Object { $_.Groups[1].Value }
Write-Host "First 20 th elements:"
$ths
