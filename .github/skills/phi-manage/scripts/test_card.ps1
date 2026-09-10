. "$PSScriptRoot\tci_lib.ps1"
$sections = @( @{ SheetTag = 'SDA Weekly Class'; RowCount = 157 } )
$body = Build-PhiCard -Product 'ARL U 2C+8A+GT1' -Group 'Client' -SubGroup 'Mobile' -Sections $sections -Filter 'Class=TEST'
Set-Content "$env:TEMP\card_test.html" $body -Encoding ASCII
Write-Host "Font declarations found:"
Select-String 'font-family' "$env:TEMP\card_test.html" | ForEach-Object { Write-Host "  $($_.Line.Substring(0, [Math]::Min(150, $_.Line.Length)))" }
Write-Host "`nFull HTML saved to $env:TEMP\card_test.html"
