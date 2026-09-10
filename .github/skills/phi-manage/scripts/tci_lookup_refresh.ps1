$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\tci_lib.ps1"

# This script adds newly-discovered MOR CommonNames to the lookup JSON
# by inferring AT_Group/AT_SubGroup from naming patterns.

$lookupPath = Join-Path $PSScriptRoot 'tci_commonname_lookup.json'
$j = Get-Content $lookupPath -Raw | ConvertFrom-Json
$known = $j.Lookup | ForEach-Object { $_.CommonName }
$noData = $j.NoDataCommonNames

# New products found in MOR but not in lookup
$newProducts = @(
    'Arcadian Shores Module'
    'Arcadian Shores PCIe'
    'Arcadian Shores x8'
    'Coral Rapids-HCC'
    'Coral Rapids-LCC'
    'Coral Rapids-UCC'
    'Coral Rapids-XCC'
    'Coral Rapids-ZCC'
    'DIAMOND RAPIDS-UCC-X1'
    'Explorer Island'
    'Hammer Lake HL'
    'Hammer Lake HM'
    'Hammer Lake HX'
    'Hammer Lake SB'
    'Hammer Lake SK'
    'Hammer Lake SL'
    'Hammer Lake SM'
    'Hammer Lake SMD'
    'Hammer Lake SML'
    'Hammer Lake SX'
    'Hope Island'
    'Jaguar Shores 2T+1'
    'Morganville (CNIC)'
    'MOUNT EDEN'
    'Newport'
    'Panther Lake U (NEX) 4P+0E+4LP_E E-Temp'
    'Serpent Lake B'
    'Serpent Lake BX'
    'Serpent Lake HL'
    'Serpent Lake HM'
    'Serpent Lake HPX'
    'Tiger Shores'
    'Tiger Shores x8 Air Cooled'
    'Tiger Shores x8 Liquid Cooled'
    'Titan Lake B'
    'Titan Lake BX'
    'Titan Lake HL'
    'Titan Lake HL [FOR LRP USE ONLY]'
    'Titan Lake HM'
    'Titan Lake HPX'
    'WIF Coral Rapids-RS'
    'WIF Iron Rapids-HCC'
    'WIF Iron Rapids-LCC'
    'WIF Iron Rapids-RS'
    'WIF Iron Rapids-UCC'
    'WIF Nova Lake AX 16C'
    'WIF Nova Lake H Int'
    'WIF Nova Lake S 16C 12Xe'
)

# Heuristic group/subgroup resolution
function Infer-GroupSubGroup($name) {
    $n = $name.ToUpper()

    # Server indicators
    if ($n -match '(RAPIDS|SHORES|DIAMOND|IRON|ARCADIAN|TIGER|JAGUAR)') {
        # Server-class products (XCC/HCC/LCC/UCC/ZCC = Server core configs)
        if ($n -match '(HCC|LCC|UCC|XCC|ZCC|RS)') { return @('Server','Server') }
        if ($n -match 'MODULE|PCIE|X8') { return @('Server','Server') }
        return @('Server','Server')
    }
    if ($n -match 'MOUNT EDEN|MORGANVILLE|NEWPORT|HOPE ISLAND|EXPLORER ISLAND') {
        return @('Server','Server')  # Infrastructure/platform codenames
    }

    # WIF entries - match the underlying product
    if ($n -match '^WIF ') {
        $base = $name.Substring(4)
        if ($base -match '(?i)Nova Lake.*S') { return @('Client','Desktop') }
        if ($base -match '(?i)Nova Lake.*(AX|H|U)') { return @('Client','Mobile') }
        if ($base -match '(?i)(Coral|Iron).*Rapids') { return @('Server','Server') }
        return @('Client','Mobile')
    }

    # Hammer Lake / Serpent Lake / Titan Lake - client
    if ($n -match 'HAMMER LAKE|SERPENT LAKE|TITAN LAKE') {
        # H/HX/HL/HM/HPX = Mobile; S/SB/SK/SL/SM/SX = Desktop; B/BX = Desktop
        if ($n -match '\s(S[A-Z]*|B[A-Z]*)$') { return @('Client','Desktop') }
        if ($n -match '\s(H[A-Z]*)$') { return @('Client','Mobile') }
        if ($n -match '\sB$') { return @('Client','Desktop') }
        return @('Client','Mobile')
    }

    # Panther Lake E-Temp
    if ($n -match 'PANTHER LAKE.*U') { return @('Client','Mobile') }

    return @('Client','Mobile')  # default fallback
}

$added = 0
foreach ($name in $newProducts) {
    if ($known -contains $name) { continue }
    if ($noData -contains $name) { continue }

    $gs = Infer-GroupSubGroup $name
    $entry = [PSCustomObject]@{
        CommonName  = $name
        AT_Group    = $gs[0]
        AT_SubGroup = $gs[1]
        Source      = 'MOR_inferred_2026-06-01'
    }
    $j.Lookup += $entry
    $added++
    Write-Host "  + $name -> $($gs[0]) / $($gs[1])"
}

Write-Host "`nAdded $added new entries. Total lookup: $($j.Lookup.Count)"

# Save
$j | ConvertTo-Json -Depth 5 | Set-Content $lookupPath -Encoding UTF8
Write-Host "Saved: $lookupPath"
