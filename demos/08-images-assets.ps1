#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 8 — Local images: generated PNGs are cataloged and searched, placed with contain and
    cover (focal point), replaced in their frame, described for accessibility, and inspected for
    resolution and distortion. Nothing is downloaded.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '08-images-assets' -OutputRoot $OutputRoot
$s = $demo.session
$folder = Join-Path $demo.folder 'images'
New-Item -ItemType Directory -Force -Path $folder | Out-Null
New-DemoImage -Path (Join-Path $folder 'warehouse-wide.png') -Width 2400 -Height 1000 -Label 'Warehouse'
New-DemoImage -Path (Join-Path $folder 'team-portrait.png') -Width 900 -Height 1350 -Label 'Team' -From '#7B2D26' -To '#F4A261'
New-DemoImage -Path (Join-Path $folder 'logo-square.png') -Width 600 -Height 600 -Label 'Logo' -From '#00655B' -To '#2A9D8F'
New-DemoImage -Path (Join-Path $folder 'thumbnail-small.png') -Width 240 -Height 135 -Label 'Small'
Copy-Item (Join-Path $folder 'logo-square.png') (Join-Path $folder 'logo-copy.png')

# The catalog lives in the daemon's PPTMCP_ASSETS_DIR (default %LOCALAPPDATA%\PowerPointMcp\assets).
$scan = Invoke-Pptcli -Step 'asset scan-folder' -Arguments @('asset', 'scan-folder', '-s', $s, '--folder', $folder)
Invoke-Pptcli -Step 'asset tag-asset' -Arguments @('asset', 'tag-asset', '-s', $s, '--path', (Join-Path $folder 'team-portrait.png'), '--tags', 'people,team', '--description', 'The project team at the 2025 offsite', '--attribution', 'Photo: internal') | Out-Null
$search = Invoke-Pptcli -Step 'asset search portrait team' -Arguments @('asset', 'search', '-s', $s, '--query', 'team', '--orientation', 'portrait')
$dups = Invoke-Pptcli -Step 'asset duplicates' -Arguments @('asset', 'duplicates', '-s', $s)

Invoke-Pptcli -Step 'asset place contain' -Arguments @('asset', 'place', '-s', $s, '--slide-index', '1', '--path', (Join-Path $folder 'warehouse-wide.png'), '--left', '40', '--top', '60', '--width', '420', '--height', '300', '--fit', 'contain', '--alt-text', 'Warehouse exterior', '--name', 'Warehouse') | Out-Null
Invoke-Pptcli -Step 'asset place cover (focal left)' -Arguments @('asset', 'place', '-s', $s, '--slide-index', '1', '--path', (Join-Path $folder 'warehouse-wide.png'), '--left', '500', '--top', '60', '--width', '300', '--height', '300', '--fit', 'cover', '--focal-x', '0.2', '--name', 'Hero') | Out-Null
Invoke-Pptcli -Step 'asset place small (low PPI)' -Arguments @('asset', 'place', '-s', $s, '--slide-index', '1', '--path', (Join-Path $folder 'thumbnail-small.png'), '--left', '40', '--top', '380', '--width', '420', '--height', '140', '--name', 'Thumb') | Out-Null
Invoke-Pptcli -Step 'asset replace (keep frame)' -Arguments @('asset', 'replace', '-s', $s, '--target', 'name:Hero', '--path', (Join-Path $folder 'team-portrait.png'), '--fit', 'cover', '--focal-y', '0.3') | Out-Null
Invoke-Pptcli -Step 'asset set-alt-text decorative' -Arguments @('asset', 'set-alt-text', '-s', $s, '--target', 'name:Thumb', '--decorative', 'true') | Out-Null
$imageSpec = @{
    kind   = 'image-text'
    id     = 'logo'
    title  = 'Our brand'
    image  = @{ path = (Join-Path $folder 'logo-square.png'); alt = 'Company logo'; fit = 'contain'; position = 'right' }
    points = @('Teal is our primary colour', 'Square logo on light backgrounds')
} | ConvertTo-Json -Compress -Depth 5
Invoke-Pptcli -Step 'compose image-text' -Arguments @('compose', 'create', '-s', $s, '--spec', $imageSpec) | Out-Null
$inspect = Invoke-Pptcli -Step 'asset inspect' -Arguments @('asset', 'inspect', '-s', $s)

$pictures = ($inspect.pictures | ForEach-Object { "$($_.name) ppi=$($_.effectivePpi) issues=$(@($_.issues) -join '/')" }) -join '; '
Stop-Demo -Demo $demo -Notes @("Catalog: added $($scan.scan.added); search hits: $(@($search.assets).Count); duplicate groups: $(@($dups.duplicateGroups).Count)", "Pictures: $pictures")
