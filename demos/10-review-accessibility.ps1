#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 10 — Review and accessibility: low-contrast text, tiny font, a picture without alt
    text, and inconsistent titles are detected with evidence; the deck is rendered (snapshot,
    contact sheet) and a before/after image diff measures a change.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '10-review-accessibility' -OutputRoot $OutputRoot
$s = $demo.session
Invoke-Pptcli -Step 'compose cards' -Arguments @('compose', 'create', '-s', $s, '--spec', '{"kind":"cards","id":"c","title":"Review me","cards":[{"heading":"Clear","body":"Readable text"},{"heading":"Busy","body":"More text"}]}') | Out-Null
$low = Invoke-Pptcli -Step 'shape add-text-box (low contrast)' -Arguments @('shape', 'add-text-box', '-s', $s, '--slide-index', '1', '--left', '60', '--top', '470', '--width', '400', '--height', '30', '--text', 'Pale grey footnote on white')
Invoke-Pptcli -Step 'textframe format-range (pale, 7 pt)' -Arguments @('textframe', 'format-range', '-s', $s, '--slide-index', '1', '--shape-index', "$($low.shapeIndex)", '--color', '#DDDDDD', '--font-size', '7') | Out-Null
$image = Join-Path $demo.folder 'chart.png'
New-DemoImage -Path $image -Width 1200 -Height 800 -Label 'Chart'
Invoke-Pptcli -Step 'asset place (no alt text)' -Arguments @('asset', 'place', '-s', $s, '--slide-index', '1', '--path', $image, '--left', '520', '--top', '360', '--width', '200', '--height', '140', '--name', 'Unlabelled chart') | Out-Null

$rules = Invoke-Pptcli -Step 'review list-rules' -Arguments @('review', 'list-rules', '-s', $s)
$findings = Invoke-Pptcli -Step 'review validate' -Arguments @('review', 'validate', '-s', $s, '--limit', '200')
$audit = Invoke-Pptcli -Step 'accessibility audit' -Arguments @('accessibility', 'audit', '-s', $s)
$before = Invoke-Pptcli -Step 'preview snapshot (before)' -Arguments @('preview', 'snapshot', '-s', $s, '--slide-index', '1', '--output-path', (Join-Path $demo.folder 'before.png'), '--overwrite', 'true', '--include-image', 'false')
Invoke-Pptcli -Step 'asset set-alt-text' -Arguments @('asset', 'set-alt-text', '-s', $s, '--target', 'name:"Unlabelled chart"', '--alt-text', 'Column chart of quarterly revenue') | Out-Null
Invoke-Pptcli -Step 'textframe format-range (fix contrast)' -Arguments @('textframe', 'format-range', '-s', $s, '--slide-index', '1', '--shape-index', "$($low.shapeIndex)", '--color', '#404040', '--font-size', '12') | Out-Null
$after = Invoke-Pptcli -Step 'preview snapshot (after)' -Arguments @('preview', 'snapshot', '-s', $s, '--slide-index', '1', '--output-path', (Join-Path $demo.folder 'after.png'), '--overwrite', 'true', '--include-image', 'false')
$diff = Invoke-Pptcli -Step 'preview compare-images' -Arguments @('preview', 'compare-images', '-s', $s, '--first-path', $before.imagePath, '--second-path', $after.imagePath)
$recheck = Invoke-Pptcli -Step 'review validate (after fixes)' -Arguments @('review', 'validate', '-s', $s, '--limit', '200')

Stop-Demo -Demo $demo -Notes @("Rules: $(@($rules.rules).Count); findings before: $($findings.totalFindings) ($((@($findings.findings) | Group-Object code | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', ')); after: $($recheck.totalFindings)", "Changed pixels: $($diff.difference.changedRatio)")
