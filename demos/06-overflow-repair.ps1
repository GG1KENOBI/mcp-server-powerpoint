#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 6 — Overflow and repair: a long summary splits onto continuation slides with measured
    fitting; a hand-made overflowing box, an off-slide shape, and misaligned titles are found by
    review validate, planned, and repaired safely.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '06-overflow-repair' -OutputRoot $OutputRoot
$s = $demo.session
$plan = Invoke-Pptcli -Step 'compose plan (long summary)' -Arguments @('compose', 'plan', '-s', $s, '--spec-path', (Join-Path $PSScriptRoot 'specs/overflow-summary.json'))
$created = Invoke-Pptcli -Step 'compose create (long summary)' -Arguments @('compose', 'create', '-s', $s, '--spec-path', (Join-Path $PSScriptRoot 'specs/overflow-summary.json'))

$long = 'This box was sized by hand and holds far more text than fits. ' * 12
Invoke-Pptcli -Step 'shape add-text-box (overflowing)' -Arguments @('shape', 'add-text-box', '-s', $s, '--slide-index', '1', '--left', '60', '--top', '420', '--width', '300', '--height', '40', '--text', $long) | Out-Null
Invoke-Pptcli -Step 'shape add-text-box (off-slide)' -Arguments @('shape', 'add-text-box', '-s', $s, '--slide-index', '1', '--left', '900', '--top', '480', '--width', '200', '--height', '40', '--text', 'Partly off the slide') | Out-Null

$before = Invoke-Pptcli -Step 'review validate (before)' -Arguments @('review', 'validate', '-s', $s, '--limit', '200')
$repairPlan = Invoke-Pptcli -Step 'review plan-repair' -Arguments @('review', 'plan-repair', '-s', $s)
Invoke-Pptcli -Step 'review repair' -Arguments @('review', 'repair', '-s', $s) | Out-Null
$after = Invoke-Pptcli -Step 'review validate (after)' -Arguments @('review', 'validate', '-s', $s, '--limit', '200')

Stop-Demo -Demo $demo -Notes @("Composition slides: $(@($created.slides).Count) (planned $(@($plan.planned).Count))", "Findings before repair: $($before.totalFindings); after: $($after.totalFindings)")
