#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 7 — Deck-wide renaming: a product name in titles, bullets, a table, a group, and speaker
    notes is found and replaced with a preview and an exact-count guard; other formatting stays.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '07-deck-text-editing' -OutputRoot $OutputRoot
$s = $demo.session
Invoke-Pptcli -Step 'compose summary' -Arguments @('compose', 'create', '-s', $s, '--spec', '{"kind":"executive-summary","id":"sum","title":"Contoso Cloud in 2025","points":["Contoso Cloud reached 1,200 customers","Contoso Cloud Edge launched in May","Partners resell Contoso Cloud in 14 countries"]}') | Out-Null
Invoke-Pptcli -Step 'compose table' -Arguments @('compose', 'create', '-s', $s, '--spec', '{"kind":"table","id":"plans","title":"Contoso Cloud plans","table":{"header":["Plan","Price"],"rows":[["Contoso Cloud Basic","$10"],["Contoso Cloud Pro","$25"]]}}') | Out-Null
Invoke-Pptcli -Step 'notes set-notes-text' -Arguments @('notes', 'set-notes-text', '-s', $s, '--slide-index', '2', '--text', 'Mention that Contoso Cloud pricing changes in January.') | Out-Null

$all = Invoke-Pptcli -Step 'deck find-text (with notes)' -Arguments @('deck', 'find-text', '-s', $s, '--find-what', 'Contoso Cloud', '--include-notes', 'true')
$tables = Invoke-Pptcli -Step 'deck find-text (tables only)' -Arguments @('deck', 'find-text', '-s', $s, '--find-what', 'Contoso Cloud', '--selector', 'kind:table')
$guard = Invoke-Pptcli -Step 'deck replace-text (wrong count refused)' -Arguments @('deck', 'replace-text', '-s', $s, '--find-what', 'Contoso Cloud', '--replace-what', 'Fabrikam One', '--include-notes', 'true', '--dry-run', 'false', '--expected-count', '1') -AllowFailure
$done = Invoke-Pptcli -Step 'deck replace-text' -Arguments @('deck', 'replace-text', '-s', $s, '--find-what', 'Contoso Cloud', '--replace-what', 'Fabrikam One', '--include-notes', 'true', '--dry-run', 'false', '--expected-count', "$($all.totalCount)")
$regex = Invoke-Pptcli -Step 'deck replace-text (regex prices)' -Arguments @('deck', 'replace-text', '-s', $s, '--find-what', '\$(\d+)', '--replace-what', '$1 USD', '--use-regex', 'true', '--dry-run', 'false')
$left = Invoke-Pptcli -Step 'deck find-text (left over)' -Arguments @('deck', 'find-text', '-s', $s, '--find-what', 'Contoso', '--include-notes', 'true')

Stop-Demo -Demo $demo -Notes @("Matches: $($all.totalCount) (tables: $($tables.totalCount)); wrong expected_count refused: $(-not $guard.success)", "Replaced: $($done.replacementCount); prices rewritten: $($regex.replacementCount); left over: $($left.totalCount)")
