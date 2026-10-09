#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 2 — Corporate restyle: a deck built with the default profile moves to a custom profile
    (demos/profiles/acme.json) with a migration preview, then typography normalization and component updates.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '02-restyle-profile' -OutputRoot $OutputRoot
$s = $demo.session

foreach ($spec in @(
    '{"kind":"cards","id":"prio","title":"Our priorities","cards":[{"heading":"Growth","body":"Enter two new markets","badge":"P1"},{"heading":"Efficiency","body":"Automate reporting"},{"heading":"People","body":"Hire 40 engineers"}]}',
    '{"kind":"kpis","id":"glance","title":"Q3 at a glance","kpis":[{"value":"$4.2M","label":"Revenue","delta":"+12% YoY","trend":"up"},{"value":"38%","label":"Gross margin","delta":"+3 pts","trend":"up"}]}',
    '{"kind":"table","id":"regions","title":"Revenue by region","table":{"header":["Region","2024","2025"],"rows":[["EMEA","3.1","3.6"],["APAC","2.4","2.9"]]}}',
    '{"kind":"comparison","id":"choice","title":"Build or buy","columns":[{"heading":"Build","points":["Full control","12 months"]},{"heading":"Buy","points":["Live in 8 weeks","Licence cost"],"tone":"positive"}]}')) {
    Invoke-Pptcli -Step "compose create $(($spec | ConvertFrom-Json).id)" -Arguments @('compose', 'create', '-s', $s, '--spec', $spec, '--profile', 'default') | Out-Null
}

$profileJson = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'profiles/acme.json') -Raw
Invoke-Pptcli -Step 'design validate-profile' -Arguments @('design', 'validate-profile', '-s', $s, '--profile-json', $profileJson) | Out-Null
# Saved to the daemon's profile folder (PPTMCP_PROFILES_DIR or %LOCALAPPDATA%\PowerPointMcp\profiles).
Invoke-Pptcli -Step 'design import-profile' -Arguments @('design', 'import-profile', '-s', $s, '--path', (Join-Path $PSScriptRoot 'profiles/acme.json'), '--overwrite', 'true') | Out-Null
$plan = Invoke-Pptcli -Step 'design apply-profile (preview)' -Arguments @('design', 'apply-profile', '-s', $s, '--profile', 'acme', '--from-profile', 'default')
$applied = Invoke-Pptcli -Step 'design apply-profile' -Arguments @('design', 'apply-profile', '-s', $s, '--profile', 'acme', '--from-profile', 'default', '--dry-run', 'false')
$typo = Invoke-Pptcli -Step 'design normalize-typography (preview)' -Arguments @('design', 'normalize-typography', '-s', $s, '--profile', 'acme')
Invoke-Pptcli -Step 'design normalize-typography' -Arguments @('design', 'normalize-typography', '-s', $s, '--profile', 'acme', '--dry-run', 'false') | Out-Null
Invoke-Pptcli -Step 'design update-components' -Arguments @('design', 'update-components', '-s', $s, '--profile', 'acme', '--from-profile', 'default', '--dry-run', 'false') | Out-Null
Invoke-Pptcli -Step 'design list-components' -Arguments @('design', 'list-components', '-s', $s, '--profile', 'acme') | Out-Null
$extracted = Invoke-Pptcli -Step 'design extract-profile' -Arguments @('design', 'extract-profile', '-s', $s, '--name', 'from-demo-deck')
$extracted.profileJson | Set-Content -LiteralPath (Join-Path $demo.folder 'extracted-profile.json') -Encoding utf8

Stop-Demo -Demo $demo -Notes @("Planned changes: $($plan.changeCount) ($(($plan.changesByProperty.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ', '))", "Typography changes planned: $($typo.changeCount)")
