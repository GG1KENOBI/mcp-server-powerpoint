#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 1 — Quarterly business review from a Russian CSV (decimal commas, ₽): title, KPI and
    summary compositions, a bound chart and table, then a data refresh after the file changes.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '01-quarterly-review' -OutputRoot $OutputRoot
$s = $demo.session
$csv = Join-Path $demo.folder 'quarterly-ru.csv'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'data/quarterly-ru.csv') -Destination $csv -Force

Invoke-Pptcli -Step 'compose create title' -Arguments @('compose', 'create', '-s', $s, '--spec-path', (Join-Path $PSScriptRoot 'specs/qbr-title.json'), '--profile', 'default') | Out-Null
Invoke-Pptcli -Step 'compose create kpis' -Arguments @('compose', 'create', '-s', $s, '--spec-path', (Join-Path $PSScriptRoot 'specs/qbr-kpis.json'), '--profile', 'default') | Out-Null
Invoke-Pptcli -Step 'compose create summary' -Arguments @('compose', 'create', '-s', $s, '--spec-path', (Join-Path $PSScriptRoot 'specs/qbr-summary.json'), '--profile', 'default') | Out-Null
$preview = Invoke-Pptcli -Step 'data preview' -Arguments @('data', 'preview', '-s', $s, '--source-path', $csv, '--culture', 'ru-RU')
Invoke-Pptcli -Step 'data create-chart' -Arguments @('data', 'create-chart', '-s', $s, '--source-path', $csv, '--culture', 'ru-RU', '--title', 'Выручка и расходы по кварталам', '--chart-type', 'column', '--takeaway', 'Выручка растёт быстрее расходов', '--slide-app-id', 'qbr-chart') | Out-Null
Invoke-Pptcli -Step 'data create-table' -Arguments @('data', 'create-table', '-s', $s, '--source-path', $csv, '--culture', 'ru-RU', '--title', 'Показатели по кварталам', '--source-note', 'Управленческая отчётность', '--slide-app-id', 'qbr-table') | Out-Null

# The source file changes: Q4 forecast is added; refresh updates the bound chart and table in place.
Add-Content -LiteralPath $csv -Value 'Q4 2025 (прогноз);1 610,0;1 150,0;28,6%' -Encoding utf8
$dry = Invoke-Pptcli -Step 'data refresh (dry run)' -Arguments @('data', 'refresh', '-s', $s, '--dry-run', 'true')
Invoke-Pptcli -Step 'data refresh' -Arguments @('data', 'refresh', '-s', $s) | Out-Null
Invoke-Pptcli -Step 'data bindings' -Arguments @('data', 'bindings', '-s', $s) | Out-Null

Stop-Demo -Demo $demo -Notes @("Columns detected: $(@($preview.columns).Count)", 'Refresh after adding a Q4 row to the CSV.')
