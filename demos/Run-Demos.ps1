#Requires -Version 7.3
<#
.SYNOPSIS
    Runs the demonstration scenarios against the installed desktop PowerPoint through pptcli and
    writes results.json and results.md with what actually happened (step timings, failures,
    slide counts, validation findings, output files). Nothing is reported for scenarios that did
    not run.

.PARAMETER Cli
    Path to pptcli.exe (default: pptcli on PATH, or the offline bundle's cli\pptcli.exe).

.PARAMETER OutputRoot
    Folder for decks, contact sheets, and results (default: .\demo-output next to this script).

.PARAMETER Only
    Scenario numbers or names to run, e.g. 1,4 or 'diagrams'.

.EXAMPLE
    pwsh -File demos\Run-Demos.ps1 -Cli C:\PowerPointMcp\cli\pptcli.exe
#>
[CmdletBinding()]
param(
    [string]$Cli,
    [string]$OutputRoot = (Join-Path $PSScriptRoot 'demo-output'),
    [string[]]$Only = @()
)
$ErrorActionPreference = 'Stop'
$PSNativeCommandArgumentPassing = 'Standard'

if (-not $Cli) {
    $bundled = Join-Path (Split-Path $PSScriptRoot -Parent) 'cli\pptcli.exe'
    $Cli = if (Test-Path -LiteralPath $bundled) { $bundled } else { 'pptcli' }
}
if (-not (Get-Command $Cli -ErrorAction SilentlyContinue)) {
    throw "pptcli was not found ('$Cli'). Pass -Cli with the full path to pptcli.exe."
}
Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') -Force
Set-DemoCli $Cli
New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null

function Get-PowerPointVersion {
    try {
        $c2r = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
        return "$($c2r.VersionToReport) ($($c2r.Platform))"
    }
    catch {
        $progId = (Get-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CurVer' -ErrorAction SilentlyContinue).'(default)'
        return $progId
    }
}

$environment = [ordered]@{
    date        = (Get-Date).ToString('o')
    windows     = [System.Environment]::OSVersion.VersionString
    powerPoint  = Get-PowerPointVersion
    cli         = (Get-Command $Cli).Source
    pwsh        = $PSVersionTable.PSVersion.ToString()
}

$scenarios = Get-ChildItem -LiteralPath $PSScriptRoot -Filter '[0-9][0-9]-*.ps1' | Sort-Object Name
if ($Only.Count -gt 0) {
    $scenarios = $scenarios | Where-Object { $name = $_.BaseName; $Only | Where-Object { $name -like "$($_.ToString().PadLeft(2, '0'))-*" -or $name -like "*$_*" } }
}

$records = [System.Collections.Generic.List[object]]::new()
foreach ($scenario in $scenarios) {
    Write-Host "`n=== $($scenario.BaseName)" -ForegroundColor Cyan
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $record = & $scenario.FullName -OutputRoot $OutputRoot
        $record = @($record)[-1]
        $records.Add([pscustomobject]@{ scenario = $scenario.BaseName; status = if ($record.stepsFailed -eq 0) { 'passed' } else { 'passed with failed steps' }; seconds = [math]::Round($watch.Elapsed.TotalSeconds, 1); result = $record; error = $null })
    }
    catch {
        Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
        $records.Add([pscustomobject]@{ scenario = $scenario.BaseName; status = 'failed'; seconds = [math]::Round($watch.Elapsed.TotalSeconds, 1); result = $null; error = $_.Exception.Message })
        # Leave nothing open: close any session this scenario left behind.
        try { & $Cli session list 2>$null | Out-Null } catch { }
    }
}

$report = [ordered]@{ environment = $environment; scenarios = $records }
$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $OutputRoot 'results.json') -Encoding utf8

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# Demonstration results')
$lines.Add('')
$lines.Add("Run on $($environment.date) — Windows $($environment.windows), PowerPoint $($environment.powerPoint), pptcli $($environment.cli).")
$lines.Add('')
$lines.Add('| Scenario | Status | Slides | Steps ok / failed | Seconds | Findings | Notes |')
$lines.Add('|---|---|---|---|---|---|---|')
foreach ($entry in $records) {
    $r = $entry.result
    $notes = if ($r) { ($r.notes -join '<br>') -replace '\|', '/' } else { $entry.error -replace '\|', '/' }
    $slides = if ($r) { $r.slides } else { '' }
    $steps = if ($r) { "$($r.stepsOk) / $($r.stepsFailed)" } else { '' }
    $found = if ($r) { "$($r.findings) $(if ($r.findingCodes) { '(' + ($r.findingCodes -join ', ') + ')' })" } else { '' }
    $lines.Add("| $($entry.scenario) | $($entry.status) | $slides | $steps | $($entry.seconds) | $found | $notes |")
}
$lines.Add('')
$lines.Add('Failed steps:')
foreach ($entry in $records | Where-Object { $_.result }) {
    foreach ($step in @($entry.result.steps) | Where-Object { -not $_.ok }) {
        $lines.Add("- $($entry.scenario) / $($step.step): $($step.error)")
    }
}
$lines | Set-Content -LiteralPath (Join-Path $OutputRoot 'results.md') -Encoding utf8
Write-Host "`nResults: $(Join-Path $OutputRoot 'results.md')" -ForegroundColor Green
