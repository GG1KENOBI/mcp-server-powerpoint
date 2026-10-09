#Requires -Version 5.1
<#
.SYNOPSIS
    Checks an installed or unpacked PowerPoint MCP bundle without changing anything: PowerPoint
    registration and version, the executables, and — with -RoundTrip — a create/close session in a
    temporary folder through pptcli (this starts PowerPoint once).
#>
[CmdletBinding()]
param(
    [string]$Root = $PSScriptRoot,
    [switch]$RoundTrip
)
$ErrorActionPreference = 'Stop'
$ok = $true
function Report([string]$Check, [bool]$Passed, [string]$Detail) {
    $mark = if ($Passed) { 'OK  ' } else { 'FAIL' }
    Write-Host ("[{0}] {1,-28} {2}" -f $mark, $Check, $Detail)
    if (-not $Passed) { $script:ok = $false }
}

$progId = (Get-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CurVer' -ErrorAction SilentlyContinue).'(default)'
Report 'PowerPoint registered' ([bool]$progId) ($(if ($progId) { $progId } else { 'Install desktop PowerPoint (Microsoft 365 or Office 2016+).' }))
$c2r = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($c2r) { Report 'Office version' $true "$($c2r.VersionToReport) $($c2r.Platform)" }
Report 'Windows' ([Environment]::OSVersion.Version.Major -ge 10) ([Environment]::OSVersion.VersionString)

$mcp = Join-Path $Root 'mcp\mcp-powerpoint.exe'
$cli = Join-Path $Root 'cli\pptcli.exe'
Report 'MCP server present' (Test-Path -LiteralPath $mcp) $mcp
Report 'CLI present' (Test-Path -LiteralPath $cli) $cli
if (Test-Path -LiteralPath $mcp) {
    $version = & $mcp --version 2>&1 | Select-Object -First 1
    Report 'MCP server starts' ($LASTEXITCODE -eq 0) "$version"
}
if (Test-Path -LiteralPath $cli) {
    & $cli --help *> $null
    Report 'CLI starts' ($LASTEXITCODE -eq 0) 'pptcli --help'
}

if ($RoundTrip -and (Test-Path -LiteralPath $cli)) {
    $folder = Join-Path ([System.IO.Path]::GetTempPath()) "pptmcp-check-$([guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory -Path $folder | Out-Null
    $deck = Join-Path $folder 'check.pptx'
    $created = (& $cli session create $deck | Out-String) | ConvertFrom-Json
    Report 'Create presentation' ([bool]$created.success) $deck
    if ($created.success) {
        $summary = (& $cli deck summary -s $created.sessionId --include-theme false | Out-String) | ConvertFrom-Json
        Report 'Deck summary' ([bool]$summary.success) "slides: $($summary.slideCount)"
        $closed = (& $cli session close $created.sessionId --save | Out-String) | ConvertFrom-Json
        Report 'Close and save' ([bool]$closed.success) ''
    }
    & $cli service stop *> $null
    Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue
}

if ($ok) { Write-Host "`nAll checks passed." -ForegroundColor Green } else { Write-Host "`nSome checks failed." -ForegroundColor Red; exit 1 }
