#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Builds an offline Windows bundle: self-contained MCP server and CLI executables (no .NET
    install needed), demos, documentation, client configuration samples, install and verify
    scripts, and SHA-256 checksums, zipped for copying to a machine without internet access.

.DESCRIPTION
    Uses the same publish settings as the release workflow: self-contained, single file,
    PublishReadyToRun=false, and never trimmed (trimming disables built-in COM interop). Nothing
    in the bundle downloads anything at run time.

.PARAMETER Version
    Version stamped into the executables (default: Version from Directory.Build.props).

.PARAMETER Runtime
    win-x64 (default) or win-arm64.

.PARAMETER OutputDir
    Folder for the bundle folder and zip (default: artifacts/offline).

.EXAMPLE
    pwsh scripts/Publish-OfflineBundle.ps1 -Version 0.4.0
#>
[CmdletBinding()]
param(
    [string]$Version,
    [ValidateSet('win-x64', 'win-arm64')]
    [string]$Runtime = 'win-x64',
    [string]$OutputDir = 'artifacts/offline'
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($Version)) {
    $props = [xml](Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw)
    $Version = @($props.Project.PropertyGroup | ForEach-Object { $_.Version } | Where-Object { $_ })[0]
    if (-not $Version) { $Version = '0.0.0' }
}
$numeric = ($Version -split '-')[0]
$name = "PowerPointMcp-$Version-$Runtime"
$output = if ([System.IO.Path]::IsPathRooted($OutputDir)) { $OutputDir } else { Join-Path $repo $OutputDir }
$bundle = Join-Path $output $name
if (Test-Path -LiteralPath $bundle) { Remove-Item -LiteralPath $bundle -Recurse -Force }
New-Item -ItemType Directory -Force -Path $bundle | Out-Null

function Publish-Executable([string]$Project, [string]$Folder) {
    $target = Join-Path $bundle $Folder
    & dotnet publish (Join-Path $repo $Project) `
        --configuration Release `
        --runtime $Runtime `
        --self-contained true `
        -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:PublishReadyToRun=false `
        -p:PublishTrimmed=false `
        -p:Version=$Version `
        -p:AssemblyVersion="$numeric.0" `
        -p:FileVersion="$numeric.0" `
        -p:NuGetAudit=false `
        --output $target | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed for $Project." }
    # Debug symbols and XML API docs are not needed at run time.
    Get-ChildItem -LiteralPath $target -File | Where-Object { $_.Extension -in '.pdb', '.xml' } | Remove-Item -Force
    return $target
}

Write-Host "Publishing MCP server ($Runtime, self-contained, untrimmed)..." -ForegroundColor Cyan
$mcp = Publish-Executable 'src/PowerPointMcp.McpServer/PowerPointMcp.McpServer.csproj' 'mcp'
Rename-Item -LiteralPath (Join-Path $mcp 'Sbroenne.PowerPointMcp.McpServer.exe') 'mcp-powerpoint.exe'
Write-Host "Publishing CLI ($Runtime, self-contained, untrimmed)..." -ForegroundColor Cyan
$cli = Publish-Executable 'src/PowerPointMcp.CLI/PowerPointMcp.CLI.csproj' 'cli'
Rename-Item -LiteralPath (Join-Path $cli 'powerpointcli.exe') 'pptcli.exe'

# Demos, documentation, skills, client configuration samples, and helper scripts.
Copy-Item -LiteralPath (Join-Path $repo 'demos') -Destination (Join-Path $bundle 'demos') -Recurse
Get-ChildItem -LiteralPath (Join-Path $bundle 'demos') -Directory -Filter 'demo-output' -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
$docs = New-Item -ItemType Directory -Force -Path (Join-Path $bundle 'docs')
foreach ($doc in 'AUTHORING-PLATFORM.md', 'OFFLINE-INSTALL.md', 'CLIENT-CONFIGURATION.md', 'VERIFICATION.md', 'CAPABILITY-MATRIX.md') {
    Copy-Item -LiteralPath (Join-Path $repo "docs/$doc") -Destination $docs
}
Copy-Item -LiteralPath (Join-Path $repo 'LICENSE') -Destination $bundle
Copy-Item -LiteralPath (Join-Path $repo 'PRIVACY.md') -Destination $docs
Copy-Item -LiteralPath (Join-Path $repo 'skills') -Destination (Join-Path $bundle 'skills') -Recurse
Copy-Item -LiteralPath (Join-Path $repo 'distribution/config') -Destination (Join-Path $bundle 'config') -Recurse
foreach ($script in 'Install-PowerPointMcp.ps1', 'Test-PowerPointMcp.ps1') {
    Copy-Item -LiteralPath (Join-Path $repo "distribution/$script") -Destination $bundle
}
@"
PowerPoint MCP $Version ($Runtime) — offline bundle

mcp\mcp-powerpoint.exe   MCP server (stdio). Self-contained: no .NET installation needed.
cli\pptcli.exe           Command-line client (same commands as the MCP tools).
Install-PowerPointMcp.ps1  Copies the bundle to %LOCALAPPDATA%\Programs\PowerPointMcp (no admin rights).
Test-PowerPointMcp.ps1     Checks PowerPoint, the executables, and (optionally) a create/close round trip.
config\                  Client configuration samples (OpenCode, Claude Code, generic MCP JSON).
demos\                   Ten demonstration scenarios: pwsh demos\Run-Demos.ps1
docs\                    Feature guide, offline install, client configuration, verification procedure.
SHA256SUMS.txt           Checksums of every file in this bundle.

Requirements: Windows 10/11, desktop PowerPoint (Microsoft 365 or Office 2016+), PowerShell 7.3+ for the scripts.
Nothing in this bundle downloads anything or contacts any service.
"@ | Set-Content -LiteralPath (Join-Path $bundle 'README.txt') -Encoding utf8

$sums = Get-ChildItem -LiteralPath $bundle -Recurse -File | Sort-Object FullName | ForEach-Object {
    $relative = [System.IO.Path]::GetRelativePath($bundle, $_.FullName) -replace '/', '\'
    "$((Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant())  $relative"
}
$sums | Set-Content -LiteralPath (Join-Path $bundle 'SHA256SUMS.txt') -Encoding ascii

$zip = Join-Path $output "$name.zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $bundle '*') -DestinationPath $zip
Write-Host "Bundle: $bundle" -ForegroundColor Green
Write-Host "Zip:    $zip ($([math]::Round((Get-Item -LiteralPath $zip).Length / 1MB, 1)) MB)" -ForegroundColor Green
