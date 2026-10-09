#Requires -Version 5.1
<#
.SYNOPSIS
    Installs the offline PowerPoint MCP bundle for the current user (no administrator rights, no
    downloads): copies it to %LOCALAPPDATA%\Programs\PowerPointMcp, verifies checksums, optionally
    adds the CLI to the user PATH, and prints ready-to-paste client configuration.

.PARAMETER Destination
    Install folder (default %LOCALAPPDATA%\Programs\PowerPointMcp).

.PARAMETER AddCliToPath
    Add the cli folder to the user PATH.

.PARAMETER SmallModel
    Print client configuration tuned for small local models (lenient argument names, no image
    content in preview results).
#>
[CmdletBinding()]
param(
    [string]$Destination = (Join-Path $env:LOCALAPPDATA 'Programs\PowerPointMcp'),
    [switch]$AddCliToPath,
    [switch]$SmallModel
)
$ErrorActionPreference = 'Stop'
$source = $PSScriptRoot

Write-Host 'Checking checksums...'
$bad = @()
foreach ($line in Get-Content -LiteralPath (Join-Path $source 'SHA256SUMS.txt')) {
    if ($line -notmatch '^(?<hash>[0-9a-f]{64})  (?<file>.+)$') { continue }
    $path = Join-Path $source $Matches.file
    if (-not (Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $Matches.hash) { $bad += $Matches.file }
}
if ($bad.Count -gt 0) { throw "Checksum mismatch or missing files: $($bad -join ', '). Copy the bundle again." }

if ((Resolve-Path -LiteralPath $source).Path -ne [System.IO.Path]::GetFullPath($Destination)) {
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    Copy-Item -Path (Join-Path $source '*') -Destination $Destination -Recurse -Force
}
$mcp = Join-Path $Destination 'mcp\mcp-powerpoint.exe'
$cli = Join-Path $Destination 'cli\pptcli.exe'

if ($AddCliToPath) {
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    $cliFolder = Split-Path $cli
    if (($userPath -split ';') -notcontains $cliFolder) {
        [Environment]::SetEnvironmentVariable('Path', ($userPath.TrimEnd(';') + ';' + $cliFolder), 'User')
        Write-Host "Added $cliFolder to the user PATH (open a new terminal to use pptcli)."
    }
}

$serverEnv = [ordered]@{}
if ($SmallModel) {
    $serverEnv['PPTMCP_LENIENT_ARGUMENTS'] = '1'
    $serverEnv['PPTMCP_PREVIEW_IMAGES'] = 'off'
}
$escaped = $mcp -replace '\\', '\\'
$envJson = if ($serverEnv.Count -gt 0) { ', "env": ' + ($serverEnv | ConvertTo-Json -Compress) } else { '' }
Write-Host "`nInstalled to $Destination" -ForegroundColor Green
Write-Host "`nClaude Code:   claude mcp add powerpoint -- `"$mcp`""
Write-Host "Generic MCP JSON (Claude Desktop, Hermes and other clients):"
Write-Host "  { `"mcpServers`": { `"powerpoint`": { `"command`": `"$escaped`", `"args`": []$envJson } } }"
Write-Host "OpenCode (opencode.json):"
Write-Host "  { `"mcp`": { `"powerpoint`": { `"type`": `"local`", `"command`": [`"$escaped`"], `"enabled`": true$(if ($serverEnv.Count -gt 0) { ', "environment": ' + ($serverEnv | ConvertTo-Json -Compress) }) } } }"
Write-Host "`nNext: run Test-PowerPointMcp.ps1 to check PowerPoint and the executables."
