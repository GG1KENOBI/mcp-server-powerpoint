#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 4 — Editable diagrams: flowchart, swimlane, matrix, hub-and-spoke, and architecture
    with glued connectors; then node edits, a new node and edge, and a relayout.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '04-diagrams' -OutputRoot $OutputRoot
$s = $demo.session
foreach ($type in 'flowchart', 'swimlane', 'matrix', 'hub-spoke', 'architecture') {
    Invoke-Pptcli -Step "diagram create $type" -Arguments @('diagram', 'create', '-s', $s, '--spec-path', (Join-Path $PSScriptRoot "specs/diagram-$type.json")) | Out-Null
}
$inspect = Invoke-Pptcli -Step 'diagram inspect approval' -Arguments @('diagram', 'inspect', '-s', $s, '--diagram-id', 'approval')
Invoke-Pptcli -Step 'diagram update-node' -Arguments @('diagram', 'update-node', '-s', $s, '--diagram-id', 'approval', '--node-id', 'cfo', '--label', 'CFO and legal review') | Out-Null
Invoke-Pptcli -Step 'diagram add-node' -Arguments @('diagram', 'add-node', '-s', $s, '--diagram-id', 'approval', '--node-id', 'archive', '--label', 'Archive request', '--connect-from', 'done') | Out-Null
Invoke-Pptcli -Step 'diagram relayout' -Arguments @('diagram', 'relayout', '-s', $s, '--diagram-id', 'approval') | Out-Null
$after = Invoke-Pptcli -Step 'diagram inspect (after)' -Arguments @('diagram', 'inspect', '-s', $s, '--diagram-id', 'approval')
Invoke-Pptcli -Step 'diagram list' -Arguments @('diagram', 'list', '-s', $s) | Out-Null

$unglued = @($after.edges | Where-Object { -not $_.attached }).Count
Stop-Demo -Demo $demo -Notes @("Approval nodes: $(@($inspect.nodes).Count) → $(@($after.nodes).Count); unglued edges after relayout: $unglued")
