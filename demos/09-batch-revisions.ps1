#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 9 — Batch with references, revision guard, checkpoint, change log, and a background
    job; a stale revision is refused before anything changes.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '09-batch-revisions' -OutputRoot $OutputRoot
$s = $demo.session
$ops = @(
    @{ id = 'box'; tool = 'shape'; action = 'add-text-box'; args = @{ slide_index = 1; left = 60; top = 60; width = 600; height = 60; text = 'Batch-created headline' } },
    @{ id = 'bold'; tool = 'textframe'; action = 'format-range'; args = @{ slide_index = 1; shape_index = '$box.shapeIndex'; match = 'headline'; bold = $true; color = '#1F4E79' } },
    @{ id = 'ids'; tool = 'deck'; action = 'assign-ids'; args = @{ selector = 'slide:1 kind:text-box'; prefix = 'demo' } },
    @{ id = 'cards'; tool = 'compose'; action = 'create'; args = @{ spec = '{"kind":"cards","id":"batch-cards","title":"Made in a batch","cards":[{"heading":"One","body":"First"},{"heading":"Two","body":"Second"}]}' } }
) | ConvertTo-Json -Depth 6 -Compress

$validation = Invoke-Pptcli -Step 'batch validate' -Arguments @('batch', 'validate', '-s', $s, '--operations', $ops)
$revision = (Invoke-Pptcli -Step 'deck fingerprint' -Arguments @('deck', 'fingerprint', '-s', $s)).revision
$run = Invoke-Pptcli -Step 'batch run (checkpoint)' -Arguments @('batch', 'run', '-s', $s, '--operations', $ops, '--expected-revision', $revision, '--checkpoint', 'true')
$stale = Invoke-Pptcli -Step 'batch run (stale revision refused)' -Arguments @('batch', 'run', '-s', $s, '--operations', $ops, '--expected-revision', $revision) -AllowFailure

$jobOps = '[{"id":"s","tool":"deck","action":"summary"},{"id":"v","tool":"review","action":"validate"}]'
$job = Invoke-Pptcli -Step 'batch start' -Arguments @('batch', 'start', '-s', $s, '--operations', $jobOps)
$status = $null
for ($i = 0; $i -lt 60; $i++) {
    $status = Invoke-Pptcli -Step "batch status ($i)" -Arguments @('batch', 'status', '-s', $s, '--job-id', $job.jobId)
    if ($status.state -notin @('queued', 'running')) { break }
    Start-Sleep -Milliseconds 500
}

Stop-Demo -Demo $demo -Notes @("Validated ops: $(@($validation.operations).Count); run completed $($run.completed)/$($run.total); changes logged: $(@($run.changes).Count); checkpoint: $($run.checkpointPath)", "Stale revision refused: $(-not $stale.success)", "Background job state: $($status.state)")
