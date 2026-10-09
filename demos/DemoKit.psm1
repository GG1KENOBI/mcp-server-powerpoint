# Shared helpers for the demonstration scenarios. Every call goes through pptcli, the same
# service the MCP server uses. Results are recorded only from what actually ran.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# JSON arguments must reach pptcli with their quotes intact (PowerShell 7.3+).
$PSNativeCommandArgumentPassing = 'Standard'

$script:Cli = 'pptcli'
$script:Steps = [System.Collections.Generic.List[object]]::new()
$script:Demo = $null

function Set-DemoCli([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $script:Cli = $Path
}

function Invoke-Pptcli {
    <#
    .SYNOPSIS
        Runs one pptcli command, parses its JSON output, records timing, and throws when the
        command reports success=false (unless -AllowFailure).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string[]]$Arguments,
        [switch]$AllowFailure
    )
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $raw = & $script:Cli @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    $watch.Stop()
    $text = ($raw | ForEach-Object { "$_" }) -join "`n"
    $json = $null
    $start = $text.IndexOf('{')
    if ($start -ge 0) {
        try { $json = $text.Substring($start) | ConvertFrom-Json -Depth 64 } catch { $json = $null }
    }
    $ok = $exitCode -eq 0 -and $null -ne $json -and ($json.PSObject.Properties.Name -notcontains 'success' -or $json.success)
    $entry = [ordered]@{
        step       = $Step
        command    = ($Arguments | Select-Object -First 2) -join ' '
        ok         = $ok
        seconds    = [math]::Round($watch.Elapsed.TotalSeconds, 2)
        error      = if ($ok) { $null } elseif ($json -and $json.PSObject.Properties.Name -contains 'errorMessage') { $json.errorMessage } else { $text.Substring(0, [math]::Min(400, $text.Length)) }
    }
    $script:Steps.Add([pscustomobject]$entry)
    $mark = if ($ok) { 'ok ' } else { 'ERR' }
    Write-Host ("  [{0}] {1,-48} {2,7:N2}s" -f $mark, $Step, $entry.seconds)
    if (-not $ok -and -not $AllowFailure) {
        throw "Step '$Step' failed: $($entry.error)"
    }
    return $json
}

function Start-Demo {
    <#
    .SYNOPSIS
        Creates the output folder and a new presentation session for one scenario.
    #>
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$OutputRoot)
    $folder = Join-Path $OutputRoot $Name
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    $script:Steps.Clear()
    $deck = Join-Path $folder "$Name.pptx"
    if (Test-Path -LiteralPath $deck) { Remove-Item -LiteralPath $deck -Force }
    $created = Invoke-Pptcli -Step 'session create' -Arguments @('session', 'create', $deck)
    $script:Demo = [ordered]@{ name = $Name; folder = $folder; deck = $deck; session = $created.sessionId; started = (Get-Date).ToString('o') }
    return [pscustomobject]$script:Demo
}

function Stop-Demo {
    <#
    .SYNOPSIS
        Validates the deck, renders a contact sheet, saves the result as a copy, closes the
        session, and returns the scenario record with measurements.
    #>
    param([Parameter(Mandatory)]$Demo, [string[]]$Notes = @())
    $session = $Demo.session
    $summary = Invoke-Pptcli -Step 'deck summary' -Arguments @('deck', 'summary', '-s', $session, '--include-theme', 'false') -AllowFailure
    $review = Invoke-Pptcli -Step 'review validate' -Arguments @('review', 'validate', '-s', $session, '--limit', '200') -AllowFailure
    $sheetPath = Join-Path $Demo.folder 'contact-sheet.png'
    Invoke-Pptcli -Step 'preview contact-sheet' -Arguments @('preview', 'contact-sheet', '-s', $session, '--output-path', $sheetPath, '--overwrite', 'true', '--include-image', 'false') -AllowFailure | Out-Null
    $final = Join-Path $Demo.folder "$($Demo.name)-final.pptx"
    Invoke-Pptcli -Step 'session save-copy-as' -Arguments @('session', 'save-copy-as', $session, $final, '--overwrite') -AllowFailure | Out-Null
    Invoke-Pptcli -Step 'session close' -Arguments @('session', 'close', $session, '--save') -AllowFailure | Out-Null

    $findings = @()
    if ($review -and $review.PSObject.Properties.Name -contains 'findings' -and $review.findings) { $findings = @($review.findings) }
    $record = [ordered]@{
        name          = $Demo.name
        started       = $Demo.started
        finished      = (Get-Date).ToString('o')
        slides        = if ($summary -and $summary.PSObject.Properties.Name -contains 'slideCount') { $summary.slideCount } else { $null }
        stepsOk       = @($script:Steps | Where-Object ok).Count
        stepsFailed   = @($script:Steps | Where-Object { -not $_.ok }).Count
        totalSeconds  = [math]::Round((@($script:Steps | Measure-Object -Property seconds -Sum).Sum), 2)
        findings      = $findings.Count
        findingCodes  = @($findings | Group-Object code | ForEach-Object { "$($_.Name)=$($_.Count)" })
        deck          = $Demo.deck
        final         = if (Test-Path -LiteralPath $final) { $final } else { $null }
        contactSheet  = if (Test-Path -LiteralPath $sheetPath) { $sheetPath } else { $null }
        notes         = $Notes
        steps         = @($script:Steps)
    }
    $record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $Demo.folder 'result.json') -Encoding utf8
    return [pscustomobject]$record
}

function New-DemoImage {
    <#
    .SYNOPSIS
        Draws a simple local PNG (gradient and label) so image scenarios need no downloads.
    #>
    param([Parameter(Mandatory)][string]$Path, [int]$Width = 1600, [int]$Height = 900, [string]$Label = 'Demo', [string]$From = '#1F4E79', [string]$To = '#9DC3E6')
    Add-Type -AssemblyName System.Drawing
    $bitmap = [System.Drawing.Bitmap]::new($Width, $Height)
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $rect = [System.Drawing.Rectangle]::new(0, 0, $Width, $Height)
            $brush = [System.Drawing.Drawing2D.LinearGradientBrush]::new($rect, [System.Drawing.ColorTranslator]::FromHtml($From), [System.Drawing.ColorTranslator]::FromHtml($To), 35.0)
            $graphics.FillRectangle($brush, $rect)
            $font = [System.Drawing.Font]::new('Segoe UI', [float]($Height / 10), [System.Drawing.FontStyle]::Bold)
            $graphics.DrawString($Label, $font, [System.Drawing.Brushes]::White, [float]($Width / 20), [float]($Height / 2.5))
            $brush.Dispose(); $font.Dispose()
        }
        finally { $graphics.Dispose() }
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
}

Export-ModuleMember -Function Set-DemoCli, Invoke-Pptcli, Start-Demo, Stop-Demo, New-DemoImage
