#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 3 — Template workflow: layouts with placeholders, slides from layouts filled by role
    (levels, picture, notes), slide import from another deck with a report, and a template preview.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

# A second deck to import from and to preview as a template.
$sourceDemo = Start-Demo -Name '03-template-source' -OutputRoot $OutputRoot
Invoke-Pptcli -Step 'source compose cards' -Arguments @('compose', 'create', '-s', $sourceDemo.session, '--spec', '{"kind":"process","id":"onboarding","title":"How onboarding works","steps":[{"label":"Apply"},{"label":"Review"},{"label":"Sign"},{"label":"Start"}]}', '--profile', 'minimal-mono') | Out-Null
Invoke-Pptcli -Step 'source compose quote' -Arguments @('compose', 'create', '-s', $sourceDemo.session, '--spec', '{"kind":"quote","id":"q","quote":{"text":"Simplicity is prerequisite for reliability.","author":"Edsger W. Dijkstra"}}', '--profile', 'minimal-mono') | Out-Null
$null = Stop-Demo -Demo $sourceDemo
$sourceDeck = Join-Path $sourceDemo.folder '03-template-source-final.pptx'

$demo = Start-Demo -Name '03-template-workflow' -OutputRoot $OutputRoot
$s = $demo.session
$layouts = Invoke-Pptcli -Step 'template list-layouts' -Arguments @('template', 'list-layouts', '-s', $s)
$content = @($layouts.layouts | Where-Object { @($_.placeholders | Where-Object role -eq 'title').Count -eq 1 -and @($_.placeholders | Where-Object role -eq 'body').Count -eq 1 })[0]
$two = @($layouts.layouts | Where-Object { @($_.placeholders | Where-Object role -eq 'body').Count -ge 2 })[0]
Invoke-Pptcli -Step 'template add-slide (content)' -Arguments @('template', 'add-slide', '-s', $s, '--layout', $content.name, '--content', '{"title":"Agenda","body":["Results",{"text":"Revenue and margin","level":2},"Priorities","Next steps"],"notes":"Keep to 20 minutes"}') | Out-Null
if ($two) {
    Invoke-Pptcli -Step 'template add-slide (two content)' -Arguments @('template', 'add-slide', '-s', $s, '--layout', $two.name, '--content', '{"title":"Options","body":["Build in-house","12 months"],"body2":["Buy a platform","8 weeks"]}') | Out-Null
}
$image = Join-Path $demo.folder 'team.png'
New-DemoImage -Path $image -Width 1600 -Height 1000 -Label 'Team offsite'
$added = Invoke-Pptcli -Step 'template add-slide (picture)' -Arguments @('template', 'add-slide', '-s', $s, '--layout', $content.name, '--content', (@{ title = 'Our team'; picture = $image } | ConvertTo-Json -Compress))
Invoke-Pptcli -Step 'template fill-placeholders' -Arguments @('template', 'fill-placeholders', '-s', $s, '--slide-index', "$($added.slideIndex)", '--content', '{"title":"Our team in 2025"}', '--remove-empty', 'true') | Out-Null
$import = Invoke-Pptcli -Step 'template import-slides' -Arguments @('template', 'import-slides', '-s', $s, '--source-path', $sourceDeck, '--slides', '1-2')
$preview = Invoke-Pptcli -Step 'template preview-template' -Arguments @('template', 'preview-template', '-s', $s, '--template-path', $sourceDeck)

Stop-Demo -Demo $demo -Notes @("Imported $(@($import.importedSlides).Count) slide(s); missing fonts: $(@($import.missingFonts) -join ', ')", "Template mapping: $((@($preview.layoutMapping) | ForEach-Object { "$($_.layout)→$($_.status)" }) -join '; ')")
