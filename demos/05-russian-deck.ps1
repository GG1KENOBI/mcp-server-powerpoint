#Requires -Version 7.3
<#
.SYNOPSIS
    Scenario 5 — Russian-language deck: Cyrillic compositions, a windows-1251 CSV, case-insensitive
    Cyrillic find/replace with a count guard, and rich text edits that keep formatting.
#>
param([Parameter(Mandatory)][string]$OutputRoot)
if (-not (Get-Module DemoKit)) { Import-Module (Join-Path $PSScriptRoot 'DemoKit.psm1') }

$demo = Start-Demo -Name '05-russian-deck' -OutputRoot $OutputRoot
$s = $demo.session
Invoke-Pptcli -Step 'compose title' -Arguments @('compose', 'create', '-s', $s, '--spec', '{"kind":"title","id":"t","title":"Стратегия развития до 2027 года","subtitle":"Проект для обсуждения"}') | Out-Null
Invoke-Pptcli -Step 'compose process' -Arguments @('compose', 'create', '-s', $s, '--spec', '{"kind":"process","id":"p","title":"Как мы запускаем продукт","steps":[{"label":"Исследование","detail":"Интервью с клиентами"},{"label":"Пилот","detail":"Три региона"},{"label":"Масштабирование","detail":"Вся страна"}]}') | Out-Null
Invoke-Pptcli -Step 'compose timeline' -Arguments @('compose', 'create', '-s', $s, '--spec', '{"kind":"timeline","id":"tl","title":"Дорожная карта","milestones":[{"date":"Янв","label":"Пилот","status":"done"},{"date":"Апр","label":"Первая волна","status":"current"},{"date":"Сен","label":"Вторая волна"}]}') | Out-Null

# A legacy CSV in windows-1251 with semicolons and decimal commas.
$csv = Join-Path $demo.folder 'регионы-1251.csv'
$text = "Регион;Продажи;Доля`r`nЦентральный;1 234,5;41,2%`r`nПриволжский;812,0;27,1%`r`nСибирский;503,25;16,8%`r`n"
try { [System.IO.File]::WriteAllText($csv, $text, [System.Text.Encoding]::GetEncoding(1251)) }
catch { [System.IO.File]::WriteAllText($csv, $text, [System.Text.UTF8Encoding]::new($true)) }
$preview = Invoke-Pptcli -Step 'data preview (cp1251)' -Arguments @('data', 'preview', '-s', $s, '--source-path', $csv, '--culture', 'ru-RU')
Invoke-Pptcli -Step 'data create-table' -Arguments @('data', 'create-table', '-s', $s, '--source-path', $csv, '--culture', 'ru-RU', '--title', 'Продажи по регионам') | Out-Null

$found = Invoke-Pptcli -Step 'deck find-text ПИЛОТ' -Arguments @('deck', 'find-text', '-s', $s, '--find-what', 'ПИЛОТ')
$preview2 = Invoke-Pptcli -Step 'deck replace-text (preview)' -Arguments @('deck', 'replace-text', '-s', $s, '--find-what', 'пилот', '--replace-what', 'опытный запуск', '--whole-words', 'true')
Invoke-Pptcli -Step 'deck replace-text' -Arguments @('deck', 'replace-text', '-s', $s, '--find-what', 'пилот', '--replace-what', 'опытный запуск', '--whole-words', 'true', '--dry-run', 'false', '--expected-count', "$($preview2.totalCount)") | Out-Null

$box = Invoke-Pptcli -Step 'shape add-text-box' -Arguments @('shape', 'add-text-box', '-s', $s, '--slide-index', '1', '--left', '60', '--top', '400', '--width', '600', '--height', '80', '--text', "Ключевой вывод: рост выручки на 12%`rРиски: валютный курс")
Invoke-Pptcli -Step 'textframe format-range' -Arguments @('textframe', 'format-range', '-s', $s, '--slide-index', '1', '--shape-index', "$($box.shapeIndex)", '--match', 'Ключевой вывод:', '--bold', 'true', '--color', '#C00000') | Out-Null
Invoke-Pptcli -Step 'textframe replace-range' -Arguments @('textframe', 'replace-range', '-s', $s, '--slide-index', '1', '--shape-index', "$($box.shapeIndex)", '--match', '12%', '--text', '12,4%') | Out-Null
Invoke-Pptcli -Step 'textframe set-paragraph-format' -Arguments @('textframe', 'set-paragraph-format', '-s', $s, '--slide-index', '1', '--shape-index', "$($box.shapeIndex)", '--paragraph', '2', '--bullet-style', 'bullet', '--space-before', '6') | Out-Null
$paragraphs = Invoke-Pptcli -Step 'textframe get-paragraphs' -Arguments @('textframe', 'get-paragraphs', '-s', $s, '--slide-index', '1', '--shape-index', "$($box.shapeIndex)")

Stop-Demo -Demo $demo -Notes @("CSV columns: $(@($preview.columns | ForEach-Object name) -join ', ')", "Cyrillic case-insensitive hits for ПИЛОТ: $($found.totalCount)", "Runs in paragraph 1 after edits: $(@($paragraphs.paragraphs[0].runs).Count)")
