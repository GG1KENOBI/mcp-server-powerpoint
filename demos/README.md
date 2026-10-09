# Demonstration scenarios

The ten scenarios exercise the authoring platform against desktop PowerPoint through `pptcli`.
`pptcli` uses the same service as the MCP server.

| # | Scenario | Shows |
|---|---|---|
| 01 | Quarterly review | Russian CSV with decimal commas, title/KPI/summary compositions, bound chart and table, refresh after the file changes |
| 02 | Restyle | Import a custom profile, apply it with a migration preview, normalize typography, update components, extract a profile |
| 03 | Template workflow | Layouts with placeholders, slides filled by role (levels, picture, notes), slide import with a report, template preview |
| 04 | Diagrams | Flowchart, swimlane, matrix, hub-and-spoke, architecture; node edits, a new node and edge, relayout, glue check |
| 05 | Russian deck | Cyrillic compositions, windows-1251 CSV, case-insensitive Cyrillic find/replace with a count guard, rich text edits |
| 06 | Overflow and repair | Measured fitting with continuation slides; overflow and off-slide findings; plan-repair and repair |
| 07 | Deck-wide renaming | Matches in titles, bullets, tables, and notes; a refused wrong count; regex replacement |
| 08 | Images | Local catalog, search, contain and cover placement, replace in frame, decorative flag, PPI and distortion checks |
| 09 | Batch and revisions | References between steps, revision guard, checkpoint, change log, stale revision refused, background job |
| 10 | Review and accessibility | Contrast, small text, alt text findings; rendered snapshots; image diff before and after fixes |

## Run

Requires Windows, desktop PowerPoint, PowerShell 7.3+, and `pptcli` (from the offline bundle,
the release, or `dotnet publish src\PowerPointMcp.CLI`).

```powershell
pwsh -File demos\Run-Demos.ps1                      # uses cli\pptcli.exe from the bundle, else pptcli on PATH
pwsh -File demos\Run-Demos.ps1 -Cli C:\path\pptcli.exe -Only 1,4
```

The output goes to `demos\demo-output\`: one folder per scenario with the deck, a final copy, a
contact sheet, and `result.json`, plus `results.json` and `results.md` for the whole run. The
results report what actually happened, including failed steps and their errors.

Every scenario generates its own pictures with System.Drawing, so nothing is downloaded. It
works on copies of the data files in its output folder.
