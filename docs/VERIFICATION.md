# Verification procedure and results

This page says how to verify the authoring platform on Windows with real PowerPoint, and records
what has been verified so far. Only record results that were observed.

## What has been verified

| Date | Environment | What ran | Result |
|---|---|---|---|
| 2026-10-09 | Linux container, .NET 10 SDK, **no PowerPoint** | Release build of the solution | Pass |
| 2026-10-09 | same | PowerPoint-free tests: `dotnet test tests/PowerPointMcp.McpServer.Tests --filter "RequiresPowerPoint!=true"` | 563 / 563 pass |
| 2026-10-09 | same | Audits: check-com-leaks, check-dynamic-casts, check-success-flag, check-core-interface-completeness, check-doc-counts | Pass |
| 2026-10-09 | same | Local UI security tests (real HttpListener on loopback) | Pass (part of the 563) |
| — | Windows + PowerPoint | Native Core tests for the new domains (64) | **Not run yet** |
| — | Windows + PowerPoint | MCP end-to-end workflow (`McpPlatformWorkflowTests`) | **Not run yet** |
| — | Windows + PowerPoint | Ten demonstration scenarios | **Not run yet** |
| — | Windows | Offline bundle build, install, and `Test-PowerPointMcp.ps1 -RoundTrip` | **Not run yet** |

Add a row for each run below, and paste the demo results table into the section at the end.

## Prerequisites

- Windows 10 or 11 with desktop PowerPoint (Microsoft 365 or Office 2016+), signed in and past
  first-run dialogs.
- .NET 10 SDK (only for building and running the tests), PowerShell 7.3+.
- Close presentations you care about. The tests start their own PowerPoint process and close only
  that process. They never terminate unrelated PowerPoint processes.

## 1. Build and run the PowerPoint-free tests

```powershell
git clone https://github.com/GG1KENOBI/mcp-server-powerpoint.git
cd mcp-server-powerpoint
git checkout feature/authoring-platform
dotnet build Sbroenne.PowerPointMcp.slnx -c Release
dotnet test tests\PowerPointMcp.McpServer.Tests -c Release --filter "RequiresPowerPoint!=true"
```

Expected: 563 tests pass, as on Linux.

## 2. Native tests for the new features

Run one feature at a time. Each class shares one PowerPoint process across its tests.

```powershell
foreach ($feature in 'Deck', 'Review', 'Composition', 'Diagram', 'Preview', 'Data', 'TextFrame', 'Design', 'Template', 'Assets') {
    dotnet test tests\PowerPointMcp.Core.Tests -c Release --filter "Feature=$feature" --logger "trx;LogFileName=$feature.trx"
}
```

`Feature=TextFrame` also runs the pre-existing textframe tests together with
`RichTextCommandsTests`.

Then run the whole native suite once, so regressions in the original domains are caught too:

```powershell
dotnet test tests\PowerPointMcp.Core.Tests -c Release
```

## 3. MCP end-to-end workflow

```powershell
dotnet test tests\PowerPointMcp.McpServer.Tests -c Release --filter "FullyQualifiedName~McpPlatformWorkflowTests|FullyQualifiedName~McpAuthoringWorkflowTests"
```

`McpPlatformWorkflowTests` drives one session through the protocol:

1. capabilities
2. template add-slide
3. compose plan and create
4. data create-chart from a Russian CSV
5. diagram create
6. deck replace-text with a count guard
7. a design dry run
8. review validate
9. preview snapshot, which must return PNG image content
10. batch validate
11. asset place and inspect
12. save-copy-as and close

## 4. Demonstration scenarios

```powershell
dotnet publish src\PowerPointMcp.CLI -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=false -o artifacts\cli
pwsh -File demos\Run-Demos.ps1 -Cli artifacts\cli\powerpointcli.exe
```

Each scenario writes the following to `demos\demo-output\<scenario>\`:

- the deck
- `<scenario>-final.pptx`
- `contact-sheet.png`
- `result.json`

The runner writes `results.json` and `results.md` with these measurements:

- step timings
- failed steps with their error messages
- slide counts
- validation findings by rule
- scenario notes, for example the planned changes, matches replaced, picture PPI, and findings
  before and after repair

Open each contact sheet and final deck and check them by eye:

1. **Quarterly review:** the Cyrillic text and ₽ render correctly. The chart and table show 4
   quarters after the refresh.
2. **Restyle:** colors, fonts, and cards follow `acme` (teal, Georgia headings). No text
   overflows.
3. **Template:** the slides use the deck's layouts. The picture fills its placeholder without
   distortion. The imported slides keep their content.
4. **Diagrams:** all five types are present. The connectors stay attached when you move a node in
   PowerPoint.
5. **Russian deck:** "пилот" is replaced as a whole word only. The bold red run keeps its format.
6. **Overflow:** the long summary splits onto continuation slides. Fewer findings remain after
   repair.
7. **Renaming:** no "Contoso" is left, notes included. Prices read "10 USD".
8. **Images:** contain and cover are correct, with no stretching. The low-PPI thumbnail is
   flagged.
9. **Batch:** the checkpoint file exists. The stale revision is refused.
10. **Accessibility:** the findings include `low-contrast`, `small-text`, and `missing-alt-text`
    before the fixes, and fewer after.

## 5. Offline bundle

```powershell
pwsh scripts\Publish-OfflineBundle.ps1 -Version 0.4.0
# copy artifacts\offline\PowerPointMcp-0.4.0-win-x64.zip to a machine without internet, then:
pwsh -File .\Install-PowerPointMcp.ps1 -AddCliToPath
pwsh -File $env:LOCALAPPDATA\Programs\PowerPointMcp\Test-PowerPointMcp.ps1 -RoundTrip
```

Disconnect the network first to confirm that nothing needs it. Then connect a client (see
[CLIENT-CONFIGURATION.md](CLIENT-CONFIGURATION.md)) and ask it to build a three-slide deck from a
CSV. Record the client and model you used.

## 6. Local UI

Set `PPTMCP_UI=1` for the server, open the printed URL, and check the following:

- The slides list matches the open deck. Selecting a slide shows PowerPoint's rendering.
- Opening the URL without the `#token=...` part shows the token message, and the API returns 401.
- From another machine, the port is not reachable, because the server listens on localhost only.

## Recording results

For each run, add a row to the table at the top with:

- the date
- the Windows and PowerPoint versions (`Test-PowerPointMcp.ps1` prints them)
- what ran
- pass and fail counts

Paste the demo `results.md` below. For failures, open an issue or a fix commit that names the
failing test or step. Do not mark a native feature as verified until its tests passed on real
PowerPoint.

## Demo results

Not run yet.
