# Offline installation on Windows

This guide installs the PowerPoint MCP server and CLI on a Windows 10 or 11 (x64) computer with
no internet access. Nothing in the bundle downloads anything at run time, and no account,
telemetry, or cloud service is involved.

## Requirements

- Windows 10 or 11, x64. Use the win-arm64 bundle on ARM devices.
- Desktop PowerPoint: Microsoft 365 Apps or Office 2016 or later, installed and activated. The web
  and store versions cannot be automated.
- PowerShell 7.3 or later for the demos. The install and test scripts also run in Windows
  PowerShell 5.1.
- No .NET installation. The executables are self-contained.

## 1. Build the bundle (on a connected machine)

```powershell
git clone https://github.com/GG1KENOBI/mcp-server-powerpoint.git
cd mcp-server-powerpoint
git checkout feature/authoring-platform
pwsh scripts\Publish-OfflineBundle.ps1 -Version 0.4.0          # add -Runtime win-arm64 for ARM
```

This produces `artifacts\offline\PowerPointMcp-0.4.0-win-x64\` and a `.zip` of it. The
executables use the release workflow's settings: self-contained, single file, no ReadyToRun, and
never trimmed (trimming disables COM interop).

## 2. Copy and install (on the offline machine)

Copy the zip by USB or a file share, then:

```powershell
Expand-Archive PowerPointMcp-0.4.0-win-x64.zip -DestinationPath C:\Temp\PowerPointMcp
cd C:\Temp\PowerPointMcp
pwsh -File .\Install-PowerPointMcp.ps1 -AddCliToPath            # add -SmallModel for local models
```

The installer does the following:

1. Verifies every file against `SHA256SUMS.txt`.
2. Copies the bundle to `%LOCALAPPDATA%\Programs\PowerPointMcp`. No administrator rights are
   needed.
3. With `-AddCliToPath`, adds `cli` to the user PATH.
4. Prints ready-to-paste configuration for Claude Code, OpenCode, and generic MCP clients.

If Windows marks the files as downloaded, unblock them:
`Get-ChildItem -Recurse | Unblock-File`.

## 3. Check the installation

```powershell
cd $env:LOCALAPPDATA\Programs\PowerPointMcp
pwsh -File .\Test-PowerPointMcp.ps1              # registry, executables, --version, --help
pwsh -File .\Test-PowerPointMcp.ps1 -RoundTrip   # also creates, inspects, and closes a deck once
```

`-RoundTrip` starts PowerPoint through the CLI daemon, then stops the daemon, which closes only
the PowerPoint process the daemon started.

## 4. Connect a client

See [CLIENT-CONFIGURATION.md](CLIENT-CONFIGURATION.md). The short version for Claude Code:

```powershell
claude mcp add powerpoint -- "$env:LOCALAPPDATA\Programs\PowerPointMcp\mcp\mcp-powerpoint.exe"
```

## Settings

All settings are optional environment variables, set in the client's MCP configuration:

| Variable | Effect |
|---|---|
| `PPTMCP_LENIENT_ARGUMENTS=1` | Accepts near-miss argument names and action spellings (helps small models). |
| `PPTMCP_PREVIEW_IMAGES=off` | Omits PNG image content from preview results (text and wireframe stay). Use for models without vision. |
| `PPTMCP_PREVIEW_DIR` | Preview cache folder (default `%LOCALAPPDATA%\PowerPointMcp\previews`). |
| `PPTMCP_PROFILES_DIR` | Saved design profiles (default `%LOCALAPPDATA%\PowerPointMcp\profiles`). |
| `PPTMCP_ASSETS_DIR` | Image catalog (default `%LOCALAPPDATA%\PowerPointMcp\assets`). |
| `PPTMCP_UI=1` | Starts the local review UI on `http://localhost:<port>/` (loopback only, token in the printed URL). |
| `PPTMCP_UI_PORT` | Fixed port for the UI (default: a free port). |

## Uninstall

1. Close the clients that use the server.
2. Run `pptcli service stop` if the CLI daemon is running.
3. Delete `%LOCALAPPDATA%\Programs\PowerPointMcp`.
4. Optionally delete `%LOCALAPPDATA%\PowerPointMcp`, which holds previews, profiles, the asset
   catalog, and checkpoints.
5. Remove the `cli` folder from the user PATH.

## Troubleshooting

- **"PowerPoint.Application is not registered."** Install or repair desktop Office. Click-to-Run
  and MSI installations both work.
- **The first call is slow.** Starting PowerPoint takes time. Later calls reuse the session.
- **A security or macro prompt appears.** The server never dismisses prompts or enables macros.
  Answer the prompt in PowerPoint, or open trusted files only.
- **Fonts look different from the source deck.** `template import-slides` and
  `review validate` (`missing-font`) report fonts that are not installed. Install the fonts, or
  run `design normalize-typography`.
