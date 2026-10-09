# Client configuration

The server is a stdio MCP server: `mcp\mcp-powerpoint.exe` with no arguments. The paths below
assume the default install folder `%LOCALAPPDATA%\Programs\PowerPointMcp`; replace `YOU` with
your Windows user name. Samples are in the bundle's `config\` folder.

## Claude Code

```powershell
claude mcp add powerpoint -- "C:\Users\YOU\AppData\Local\Programs\PowerPointMcp\mcp\mcp-powerpoint.exe"
```

Or add it to a project's `.mcp.json` (see `config\claude-code.mcp.json`). Claude models read
images, so leave `PPTMCP_PREVIEW_IMAGES` unset to get rendered previews.

## OpenCode with a local model (llama.cpp)

`opencode.json` (global `%USERPROFILE%\.config\opencode\opencode.json` or per project):

```json
{
  "$schema": "https://opencode.ai/config.json",
  "provider": {
    "llamacpp": {
      "npm": "@ai-sdk/openai-compatible",
      "name": "llama.cpp (local)",
      "options": { "baseURL": "http://127.0.0.1:8080/v1" },
      "models": { "bonsai": { "name": "Bonsai (local)" } }
    }
  },
  "model": "llamacpp/bonsai",
  "mcp": {
    "powerpoint": {
      "type": "local",
      "command": ["C:\\Users\\YOU\\AppData\\Local\\Programs\\PowerPointMcp\\mcp\\mcp-powerpoint.exe"],
      "enabled": true,
      "environment": { "PPTMCP_LENIENT_ARGUMENTS": "1", "PPTMCP_PREVIEW_IMAGES": "off" }
    }
  }
}
```

Start llama.cpp's server with tool calling enabled, for example:

```powershell
llama-server -m bonsai.gguf --jinja -c 32768 --port 8080
```

The model file name and the context size depend on the build you use. The provider and npm
package names follow OpenCode's documentation for OpenAI-compatible endpoints. Check them
against your OpenCode version, because they change between releases.

## Hermes agent (or any client with a JSON or YAML MCP list)

Clients that take a list of MCP servers need three fields: the command, its arguments (none),
and environment variables. For a YAML configuration:

```yaml
mcp_servers:
  powerpoint:
    command: "C:\\Users\\YOU\\AppData\\Local\\Programs\\PowerPointMcp\\mcp\\mcp-powerpoint.exe"
    args: []
    env:
      PPTMCP_LENIENT_ARGUMENTS: "1"
      PPTMCP_PREVIEW_IMAGES: "off"
```

For a JSON configuration, see `config\generic-mcp.json`. Check the key names against your
client's documentation (`mcp_servers` or `mcpServers`, `env` or `environment`).

## Tips for small local models

- Set `PPTMCP_LENIENT_ARGUMENTS=1`. Near-miss argument names (`slideIndex`, `slide`) and action
  spellings (`AddTextBox`) are accepted. In strict mode the error names the canonical argument.
- Set `PPTMCP_PREVIEW_IMAGES=off` when the model cannot read images. `preview snapshot` still
  returns findings and a text wireframe.
- Ask the model to call `capabilities` (action `workflows`) first, then follow a recipe.
- Prefer high-level tools, because one call does the work of dozens of shape calls:
  - `compose create` with a short JSON spec
  - `template add-slide` with content by role
  - `data create-chart` from a CSV
- If the client can limit tools per agent, keep `capabilities`, `presentation`, `deck`,
  `compose`, `template`, `review_read`, and `preview_read` for drafting, and add others as
  needed.
- Bulk changes are dry runs by default; the model must pass `dry_run=false` to apply them.

## Local review UI

Add `PPTMCP_UI=1` to the server's environment. The server prints
`PowerPoint MCP local UI: http://localhost:<port>/#token=...` to stderr. The `capabilities` tool
returns the same URL. Open it in a browser on the same computer. The page is read-only and shows:

- slides with PowerPoint-rendered previews
- review findings
- picture checks
- design profiles
