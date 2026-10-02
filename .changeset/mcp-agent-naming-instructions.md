---
"powerpointmcp": patch
---

The MCP server now tells AI clients its naming rules when they connect: kebab-case action names (for example `add-text-box`), camelCase `sessionId`/`filePath`/`targetPath` on the presentation tool, snake_case `session_id`/`slide_index`/`shape_index` on all other tools, 1-based indexes, the required position and size for `add-text-box`/`add-rectangle`, and how to save (there is no generic `save` action).
