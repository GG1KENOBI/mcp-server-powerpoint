---
"powerpointmcp": patch
---

Fix the standalone MCP server exe failing every PowerPoint call with "Built-in COM has been disabled via a feature switch". The release exe is no longer trimmed, so PowerPoint automation works again (the exe is larger as a result).
