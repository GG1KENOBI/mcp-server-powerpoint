---
"powerpointmcp": minor
---

The MCP server now accepts near-miss tool calls from smaller and local AI models. Action names in other spellings (for example `add_text_box` or `addTextBox` for `add-text-box`) and parameter names in the other case style (`session_id` on the presentation tool, `sessionId` or `slideIndex` on the other tools) are mapped to the canonical name when exactly one name matches. Errors now suggest the closest action or parameter, list every missing required parameter, and include a short example call. Tool schemas and canonical names are unchanged.
