---
name: powerpoint-mcp
description: >
  Use the configured PowerPoint MCP server to inspect or edit presentations on Windows.
  Covers session safety, live tool discovery, saving, and visual verification; not general
  presentation design. Triggers: PowerPoint MCP, presentation tools, edit a PowerPoint deck.
compatibility: Windows with Microsoft PowerPoint desktop installed and the PowerPoint MCP server configured.
---

# PowerPoint MCP

Use this skill only for tasks performed through the configured PowerPoint MCP server. For
general presentation design, use the optional `powerpoint-deck-design` skill. Tool schemas
advertised by the live server define the available actions and arguments; do not rely on a
copied command catalog.

## Safe editing loop

1. Call `capabilities` (no session needed) for the installed PowerPoint version, every tool
   with its read-only actions, settings, and workflow recipes. The live tool schemas define the
   arguments. Use a `{domain}_read` tool when it provides the inspection action you need.
2. Create or open a presentation once. Pass the returned `presentation_session_id` to later
   calls; do not open a file a second time after `create`.
3. Understand before editing: `deck summary`, then `deck inspect-objects` or `deck find` with a
   selector (`kind:table slide:3`, `text:"Q3*"`, `appid:...`). Results carry stable
   `slide_id`/`shape_id` values; indices start at 1.
4. Prefer the high-level tools: `template add-slide` for the deck's own layouts, `compose plan`
   then `compose create` for designed slides, `diagram create`, `data create-table/create-chart`
   from CSV/XLSX, `asset place` for local images (never stretched).
5. Bulk changes preview first: `deck replace-text`, `design apply-profile`,
   `design normalize-typography`, `review repair`, and `asset embed-linked/fix-links` are dry
   runs until `dry_run=false`. Pass `expected_count` or `expected_revision` when it matters.
6. Check the result: `review validate` lists overflow, overlap, off-slide objects, small fonts,
   contrast, alt text, and inconsistencies with evidence; `preview snapshot` renders a slide
   (image content when the client supports it, plus a text wireframe that does not need images).
7. Keep the original: save with `presentation save-as` to a new path unless the user asked to
   overwrite. Close with `save: true` only when the user wants the changes kept. Close returns
   before PowerPoint finishes its background cleanup; do not wait for its process to exit.

The server requires Windows and desktop PowerPoint; it is not a headless or cross-platform
converter. See the [workflow guide](https://powerpointmcpserver.dev/reference/workflows/) and
[behavioral rules](https://powerpointmcpserver.dev/reference/behavioral-rules/) for edge cases.
