# Agent Authoring Platform

This document describes the authoring upgrade of the PowerPoint MCP server. It covers what was
built, how to use it, how it was verified, and what is still open. It is written for people who
install the server on Windows and for the agents and models that drive it.

**Status in one paragraph.** Every feature below is implemented in code on the
`feature/authoring-platform` branch. The pure logic of every feature has unit tests that run
without PowerPoint (450 new test cases; 563 PowerPoint-free tests in total, all passing). The
build and the repository audits pass: COM release patterns, documented dynamic casts, the
Success/ErrorMessage invariant, interface completeness, and documentation counts. Behaviour
against real PowerPoint is covered by 64 new native tests, one MCP end-to-end workflow test, and
ten demonstration scenarios.

**None of the native tests, workflow tests, or demonstrations has been run yet.** They were
written and compiled in a Linux container without Windows or PowerPoint. Native behaviour
therefore stays unverified until the procedure in [VERIFICATION.md](VERIFICATION.md) has been run
on Windows. This document never claims a PowerPoint result that was not observed.

## Contents

- [Architecture](#architecture)
- [Tools at a glance](#tools-at-a-glance)
- [Workflows](#workflows)
- [Feature guide](#feature-guide)
- [Schemas and persistent tags](#schemas-and-persistent-tags)
- [Safety, offline operation, and reliability](#safety-offline-operation-and-reliability)
- [Verification status](#verification-status)
- [Known limitations](#known-limitations)
- [Commits](#commits)

## Architecture

The existing layers are unchanged: COM interop runs on an STA queue, Core holds the domain
logic, a single shared service dispatches commands, and the MCP server and the CLI are two
front ends on top of it. New features are new Core domains with `[ServiceCategory]` and
`[McpTool]` attributes. The generators turn each domain into an MCP tool, a `_read` tool for its
read-only actions, and a CLI command, so the MCP server and the CLI expose the same operations.

```
MCP client ──stdio──► McpServer ─┐                      ┌─► Core domains (deck, compose, diagram, review,
                                 ├─► PowerPointMcpService┤    preview, batch, data, design, template, asset,
pptcli ──named pipe──► daemon ───┘   (one dispatch,      │    and the original 17)
Local UI (opt-in) ──loopback HTTP┘    busy guard,        └─► ComInterop: PresentationBatch (STA queue)
                                      preview invalidation)        └─► desktop PowerPoint (COM)
```

Design rules that were kept:

- **One backend.** The optional local UI and the batch engine both call
  `PowerPointMcpService.ProcessAsync`. There is no second implementation of any command.
- **No model in the server.** Everything is deterministic code: layout math, text estimation and
  measurement, validation rules, color and type migration, CSV/XLSX typing, and diagram layout.
- **Text first.** Every result is JSON text. Preview images are added as MCP image content when
  the client supports them, and the text wireframe and findings are always present, so models
  that cannot read images lose no information.
- **Pure planners, thin COM.** Planning is pure and unit-tested: layout, fitting estimates,
  selectors, fingerprints, repair plans, color maps, typography, placeholder assignment, data
  typing, diagram ranking, and image headers. The COM code reads facts and applies planned
  changes. Every COM object is released in a `finally` block.
- **Measured, not guessed.** Compose and repair measure the rendered text height in PowerPoint
  (`TextRange.BoundHeight`) after placing it. They shrink fonts within the profile's limits, or
  split content onto continuation slides. Text is never shortened.

## Tools at a glance

There are 53 MCP tools with 305 operations across 27 domains, plus 25 read-only aliases. New in
this upgrade:

| Tool | Actions | Purpose |
|---|---|---|
| `capabilities` | overview, tools, workflows, environment | Needs no session and never starts PowerPoint. Reports versions, tools with their read-only actions, settings, catalogs, and workflow recipes. |
| `deck` | summary, inspect-objects, find, fingerprint, assign-ids, find-text, replace-text | Whole-deck inspection, stable addressing, semantic selectors, revisions, deck-wide text replacement. |
| `compose` | kinds, plan, create, get-spec, update-text | Validated composition specs (14 kinds) → native slides with measured fitting and a semantic map. |
| `diagram` | types, create, inspect, list, update-node, add-node, remove-node, add-edge, remove-edge, relayout | Editable diagrams with glued connectors and node ids. |
| `review` | validate, plan-repair, repair, list-rules | 20 validation rules with evidence; repair plans; safe repair. |
| `preview` | snapshot, contact-sheet, compare-images, clear-cache | PowerPoint-rendered PNGs (MCP image content), wireframes, contact sheets, pixel diffs. |
| `batch` | validate, run, start, status, cancel | Validated multi-step batches with references, revision guard, checkpoints, and background jobs. |
| `data` | preview, create-table, create-chart, bind, refresh, bindings | CSV/XLSX ingestion, bound tables and charts, refresh. |
| `design` | list/get/validate/save/import/export/delete/extract-profile, apply-profile, normalize-typography, list-components, update-components | Design profiles, migration with preview, typography, components. |
| `template` | list-layouts, add-slide, fill-placeholders, import-slides, preview-template | Corporate templates: layouts, slides filled by role, slide import with a report, template preview. |
| `asset` | scan-folder, search, tag-asset, duplicates, inspect, place, replace, set-alt-text, embed-linked, fix-links | Local image catalog and picture operations. |

Existing tools gained actions:

| Tool | New actions |
|---|---|
| `textframe` | get-paragraphs, replace-range, format-range, set-paragraph-format, set-text-frame |
| `table` | set-data, apply-style, set-column-widths, format-numbers, conditional-format, set-range-style |
| `chart` | set-data, set-axis, set-data-labels, set-series-style, set-missing-values, apply-style, get-details |

## Workflows

Agents should start with `capabilities` (action `workflows`) for these recipes. A short version:

1. **Understand a deck.** `presentation open` → `deck summary` (slides, theme, layouts, revision)
   → `deck inspect-objects slide_index=N` or `deck find selector=...` → `review validate`.
2. **Build slides.** `design list-profiles` (the profile `theme` means the deck's own style) →
   `compose kinds` → `compose plan` (warnings, estimated fit) → `compose create` →
   `review validate` → `preview contact-sheet`.
3. **Use the deck's layouts.** `template list-layouts` → `template add-slide layout=... content={...}`.
4. **Data.** `data preview` → `data create-table`/`create-chart` → `data refresh` after the file
   changes.
5. **Edit text everywhere.** `deck find-text` → `deck replace-text` (dry run) →
   `deck replace-text dry_run=false expected_count=N`.
6. **Restyle.** `design apply-profile` (dry run) → apply → `design normalize-typography` →
   `review validate`.
7. **Images.** `asset scan-folder` → `asset search` → `asset place`/`replace` → `asset inspect`.
8. **Many edits.** `batch validate` → `batch run` (with `checkpoint=true` and `expected_revision`).
9. **Finish.** `presentation save-as` to a new path, so the original file is kept, then
   `presentation close`.

## Feature guide

Each subsection follows one section of the requirements. "Native tests" lists the real-COM tests
that were written for the feature. None has been run here; see
[Verification status](#verification-status).

### Inspection and semantic addressing

- `deck summary` returns, in one bounded call:
  - slide size and aspect ratio
  - sections
  - theme name, colors, and fonts
  - masters with their layouts and which layouts are in use
  - document properties
  - one page of slides, each with:
    - id
    - title
    - layout
    - hidden state
    - section
    - notes preview
    - shape count
    - fingerprint

  The deck revision is included when the page covers every slide.
- `deck inspect-objects` lists objects with the following, in a compact or a full detail level,
  with pagination:
  - `slide_id`/`shape_id`, which are stable across reordering
  - persistent `PPTMCP_ID`
  - kind and role (from tags or the placeholder type)
  - geometry, rotation, z-order, and visibility
  - text with font sizes and names
  - autofit state and measured text bounds
  - group membership
  - table, chart, link, connector, and crop metadata
  - pixel size and alt text
- Selectors combine terms such as:
  - `kind:table slide:3`
  - `text:"Q3*" font<12`
  - `appid:kpi-1`
  - `name:"Title 1"`
  - `role:title`
  - `group:Cards`
  - `tag:KEY=VALUE`
  - `component:card`
  - geometry comparisons

  `deck find require_single=true` fails with the list of candidates when more than one object
  matches.
- `deck fingerprint` gives per-slide and deck fingerprints; `expected_revision` refuses stale
  edits. `deck assign-ids` adds persistent ids and repairs ids duplicated by copying (later copies
  get a `~N` suffix).

Native tests: `DeckCommandsTests` (9).

### Composition schema and the semantic map

- Schema `pptmcp.composition/1` has 14 kinds: title, section, executive-summary, comparison,
  cards, kpis, image-text, table, chart, timeline, process, hierarchy, quote, and appendix.
  `compose kinds` returns an example for each kind. Validation errors carry JSON paths.
- `compose plan` returns regions, element boxes, estimated text heights, and warnings without
  touching the deck. `compose create` builds native, editable objects (text boxes, shapes,
  tables, charts, pictures, connectors) and tags every part with `PPTMCP_ID` (`slide-id/key`),
  `PPTMCP_ROLE`, and component identity. It also stores the spec on the first slide
  (`PPTMCP_SPEC`).
- `compose get-spec` returns the stored spec for edits. `create replace=true` re-renders it.
  `update-text` replaces one part's text, keeps its formatting, and reports whether it still fits.

Native tests: `ComposeCommandsTests` (8).

### Layout engine and fitting

- The pure `LayoutEngine` works in points (rounded to 0.1 pt) on a 12-column grid with gutters,
  margins, and title, footer, and source areas from the profile. It provides columns, rows,
  stacks, grids, insets, and centering.
- Fitting policies work per element (`each`), per stack (`stack`), per paragraph
  (`paragraphs`), and per table:
  1. Estimate the fit with `TextEstimator`.
  2. Render.
  3. Measure in PowerPoint.
  4. Shrink proportionally down to `min_font_size`.
  5. Split onto continuation slides. Tables split by the rows that fit.
- Pictures use contain, or cover with a focal point. They are never stretched.

### Design profiles and components

- Schema `pptmcp.design-profile/1`:
  - color tokens
  - fonts (heading, body, mono)
  - type scale
  - line and paragraph spacing
  - spacing tokens
  - margins and reserved areas
  - grid
  - minimum font size
  - table, chart, footer, and source styles
  - versioned component styles (card, kpi, badge, callout, legend, timeline, section-header,
    quote, process-step, node)

  Profiles inherit from a `base`, token references resolve, and unknown properties are reported.
- The `theme` profile is derived from the open deck. The built-in profiles are `default`,
  `corporate-blue`, `high-contrast`, and `minimal-mono`. Saved profiles live in
  `PPTMCP_PROFILES_DIR` or `%LOCALAPPDATA%\PowerPointMcp\profiles`.
- `design` actions:
  - **list, get, validate, save, import, export, delete:** manage profiles. Built-in profiles are
    protected.
  - **extract-profile:** drafts a profile from the deck, with a note on how each value was
    inferred.
  - **apply-profile:** produces a migration preview, then applies it:
    - Colors that equal a source-profile token are remapped to the target's value for that
      token. The source is `from_profile`, else the slide's `PPTMCP_PROFILE` tag, else the
      current theme.
    - Theme-bound colors are left to follow the theme update.
    - Fonts and type-scale sizes migrate by role.
    - Tables and charts are restyled.
    - The theme colors and fonts of every master are written.
    - Component tags move to the profile's versions.
  - **normalize-typography:** unifies fonts by role and keeps symbol fonts. It sets the title
    and subtitle sizes and snaps other sizes to the type scale within a tolerance. It raises text
    below the minimum size, and reports sizes that are far off-scale instead of forcing them.
  - **list-components / update-components:** show component instances and their versions, and
    bring outdated ones to the profile's version.

Native tests: `DesignCommandsTests` (5).

### Templates and reuse

- `template list-layouts` lists every master's layouts with their placeholders (role, type,
  name, bounds, in reading order) and how many slides use each layout.
- `template add-slide` finds a layout by name (exact match ignoring case, else a unique partial
  match, else an error listing the choices). It fills placeholders by role: `title`, `subtitle`,
  `body`, `body2`… with indent levels, `picture` (local file, cover in picture placeholders,
  contain in content placeholders), and `notes`. Unmatched keys and empty placeholders are
  reported. `fill-placeholders` does the same for an existing slide.
- `template import-slides` copies `"1-3,5"` from another deck. It reports each slide's source and
  new layout, fonts that are not installed, and validation findings.
- `template preview-template` opens the template read-only and windowless. It maps each layout
  in use to the template (matched, similar, or missing) and lists theme color, font, and
  slide-size differences. It never opens the session's own file twice and closes only what it
  opened.

Native tests: `TemplateCommandsTests` (7).

### Rich text

- `textframe get-paragraphs` returns paragraphs (alignment, level, spacing, line spacing in lines
  or points, indents, bullet or numbering) and runs (font, size, bold, italic, underline,
  `#RRGGBB` color, hyperlink, character spacing, baseline offset). It also returns the frame's
  margins, anchor, wrap, autofit, orientation, columns, and rotation.
- `replace-range` and `format-range` select text by start and length, by the Nth (or last) match,
  or by paragraphs. Replaced text takes the formatting of the first replaced character, so other
  runs keep theirs. Formatting options:
  - hyperlinks: URLs, `mailto:`, files, and `slide:N` links to other slides
  - character spacing
  - superscript and subscript
- `set-paragraph-format` sets alignment, spacing, levels, indents, bullets, and numbering styles.
  `set-text-frame` sets margins, vertical anchor, wrap, autofit, text orientation, columns, and
  rotation.
- `deck find-text` and `deck replace-text` search:
  - every text frame
  - group members
  - table cells
  - speaker notes, with `include_notes`

  Matching is literal (Unicode case folding, so Cyrillic works) or .NET regex with `$1`
  replacements, and a newline in a literal query matches a paragraph or line break. Replacement
  is a dry run by default and lists each match with context. `expected_count` refuses to change
  anything when the count differs. Each hit's text is re-checked before it is written.

Native tests: `RichTextCommandsTests` (8).

### Tables and charts

- Table actions:
  - set-data with sizing
  - styles from profiles
  - column widths
  - number formats, using a subset of Excel format codes
  - conditional formatting by rules
  - range styles
- Chart actions:
  - set-data, writing the embedded workbook so the chart stays editable in PowerPoint
  - axis scale, format, and title
  - data labels
  - series colors, lines, and markers
  - handling of missing values
  - profile styling
  - details: series, values, axes, and labels as text

### CSV and XLSX

- CSV handling:
  - encoding detection (UTF-8 with or without BOM, UTF-16, windows-1251, and others through
    code pages)
  - delimiter detection that does not mistake decimal commas for delimiters
  - culture-aware numbers, percentages, and dates (`ru-RU` handles `1 250,5` and `27,8%`)
  - mixed columns stay text and are reported
- XLSX is read without Excel from the Open XML parts. The reader handles shared and inline
  strings, number formats (dates and percentages), the 1900/1904 date systems, and cached
  formula values.
- Tables and charts remember their source in `PPTMCP_BINDING`. `data refresh` re-reads the file
  (hash-checked), updates cell text and chart data in place, and keeps formatting.

Native tests: `DataCommandsTests` (6).

### Diagrams

- Schema `pptmcp.diagram/1` supports flowchart, swimlane, matrix, hub-spoke, and architecture.
  The layered layout handles back edges (cycles) and reduces crossings. Connectors are glued to
  their nodes, so they follow when nodes move.
- Node and edge edits change shapes in place. `relayout` recomputes positions from the stored
  model. `inspect` reports which connectors are glued.

Native tests: `DiagramCommandsTests` (6).

### Images and assets

- **Catalog** (`PPTMCP_ASSETS_DIR` or `%LOCALAPPDATA%\PowerPointMcp\assets`):
  - `scan-folder` reads headers for PNG, JPEG (with EXIF orientation), GIF, BMP, TIFF, WebP, SVG,
    EMF, and WMF, and hashes each file. Rescans are incremental.
  - `search` ranks words in file names, tags, descriptions, and folders (any script), and filters
    by orientation, minimum width, and tag.
  - `tag-asset` sets tags, a description (also the default alt text), and attribution.
  - `duplicates` groups identical files and repeated placements.
- **Deck pictures:**
  - `inspect` reports effective PPI (from the uncropped picture size), distortion, crop, alt
    text, decorative flag, attribution, and linked or broken files, and flags issues.
  - `place` uses contain, or cover with a focal point. It defaults to the content area and never
    stretches.
  - `replace` puts the new image in the same frame and keeps z-order, name, tags, and alt text.
    Lost animations are reported.
  - `set-alt-text` sets alt text, the PowerPoint Decorative flag where available plus the
    `PPTMCP_DECORATIVE` tag, and attribution.
  - `embed-linked` keeps the crop.
  - `fix-links` relinks by unique file name under a local folder.
- URLs are refused. Nothing is downloaded or generated.

Native tests: `AssetCommandsTests` (4).

### Preview and validation

- `preview snapshot` renders a slide with PowerPoint into a cache. The cache is keyed by session,
  slide fingerprint, and a change counter that the service bumps after every non-read-only
  command. The PNG is returned as MCP image content (up to 4 MB) unless `include_image=false` or
  `PPTMCP_PREVIEW_IMAGES=off`. The result also includes an ASCII wireframe and the slide's
  findings.
- `contact-sheet` tiles slides. `compare-images` reports the share of changed pixels and the
  changed area in pixels and points.
- `review validate` runs 20 rules, each with a certainty and a detection method:
  - placement: `off-slide`, `outside-safe-area`
  - text: `text-overflow`, `small-text`, `missing-font`
  - overlap: `text-overlap`, `partial-overlap`
  - readability: `low-contrast` (WCAG ratio)
  - pictures: `image-distortion`, `missing-asset` (broken links), `missing-alt-text`
  - structure: `empty-placeholder`, `missing-title`, `duplicate-title`
  - consistency: `inconsistent-title-position`, `inconsistent-margins`, `near-misaligned`,
    `footer-inconsistent`, `duplicate-footer`
  - identity: `duplicate-app-id`

  Each finding carries evidence and a suggestion. Picture resolution (PPI) is reported by
  `asset inspect`, and reading order by the existing `accessibility` tool.

Native tests: `PreviewCommandsTests` (4), `ReviewCommandsTests` (7).

### Repair

`review plan-repair` lists the actions for repairable findings and the reason when a finding
cannot be repaired. `review repair` applies only planned fixes:

- moving objects back inside the slide
- measured font scaling that keeps the proportions between runs
- expanding or shrinking boxes
- aligning near-misaligned edges and inconsistent titles
- assigning ids

It validates again afterwards and accepts `expected_revision`.

### Batch, revisions, and recovery

- A batch is a JSON array of `{id, tool, action, args}`. Before anything runs, every operation is
  validated against the command catalog, and arguments are probed with placeholder values for
  references (`"$op1.shapeIndex"`).
- `run` executes the operations in order through the normal service dispatch. Its options:
  - `mode=stop` or `mode=continue` on failure
  - an optional checkpoint copy (SaveCopyAs) with restore steps
  - a change log comparing the deck before and after
  - an `expected_revision` guard
- `start`, `status`, and `cancel` handle background jobs. While a job runs, other write commands
  to that session get a SessionBusy error; read-only commands are allowed.

### Agent usability and capability discovery

- The `capabilities` tool is described above. The server instructions give the workflow in a few
  lines. Every tool takes an `action` enum. Read-only work can use the `_read` tools, which
  carry the MCP read-only hint.
- `PPTMCP_LENIENT_ARGUMENTS=1` accepts near-miss argument names (camelCase, aliases) and action
  spellings, for small local models. In the default strict mode, errors name the canonical
  argument.
- Errors are actionable. They list valid ranges, candidate objects, layout names, and the
  argument to change.

### Local UI

- The UI is off by default. With `PPTMCP_UI=1`, the MCP server (or the CLI daemon) serves a small
  review page on `http://localhost:<port>/`. `PPTMCP_UI_PORT` fixes the port; otherwise a free
  port is chosen. The URL with a token is printed to stderr and returned by `capabilities`.
- Views:
  - slides with PowerPoint-rendered previews and wireframes
  - review findings
  - pictures
  - design profiles and components
  - capabilities
- Security:
  - only loopback clients are accepted
  - the Host header must name localhost on the server's port (protection against DNS
    rebinding)
  - any Origin other than its own is rejected, and no CORS headers are sent
  - a per-run token is delivered in the URL fragment and sent as a header
  - only read-only commands are allowed
  - only files in the preview folder are served
  - HTML, CSS, and JS are embedded, with no CDN
  - a strict Content-Security-Policy is set

  These rules are covered by unit tests that run a real listener.

### COM reliability and distribution

- Unchanged and relied on:
  - the STA queue
  - the message filter
  - owned-process cleanup that verifies the PID and the process start time
  - closing without killing unrelated PowerPoint processes
- New code:
  - releases every COM object in `finally` blocks
  - avoids nested `Execute` calls, because composed operations make sequential Core calls
  - never opens the session's own file twice
- `scripts/Publish-OfflineBundle.ps1` builds the offline bundle. It produces self-contained
  single-file executables with the release workflow's settings, never trimmed. The bundle also
  contains demos, docs, client configuration samples, install and verify scripts, and SHA-256
  sums. See [OFFLINE-INSTALL.md](OFFLINE-INSTALL.md) and
  [CLIENT-CONFIGURATION.md](CLIENT-CONFIGURATION.md).
- Target frameworks, trimming, and AOT settings were not changed.

### Demonstrations

The `demos/` folder has ten PowerShell scenarios, each run through `pptcli`:

1. quarterly review from a Russian CSV
2. restyle to a custom profile
3. template workflow
4. diagrams
5. Russian-language deck
6. overflow and repair
7. deck-wide renaming
8. images and assets
9. batch, revisions, and jobs
10. review and accessibility

`Run-Demos.ps1` records step timings, failures, slide counts, findings, decks, and contact sheets
in `results.json` and `results.md`. The specs, profile, and CSV that the scenarios use are checked
by `DemoInputsTests`.

## Schemas and persistent tags

| Schema | Where |
|---|---|
| `pptmcp.composition/1` | `compose kinds` returns an example per kind; `Composition/CompositionSpec.cs` |
| `pptmcp.design-profile/1` | `design get-profile profile=default`; `Design/DesignProfile.cs` |
| `pptmcp.diagram/1` | `diagram types` returns examples; `Diagram/DiagramSpec.cs` |
| Batch operations | `batch validate`; `Batch/BatchPlan.cs` |
| Placeholder content | `template add-slide`; `Template/PlaceholderContent.cs` |

The following tags are written on slides and shapes. They are ordinary PowerPoint tags and
survive save, copy, and reorder:

- `PPTMCP_ID`, `PPTMCP_ROLE`, `PPTMCP_COMPONENT` (`name@version`), `PPTMCP_PROFILE`
- `PPTMCP_COMPOSITION`, `PPTMCP_SPEC`, `PPTMCP_PART`, `PPTMCP_META_*`
- `PPTMCP_BINDING`
- `PPTMCP_DIAGRAM*`, `PPTMCP_NODE`, `PPTMCP_EDGE*`, `PPTMCP_LANE`, `PPTMCP_LAYER`, `PPTMCP_QUADRANT`
- `PPTMCP_IMG_PX`, `PPTMCP_IMG_SOURCE`, `PPTMCP_ATTRIBUTION`, `PPTMCP_DECORATIVE`
- `PPTMCP_ALLOW_OVERLAP`

## Safety, offline operation, and reliability

- No cloud APIs, telemetry, accounts, or runtime downloads. Images are local files only.
- Bulk changes preview by default:
  - `deck replace-text`
  - `design apply-profile`
  - `design normalize-typography`
  - `review repair` (plan-repair shows the plan)
  - `asset embed-linked`, `asset fix-links`
  - `batch` (validate first, checkpoint optional)
- Nothing saves over the original file unless the user saves the session. The documentation and
  the server instructions recommend `presentation save-as` to a new path.
- The server never enables macros or dismisses security prompts. It never terminates PowerPoint
  processes it did not start.

## Verification status

What ran in this environment (Linux container, .NET 10 SDK, no PowerPoint):

| Check | Result |
|---|---|
| Release build of the whole solution (`Sbroenne.PowerPointMcp.slnx`) | passes |
| PowerPoint-free tests (`RequiresPowerPoint!=true`; McpServer.Tests) | **563 / 563 pass**, 450 of them new |
| MCP protocol tests (tool list, schemas, read-only aliases, argument normalization) | pass (included above) |
| Local UI security tests (real HttpListener on loopback) | pass (included above) |
| `check-com-leaks`, `check-dynamic-casts`, `check-success-flag`, `check-core-interface-completeness` | pass |
| `check-doc-counts` (53 tools / 305 operations / 27 domains) | pass |
| CLI `--help` for the new commands | runs |

What was written but **not run** (needs Windows and PowerPoint):

| Item | Count |
|---|---|
| Native Core tests for the new domains (`tests/PowerPointMcp.Core.Tests`) | 64 tests in 10 classes |
| MCP end-to-end workflow (`McpPlatformWorkflowTests`) | 1 long test |
| Demonstration scenarios (`demos/`) | 10 |
| Offline bundle build and install scripts | not run |

Follow [VERIFICATION.md](VERIFICATION.md) on Windows and record the results there. Expect some
native tests to need fixes: the code follows the PowerPoint object model closely, but COM
behaviour often differs from documentation in edge cases (for example, `TextRange2` indexing,
`Crop` on picture placeholders, `Decorative` availability, and font names that come from the
theme).

## Known limitations

- **Native behaviour is unverified.** See above.
- `NumberFormatter` covers a subset of Excel number formats. It has no thousands scaling with a
  trailing comma, no conditional sections, and no fractions.
- Theme-bound fonts are detected by name. With `update_theme`, runs that use the current theme
  font are assumed to be theme-bound and are left to follow the theme.
- Charts, SmartArt, and OLE objects are not searched by `deck find-text`.
- Checkpoint restore is manual: open the checkpoint copy. The result lists the steps.
- The preview cache is invalidated by changes made through the server. After editing in the
  PowerPoint window, pass `refresh=true`.
- `asset replace` and `asset embed-linked` re-create the picture, so animations on it are lost
  (this is reported). Pictures inside groups must be ungrouped first.
- The local UI is read-only by design.
- The CLI does not expose the `capabilities` tool. Use `--help`, or the MCP tool.

## Commits

On `feature/authoring-platform`, after the synthetic upstream sync `b6012a0`:

| Commit | Content |
|---|---|
| c5db9ac | Near-miss argument normalization, actionable errors, `table set-data` |
| 8dcd94f | `deck`: inspection, addressing, selectors, fingerprints, persistent ids |
| a4212b1 | `review`: validation with evidence, repair plans, measured repair |
| 66b6da5 | `compose`: design profiles, layout engine, compositions |
| 57e2344 | `diagram`: editable diagrams with glued connectors |
| ddb248e | `preview`: rendered snapshots, contact sheets, image diffs |
| 5bc3f3e | `batch`: validated batches, jobs, checkpoints, change log |
| 664edb5 | `data`: CSV/XLSX, refresh; advanced table and chart actions |
| e57ec7a | Deck-wide find/replace and rich text editing |
| 5d5faff | `design` apply/typography/components; `template` workflows |
| 031d955 | `asset`, `capabilities`, local UI |
| (this) | Documentation, demos, offline bundle, workflow test, counts |
