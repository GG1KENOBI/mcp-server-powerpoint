# Capability Matrix — Agent Authoring Platform

Base commit: `cd8af40` (upstream `main`, v0.3.2). Built from the source (Core interfaces, generated
manifest, `PresentationTools.cs`, service dispatch), not from README counts. Delivery status is in the
last section and in the delivery report (`docs/AUTHORING-PLATFORM.md`).

Legend for verification: **U** = unit tests (pure logic, run in CI without PowerPoint),
**P** = protocol tests (MCP in-memory transport, no PowerPoint), **N** = native tests against real
PowerPoint (Windows only), **D** = documented manual procedure.

## Existing surface at the base commit

| Domain | Actions (from generated manifest / hand-written enum) |
|---|---|
| presentation (hand-written) | create, open, close, list, test, save-as, save-copy-as, apply-template, get-theme-name, get/set-final, document and custom properties, tags |
| slide | add-blank, get-count, inspect (bounded overview: name, layout, shape count, text preview), delete, duplicate, move-to, backgrounds, sections, comments, hidden, master shapes, import-from-file, tags |
| shape | 57 actions: add rectangle/text box/WordArt/auto shape/line/connector/attached connector, position/size/fill/line/effects, duplicate/copy, align/distribute, group/ungroup/merge, name/alt text/hyperlink, links, placeholders (list/set text/set image), tags |
| textframe | set/get text, find/replace (one shape), font size/name/color/bold/italic/underline, alignment, bullets, autosize |
| table | add, cell text, rows/columns insert/delete, cell fill/border, merge |
| chart | add (one series), add series, titles, axis titles, legend, replace data, style/color style, data table |
| image | add picture (embed/link), brightness/contrast, recolor, transparency, crop and crop frame, compress |
| layout / master / pagesetup | set (built-in `ppLayout*` only)/get/list/delete layouts; theme colors/fonts, title/body fonts, backgrounds; slide size, footer |
| notes, animation, smartart, media, customshow, accessibility, export | as named; export renders PDF and PNG files; accessibility audit (alt text, empty title, reading order) |

Every generated domain is exposed as a write tool and, for read actions, a `<domain>_read` tool
(32 tools). Sessions use `presentation_session_id`.

## Requested capabilities

| # | Requested capability | Existing implementation / entry points | Missing behavior | Proposed extension | Verify |
|---|---|---|---|---|---|
| 3.1 | Presentation inspection (size, sections, theme, masters, layouts, metadata) | Scattered: `pagesetup get-settings`, `slide get-section-*`, `presentation get-theme-name`, `master list-masters`, `layout list-layouts`, document properties | One bounded call returning all of it | `deck summary` (new `deck` domain) reusing the same COM members | N, P |
| 3.2 | Slide inventory (titles, notes, hidden, objects) | `slide inspect` overview (no title, notes, hidden, ids) | Titles, notes, hidden state, stable ids, pagination offset | `deck summary` slide page with `start_slide`/`max_slides` | N |
| 3.3 | Object inspection (type, bounds, rotation, z-order, visibility, text, formatting, groups, table/chart/picture/connector/placeholder metadata) | Per-property getters; `list-placeholders` | One call, filters, pagination, compact/detailed modes, group members, ids | `deck inspect-objects` with selector filter, `detail=compact|full`, paging | N, U (selector) |
| 3.4 | Stable addressing | Index-based everywhere; `Slide.SlideID`/`Shape.Id` not exposed; tags exist | IDs in results, id-based targeting, persistent app ids, duplicate-tag repair | `slide_id`/`shape_id` in all new outputs; `PPTMCP_ID` tag via `deck assign-ids` with duplicate detection | U, N |
| 3.5 | Semantic selectors and ambiguity errors | None | Selector language, resolution, candidates on ambiguity | Pure `SelectorEngine` (title, tag, name, role, kind, text, group, geometry, font) + `deck find` | U, N |
| 4 | High-level composition schema and 12 slide kinds | Docs-only composition recipes | Validated, versioned spec; executor; semantic map | `compose` domain: `validate`, `slide`, `deck` from JSON spec v1, pure planner + Core-reusing executor | U (planner/validation), N |
| 5 | Layout engine and fitting | None (agents compute coordinates) | Regions, stacks, cards, columns, fitting policies | Pure `LayoutEngine` + `TextFitEstimator`; measured re-fit with PowerPoint `BoundHeight` after placement | U, N |
| 6 | Design-system profiles and components | Theme readers (`master get-theme-*`), `apply-template` | Profile schema, CRUD, apply, extract, components | `design` domain: profile `validate/list/save/export/extract/apply/preview-apply`; component metadata tags | U, N |
| 7 | Templates and cross-presentation reuse | `apply-template`, `slide import-from-file`, `layout list-layouts` (custom layouts not applicable) | Create slides from custom layouts, placeholder role mapping, preview | `compose` layout-based slides via `CustomLayouts` and placeholder roles; `design preview-apply` | N |
| 8 | Rich text and precise editing | One-shape find/replace, whole-frame formatting | Paragraph/run inspection, paragraph formatting, margins/anchor, range hyperlinks, deck find/replace preview | `textframe` new actions + `deck find-replace` (preview by default) | U (replacement planner), N |
| 9 | Advanced tables/charts; CSV/XLSX | Per-cell table text; single-series chart, replace data | Bulk fill, number formats, sizing, header styles, pagination, refresh; axis scale/labels; file ingestion | `table set-data`, `format-*`, `paginate`; `data` domain for CSV/XLSX preview and bound refresh; chart axis/labels | U (parsers/formatters), N |
| 10 | Editable diagrams | Attached connectors, auto shapes, grouping exist | Layouts and node ids | `compose` diagram kinds (process, timeline, hierarchy, hub-spoke, matrix, swimlane, flowchart) using shapes + attached connectors + `PPTMCP_NODE` tags | U (layout), N |
| 11 | Images and assets | add-picture, crop, compress, link info | File metadata, contain/cover/focal placement, catalog, distortion check | Pure `ImageFileInfo` (PNG/JPEG/GIF/BMP headers) + placement math; `asset` catalog in `design` domain | U, N |
| 12 | Preview and visual validation | Export to PNG file; my earlier layout check (fork) | Image content in MCP results, fingerprints, contact sheets, broader findings | `preview` MCP tool (hand-written, returns `ImageContentBlock`) + `deck validate` | P, N |
| 13 | Safe repair | None | Plan + apply for bounds, overflow, alignment | `deck repair` (`dry_run` default true) | U (planner), N |
| 14 | Batch, revisions, recovery | None | Ordered batch, prevalidation, partial-failure report, checkpoints, expected revision, change log | `batch` MCP tool over the shared service; fingerprint from inspection snapshot | P, N |
| 15 | Agent usability, capability discovery | Strict validation, read/write tool split | Near-miss normalization (fork work), capabilities report | Port normalizer; `deck capabilities` (PowerPoint version, operations, limits) | P, N |
| 16 | Local UI | None | Loopback viewer over the same service | Optional `--ui` HttpListener on 127.0.0.1 with embedded assets | P (handler logic), D |
| 17 | COM reliability and offline distribution | STA queue, message filter, owned-process cleanup, untrimmed self-contained publish | Audit notes, offline install docs | Review + docs; no change to working publish settings | D |
| 18 | Demonstrations | None | Scenario decks and renders | Scripted scenarios (`demos/`) runnable on Windows through the CLI/MCP; results recorded only when executed | D |

## Delivery status

Delivered on `feature/authoring-platform`. **Implemented** means the code is in place and its
pure logic is unit-tested. **Native** gives the real-PowerPoint tests that were written. None of
them has been run yet, because the work was done without Windows; see
[VERIFICATION.md](VERIFICATION.md).

| # | Delivered as | Unit tests (run, pass) | Native tests (written, not run) |
|---|---|---|---|
| 3.1–3.5 | `deck` summary, inspect-objects, find (selectors), fingerprint, assign-ids | Deck (21) | DeckCommandsTests (9) |
| 4 | `compose` with 14 kinds, plan, create, get-spec, update-text | Composition (45) | ComposeCommandsTests (8) |
| 5 | LayoutEngine, TextEstimator, measured fitting and splitting in the renderer | Composition (included above) | ComposeCommandsTests |
| 6 | `design` profiles, apply with migration preview, typography, components | Design (6) + Composition profile tests | DesignCommandsTests (5) |
| 7 | `template` list-layouts, add-slide, fill-placeholders, import-slides, preview-template | Template (5) | TemplateCommandsTests (7) |
| 8 | `textframe` get-paragraphs, replace-range, format-range, set-paragraph-format, set-text-frame; `deck` find-text and replace-text | Text (14) | RichTextCommandsTests (8) |
| 9 | `table` and `chart` advanced actions; `data` CSV/XLSX with bound refresh | Data (10) + table parser tests | DataCommandsTests (6) |
| 10 | `diagram` with five types, glued connectors, edits, relayout | Diagram (11) | DiagramCommandsTests (6) |
| 11 | `asset` catalog, inspect, place, replace, alt text, embed, fix-links | Assets (16) | AssetCommandsTests (4) |
| 12 | `preview` snapshot (MCP image content), contact sheet, image diff; `review validate` | Preview (6), Review (45) | PreviewCommandsTests (4), ReviewCommandsTests (7) |
| 13 | `review` plan-repair and repair | Review (included above) | ReviewCommandsTests |
| 14 | `batch` validate, run, start, status, cancel; fingerprints, checkpoints | Batch (13) | covered by McpPlatformWorkflowTests and demo 9 |
| 15 | `capabilities` tool, workflow instructions, lenient arguments | protocol and normalization tests, Assets (capabilities) | McpPlatformWorkflowTests |
| 16 | Local UI (`PPTMCP_UI=1`) on the shared service | LocalUi (5, real loopback listener) | manual steps in VERIFICATION.md |
| 17 | Offline bundle, install and verify scripts; COM rules kept | — | VERIFICATION.md step 5 |
| 18 | Ten demo scenarios with a runner that records results | DemoInputs (5: specs, profile, CSV) | demos/Run-Demos.ps1 |

