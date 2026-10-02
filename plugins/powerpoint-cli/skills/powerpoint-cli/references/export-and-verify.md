> **CLI syntax:** Shared guides may use MCP calls as shorthand. Use `cli-commands.md` or live `--help` for exact commands and kebab-case options.

# Export & Visual Verification

Reference for PowerPoint's native PDF delivery and image rendering actions. The image actions
provide the multimodal "look at the result" verification loop. Single-property reads
(`textframe(action: "get-text", ...)`, `shape(action: "get-count", ...)`,
`chart(action: "get-chart-data", ...)`) cannot catch layout problems. `slide(action:
"check-layout", ...)` catches the measurable ones (overflow, overlap, off-slide, tiny text, empty
placeholders) without a vision model; a rendered image catches the rest (colors, balance, chart
proportions).

## REQUIRED: Verify After Visual Changes

**You MUST export and look at the result after creating or repositioning any visual content** —
shapes, tables, charts, images, or significant text/formatting changes. Do not save and close a
session without this step when the task involves visual output.

```
1. chart(action: "add-chart", ...) / table(action: "add-table", ...) / image(action: "add-picture", ...) / shape(action: "set-position", ...)
2. slide(action: "check-layout", session_id: ..., slide_index: ...)                            ← fix every error/warning first
3. export(action: "export-slide-to-image", session_id: ..., slide_index: ..., output_path: ...)  ← REQUIRED — never skip
4. Inspect the returned image for anything check-layout cannot judge (balance, colors, charts)
5. If issues found → fix → check-layout / export again → repeat until it looks right
6. presentation(action: "close", sessionId: ..., save: true)
```

This rule applies even if the operation reported `success: true` — a successful COM call only
confirms the API accepted the parameters, not that the result looks correct.

## Deterministic Layout Check (No Vision Model Needed)

`slide(action: "check-layout", session_id: ..., slide_index: ...)` measures the slide through
PowerPoint itself and reports problems with the exact shapes and numbers involved. Omit
`slide_index` to check every slide at once. It is read-only and cheap; run it after every batch of
edits, before exporting images.

| `code` | `severity` | Meaning | Typical fix (also given in `suggestion`) |
|--------|------------|---------|------------------------------------------|
| `text-overflow` | error | The rendered text is taller or wider than its shape | `shape(action: "set-size", ...)` to the suggested size, shorten the text, or lower the font size |
| `off-slide` | error / warning | The shape is entirely (error) or partly (warning) outside the slide | `shape(action: "set-position", ...)` to the suggested `left`/`top`, or `set-size` |
| `text-overlap` | error | Two shapes with text (or tables) overlap | Move the lower shape to the suggested `top`, or place them side by side |
| `partial-overlap` | warning | Shapes overlap without one containing the other | Separate them, or put one fully inside the other if layering is intended |
| `small-text` | warning | Text below 10 pt | `textframe(action: "set-font-size", ...)` to 12 pt or more, or move detail to speaker notes |
| `empty-placeholder` | warning | A content placeholder has no text, picture, table, or chart | `shape(action: "set-placeholder-text", ...)` or `shape(action: "delete", ...)` |
| `near-misaligned` | info | Left or top edges differ by a few points | `shape(action: "set-position", ...)` or `shape(action: "align", ...)` |

`layoutOk` is `true` when no error or warning remains (info items are optional polish). One shape
fully inside another (text on a card, a caption on a photo) and full-slide background shapes are
treated as intentional and not reported.

Use `slide(action: "inspect", ...)` when you need the full picture: every shape's index, kind,
text, font sizes, position, size, and rendered text bounds in one call.

check-layout does not judge colors, contrast, visual balance, chart internals, or whether the
content makes sense — that is what the exported image is for.

## Actions

| Tool | Action | Parameters | Notes |
|------|--------|------------|-------|
| `export` | `export-to-pdf` | `session_id`, `output_path`, `overwrite` (default `false`) | Creates a PDF without changing the open presentation's file path. Refuses to replace an existing file unless `overwrite` is explicitly true. |
| `export` | `export-slide-to-image` | `session_id`, `slide_index`, `output_path`, `format` (default `"PNG"`), `width`, `height` (optional pixels) | Renders exactly one slide to a single image file. |
| `export` | `export-all-slides-to-images` | `session_id`, `output_directory`, `format` (default `"PNG"`) | Renders every slide; PowerPoint names files `Slide1.PNG`, `Slide2.PNG`, etc. in the given directory. |

- `format` accepts any PowerPoint export filter name: `"PNG"`, `"JPG"`, `"GIF"`, `"BMP"`, `"TIF"`,
  `"WMF"`, `"EMF"`. Default to `"PNG"` unless the user needs a specific format.
- `width`/`height` on `export-slide-to-image` control output pixel dimensions; omit them to use
  PowerPoint's default rendering size.
- `export-all-slides-to-images` creates `output_directory` if it doesn't already exist.

## When to Use Each

| Situation | Use |
|-----------|-----|
| Delivering a finished deck as a PDF | `export(action: "export-to-pdf", ...)` after saving the presentation |
| Just added/changed one slide | `export(action: "export-slide-to-image", ...)` on that slide only |
| Finished building a whole deck | `export(action: "export-all-slides-to-images", ...)` once, review every image |
| Iterating on a single slide's layout | `export-slide-to-image` repeatedly on that slide during the fix loop |

Prefer the single-slide export while iterating on one slide — exporting the whole deck on every
fix cycle wastes calls once you've localized the issue to one slide.

## Common Issues to Look For

| Problem | Likely fix |
|---------|-----------|
| Text visually cut off / overflowing its box | Shorten `text`, reduce `font_size`, or grow the shape with `shape(action: "set-size", ...)` |
| Two shapes overlapping | `shape(action: "set-position", ...)` on one of them, using the positioning reference in `deck-builder.md` |
| Chart illegible / too small | Increase `width`/`height` on the chart, or reduce category count |
| Table cell text overflowing | Shorten cell text or increase the table's `height` when creating it (tables can't be resized post-creation without re-adding — see `tables.md`) |
| Image stretched/squashed | Recompute `width`/`height` to match the source image's aspect ratio |

## The Fix Loop

```
For each slide with visual content {
  1. Build/modify the slide
  2. slide(action: "check-layout", ...) → apply each suggestion → repeat until layoutOk is true
  3. export(action: "export-slide-to-image", ...) → inspect
  4. If ANY issue found → fix it → check-layout and export again
  5. Move to the next slide only when it looks right
}
```

Expect 1-2 fix cycles per visually complex slide (charts, tables, multi-shape layouts) — this is
normal, not a sign something went wrong the first time.

## After the Full Deck

Before the final `presentation(action: "close", sessionId: ..., save: true)`, run
`slide(action: "check-layout", session_id: ...)` without `slide_index` to check every slide, then
`export(action: "export-all-slides-to-images", ...)` once as a final pass over the whole deck,
confirming no slide was missed and the deck reads coherently start to finish.
