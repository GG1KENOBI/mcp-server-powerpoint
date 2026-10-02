---
"powerpointmcp": minor
---

Add `slide inspect`, which returns every shape on a slide in one call (index, kind, placeholder type, position, size, rotation, z-order, text, font-size range, autofit mode, rendered text bounds, and table size), and `slide check-layout`, which finds text overflow, off-slide shapes, overlapping text, partial overlaps, text below 10 pt, empty placeholders, and near-misaligned edges on one slide or the whole deck, each with a concrete fix. Agents can now verify and repair layouts without a vision model.

Add `table set-data` to fill a whole table from rows of separated cells (pasted markdown tables work) in one call instead of one call per cell. Data that does not fit is rejected before anything is written.
