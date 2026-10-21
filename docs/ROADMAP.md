# Roadmap

Sheets ships first: it is the most self-contained app and the quickest to make
genuinely useful. Items marked ✅ are done.

## Phase 0 — Foundation ✅

- ✅ Monorepo: Cargo workspace + pnpm workspace
- ✅ Core `Op` trait and undo/redo history
- ✅ Starter models for Docs and Slides

## Phase 1 — Sheets MVP (in progress)

The detailed, feature-by-feature list (with priorities and milestones M1–M7) is
[docs/sheets/EXCEL_FEATURES.md](sheets/EXCEL_FEATURES.md).

- ✅ Formula engine: parser with Excel precedence, 46 functions, error values
- ✅ Dependency graph with incremental recalculation and cycle detection
- ✅ Canvas grid with virtual scrolling, selection, keyboard navigation
- ✅ Formula bar, multiple sheets, status bar
- ✅ Copy/paste (tab-separated), undo/redo
- ✅ `.xlsx` and `.csv` import/export
- ✅ IME-friendly typing
- ✅ Cell formatting: bold/italic/underline/strike, colours, alignment
- ✅ Number formats (₹ with Indian grouping, $, %, dates, times), auto-detected from typing
- ✅ Column resizing and autofit; insert/delete rows and columns with reference rewriting
- ✅ Fill handle (series, day/month names, formulas) and relative-reference copy/paste
- ✅ Formula autocomplete, syntax hints, point-and-click references with coloured highlights
- ✅ Find, sort, right-click menu, sheet rename/delete
- ✅ 127 functions including `XLOOKUP`, `SUMIFS`, statistics, text, date & time
- ✅ Autosave to browser storage
- ⬜ Borders, font family/size, merged cells, wrap text, row heights
- ⬜ Cross-sheet references (`Sheet2!A1`), whole-column references (`A:A`), named ranges
- ⬜ Replace, filter, freeze panes, conditional formatting, data validation
- ⬜ Financial functions, dynamic arrays (`FILTER`, `SORT`, `UNIQUE`)
- ⬜ Charts
- ⬜ Volatile functions (`TODAY`, `RAND`) refreshing on every recalculation

## Phase 2 — Docs MVP

- Text shaping (`rustybuzz`), line breaking, pagination
- Canvas editor: caret, selection, IME
- Styles, lists, tables, images, headers/footers
- `.docx` import/export, PDF export

## Phase 3 — Slides MVP

- Shape editor (move/resize/rotate), text boxes, images
- Themes, layouts, presenter mode
- `.pptx` import/export

## Phase 4 — Desktop and cloud

- Tauri desktop app (macOS, Windows, Linux)
- Server: accounts, file storage, sharing
- Real-time co-editing (CRDT), comments, version history

## Phase 5 — Advanced

- Charts and pivot tables, track changes, mail merge, scripting, AI assistant, mobile
