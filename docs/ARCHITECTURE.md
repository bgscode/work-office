# Architecture

## Layers

```
apps/web                      React UI (ribbon, formula bar, canvas grid)
   │  imports
packages/engine               @office/engine: loads the WebAssembly module
   │  calls (wasm-bindgen)
crates/wasm                   thin bindings: type conversion only
   │
crates/sheets-engine          cells, formulas, functions, recalculation
crates/formats                .xlsx / .csv  (only crate that knows files)
crates/docs-engine            Docs model (early)
crates/slides-engine          Slides model (early)
   │
crates/core                   NodeId, Op trait, undo/redo History
crates/collab                 real-time co-editing (planned)
services/server               auth, storage, collaboration (planned)
```

The desktop app (`apps/desktop`, planned) will wrap `apps/web` with Tauri and
call the same Rust crates natively instead of through WebAssembly.

## App shell

`apps/web/src/App.tsx` picks the page from the URL hash (`router.ts`): `#/` is the Home page (`home/HomePage.tsx`), `#/sheets/<document id>` opens a workbook. `documents.ts` stores every document's bytes (`.xlsx` for Sheets) and a Recent index in `localStorage`; Home lists them and the app autosaves into them. The title bar (`chrome/TitleBar.tsx`) belongs to the shell; each app fills its middle through a portal (`chrome/titleSlot.ts`).

## Saving and opening

`sheets/SheetsLoader.tsx` opens a workbook (from IndexedDB or a picked file) behind an "Opening…" screen. The document is wrapped by `sheets/journal.ts`, which records every call that changes it; autosave stores that journal in IndexedDB after each edit and rewrites the full `.xlsx` only when it is cheap or after a pause. Opening replays the journal over the last full save. Numbers and how to benchmark: [PERFORMANCE.md](PERFORMANCE.md).

## What happens when you type `=SUM(A1:A3)` and press Enter

1. `Grid.tsx` (the focused cell input) reports the draft; `SheetsApp.tsx` commits it with
   `doc.setInputs(sheet, rows, cols, inputs)`.
2. `crates/wasm` wraps the edit in a `SetCells` op and applies it through
   `History`, which records the inverse for undo.
3. `Sheet::set_inputs` parses the formula, updates the dependency graph and
   recalculates only the affected cells.
4. React bumps a revision counter; the grid repaints visible cells by asking
   the engine for `displayBlock` / `kindBlock` / `styleBlock` of the viewport
   (`apps/web/src/sheets/paint.ts`).

## Data tools

AutoFilter (`filter.rs`), Data Validation (`validation.rs`) and Conditional
Formatting (`conditional_format.rs`) live in `Sheet::data_tools()` and are
changed as a whole with `SetDataTools`. A filter hides rows only when it is
applied (as in Excel); those rows join the user's hidden rows in
`Sheet::is_row_hidden`, so `SUBTOTAL`, the grid and the row layout all see
them. Conditional formats are evaluated per painted block (`cfBlock`); the
range statistics they need (ranks, averages, duplicates) are computed once per
change and cached in `crates/wasm`. The UI is in
`apps/web/src/sheets/features/dataTools.tsx`; `.xlsx` reading and writing in
`crates/formats/src/xlsx_data_tools.rs`, within the single worksheet pass.

## Charts and pictures

Drawings live in `Sheet::drawings()` and are anchored to cells (a cell plus a
pixel offset for each corner), so they follow column widths, row heights and
inserted or deleted lines. A chart stores references to its data, never the
numbers: `chart_view` reads the cells each time the chart is drawn, so it
updates live. The UI turns anchors into pixels with the grid's layout
(`apps/web/src/sheets/charts/DrawingLayer.tsx`) and draws charts on a canvas
(`charts/renderChart.ts`). Shapes (`shape.rs`) are DrawingML preset geometries drawn by `charts/renderShape.ts`. `.xlsx` drawings, charts and media are read and
written by `crates/formats/src/xlsx/drawings.rs`.

## Printing

`apps/web/src/sheets/print/`: `printJob.ts` collects what to print from the
workbook (selection, print area or used range, sizes from the grid layout,
formatted cells, merges, charts rendered to images); `pagination.ts` splits it
into pages like Excel (down, then over; print titles repeated); `printHtml.ts`
writes one HTML page per sheet of paper, used both by the preview
(`PrintView.tsx`) and by the real print in a hidden frame — so the preview is
exactly what prints, and the browser's "Save as PDF" gives PDF export. Page
setup lives in the engine (`page_setup.rs`, `SetPageSetup` op) and is saved in
`.xlsx` by `xlsx_page_setup.rs`.

## Dynamic arrays

`formula/arrays.rs` evaluates a formula to a value or an array (`eval_at_any`).
An array result is stored by `Sheet::store_result`: the formula cell keeps the
top-left value and the rest of the block is "spilled" — kept with the formula
(`spills`), with each covered cell pointing at it (`spill_owner`) so
`Sheet::value` reads it like any cell. A block that would overwrite data shows
`#SPILL!` and is remembered (`spill_blocked`) so typing or clearing in it
recalculates the formula. Formulas reading spilled cells count as dependents of
the spilling formula; when a spill changes shape, its readers are recalculated in
another round.

## Defined names

`Workbook::names` (`names.rs`) holds each name with its parsed `refers_to`.
Formulas resolve a name through `CellSource::name_expr` (LET variables first).
Formulas using names are listed per sheet (`name_users`) and indexed like
cross-sheet links, so changing a named range's cells recalculates them; `SetNames`
(undoable) recalculates them all. Sheet renames, deletes and row/column
inserts rewrite `refers_to` with the same `refs.rs` helpers formulas use.

## Kinds of edits

| Edit | Engine op | Undo |
|---|---|---|
| Typing, paste, clear | `SetCells` | previous inputs |
| Formatting | `SetStyles` | previous styles |
| Fill handle, paste within Office | `Batch(SetCells, SetStyles)` | both |
| Column width | `SetColWidths` | previous widths |
| Insert/delete rows or columns, sort | `ReplaceSheet` (formulas rewritten by `formula/refs.rs`) | the old sheet |
| Filter, data validation, conditional formatting | `SetDataTools` (`data_tools.rs`) | previous settings |
| Insert, move, resize, format or delete a chart or picture | `SetDrawings` (`chart.rs`) | previous drawings |
| Page setup (Page Layout tab, Page Setup dialog) | `SetPageSetup` (`page_setup.rs`) | previous settings |
| Remove duplicates, custom sort, sort from a filter | `ReplaceSheet` | the old sheet |
| Add/rename/delete sheet | `InsertSheet`, `RenameSheet`, `RemoveSheet` | the inverse op |

## Design rules

1. **Logic in Rust, presentation in TypeScript.** Anything testable without a browser belongs in a crate.
2. **Every document edit is an `Op`.** Undo/redo today, collaboration and autosave later, all build on the same stream.
3. **Dependencies point downward only.** Engines never import UI code or each other; `formats` is the only crate that knows several engines.
4. **Canvas, not DOM, for document content.** DOM rendering doesn't scale to a million rows; the grid paints only what is visible.
5. **The engine owns the data.** React state holds only UI concerns (selection, the edit in progress).
6. **Docs change with code.** See [CONTRIBUTING.md](../CONTRIBUTING.md#before-you-open-a-pull-request).

## Why Rust + TypeScript

| Need | Choice | Reason |
|---|---|---|
| Heavy computation (recalculation, layout, file parsing) | Rust | C++-level speed with memory safety |
| Run on web and desktop | Rust → WebAssembly + Tauri | One engine everywhere |
| Rich UI | React + TypeScript | Largest ecosystem, easy for contributors |
| Co-editing (planned) | CRDT (`yrs` or Loro) | Proven conflict-free merging |

LibreOffice and ONLYOFFICE are written in C++; Rust gives similar performance
with code that is safer for many contributors to change.
