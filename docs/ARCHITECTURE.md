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

## Kinds of edits

| Edit | Engine op | Undo |
|---|---|---|
| Typing, paste, clear | `SetCells` | previous inputs |
| Formatting | `SetStyles` | previous styles |
| Fill handle, paste within Office | `Batch(SetCells, SetStyles)` | both |
| Column width | `SetColWidths` | previous widths |
| Insert/delete rows or columns, sort | `ReplaceSheet` (formulas rewritten by `formula/refs.rs`) | the old sheet |
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
