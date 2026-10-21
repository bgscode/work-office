# Contributing

Thanks for helping build an open-source office suite! This guide gets you from
clone to merged pull request.

## Setup

1. Install [Rust](https://rustup.rs) (stable), then `rustup target add wasm32-unknown-unknown`.
2. Install [wasm-pack](https://rustwasm.github.io/wasm-pack/installer/) (`cargo install wasm-pack`).
3. Install Node.js 20+ and [pnpm](https://pnpm.io/installation) 10.
4. `pnpm install`, then `pnpm dev` and open http://localhost:5173.

After you change Rust code, re-run `pnpm build:wasm` (or restart `pnpm dev`) so the browser picks it up.

## Where things live

| Path | What | Language |
|---|---|---|
| `crates/core` | Ids, the `Op` trait, undo/redo `History` | Rust |
| `crates/sheets-engine` | Spreadsheet model, formula parser/evaluator, functions, recalculation | Rust |
| `crates/formats` | `.xlsx` and `.csv` import/export | Rust |
| `crates/wasm` | Bindings exposed to JavaScript | Rust |
| `crates/docs-engine`, `crates/slides-engine` | Early models for Docs and Slides | Rust |
| `packages/engine` | `@office/engine`: loads the WebAssembly build | TypeScript |
| `packages/ui` | Shared React components (ribbon) | TypeScript |
| `apps/web` | The web app; Sheets UI is in `src/sheets/` | TypeScript/React |
| `docs/` | Architecture, roadmap | Markdown |

Rule of thumb: **logic goes in Rust, presentation goes in TypeScript.** If you
can unit-test it without a browser, it belongs in a crate.

## Common tasks

### Add a spreadsheet function

1. Implement it in the matching module under `crates/sheets-engine/src/functions/`
   (`math.rs`, `stats.rs`, `text.rs`, `logical.rs`, `lookup.rs`, `conditional.rs`, `datetime.rs`, `info.rs`).
   Signature: `fn name(args: &[Arg]) -> Result<Value, ErrorKind>`.
   Helpers on `Arg` (`number()`, `text()`, `boolean()`, `scalar()`, `as_range()`) do Excel's type coercion for you.
2. Register it in `FUNCTIONS` in `functions/mod.rs` with its syntax and a one-sentence summary
   (both appear in the formula autocomplete). Keep the list alphabetical; a test checks.
   If Excel added the function after 2007, also add it to `NEWER_FUNCTIONS` in `crates/formats/src/xlsx.rs`.
3. Add tests next to it using `test_util::run("YOURFUNC(A1:A5)")`. The fixture data is described in `functions/mod.rs`.
4. Add a row to `crates/sheets-engine/FUNCTIONS.md` (a test fails if you forget).

Match Excel's behaviour, including edge cases and error values. If you deviate, say so in `FUNCTIONS.md`.

### Change what JavaScript can call

Add a method in `crates/wasm/src/lib.rs` with `#[wasm_bindgen]`, run
`pnpm build:wasm`, and the TypeScript types in `packages/engine/pkg` update
automatically. Keep these methods thin: convert types and call the engine.

### Change the grid UI

Sheets UI files in `apps/web/src/sheets/`:

| File | What |
|---|---|
| `SheetsApp.tsx` | All UI state; turns actions into engine calls; ribbon, menus, autocomplete logic |
| `Grid.tsx` | Mouse and keyboard on the grid: selection, point mode, fill handle, column resize |
| `paint.ts` | Draws the visible cells, styles, highlights and headers on the canvas |
| `layout.ts` | Row/column positions with custom sizes and hidden lines (fast for 1M rows) |
| `sheetGeometry.ts` | Builds the geometry from the engine, incl. automatic row heights |
| `panes.ts`, `merges.ts` | Frozen panes; merged cells and Excel-style navigation |
| `formulaText.ts` | Reading formula text: references, word at caret, enclosing function |
| `geometry.ts`, `clipboard.ts` | Selection maths, tab-separated clipboard text |
| `FormulaBar`, `FormulaAssist`, `ContextMenu`, `FindBar`, `SheetTabs` | Smaller components |

Pure helpers have Vitest tests next to them (`*.test.ts`).

### Add a ribbon command

Commands are plain data (`RibbonCommand` in `packages/ui/src/Ribbon.tsx`):
`{ id, label, onRun, group, icon, size, shortcut, title, active, menu, render }`.
Put them in the Excel tab and group where Excel has them — groups in use:
Clipboard, Font, Alignment, Number, Styles, Cells, Editing, Sort & Filter,
Data Tools, Charts, Function Library, Window, Zoom, Print. Icons are inline SVGs
in `packages/ui/src/icons.tsx` (16×16 grid). App chrome styles live in
`apps/web/src/excel.css`; grid colours in `apps/web/src/sheets/paint.ts`.

## Before you open a pull request

```sh
pnpm lint        # cargo fmt --check + clippy with warnings as errors
pnpm test        # all Rust and TypeScript tests
pnpm typecheck
```

CI runs the same commands. In addition:

- **Tests:** new behaviour needs tests; bug fixes need a test that failed before the fix.
- **Docs:** update docs in the same pull request as the code. That means
  `FUNCTIONS.md` for functions, module docs for engine changes, `README.md` for
  user-visible features, `docs/ARCHITECTURE.md` for structural changes, and the
  `Unreleased` section of `CHANGELOG.md` for anything a user would notice.
- **Scope:** one topic per pull request. Small PRs get reviewed faster.
- **Commits:** write a short imperative summary line ("Add XLOOKUP"), and explain *why* in the body if it isn't obvious.

## Good first issues

Pick any ⬜ row from [docs/sheets/EXCEL_FEATURES.md](docs/sheets/EXCEL_FEATURES.md). Also:

- Spreadsheet functions from the "Not yet supported" list in `crates/sheets-engine/FUNCTIONS.md`
- Column resizing in the grid
- Cross-sheet references (`Sheet2!A1`)

## Licensing

By contributing you agree that your work is dual-licensed under MIT and Apache-2.0, as described in [README.md](README.md#license).
