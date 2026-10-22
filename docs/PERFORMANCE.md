# Performance

Office Sheets should feel instant on workbooks with hundreds of thousands of
rows. This page records how we measure that and the current numbers.

## Benchmark workbook

3 sheets × 100,000 rows × 10 columns = 3,000,000 cells, including 300,000
formulas (`=C2*D2`), text, numbers, dates and booleans. 21.8 MB as `.xlsx`.

Generate it and time the engine natively:

```sh
OFFICE_PROFILE=1 cargo run --release -p office-formats --example bench_open        # 3 × 100k rows
OFFICE_PROFILE=1 cargo run --release -p office-formats --example bench_open -- 20000 1
cargo run --release -p office-formats --example bench_walk                         # XML reader only
```

`bench_open` writes the file to `/tmp/bench_open.xlsx`; open that in the app to
measure the browser.

## Results (Apple Silicon laptop, Chrome)

| | Before | Now |
|---|---|---|
| Open the workbook in the browser | did not finish (minutes, debug WebAssembly build) | **2.6 s** to an interactive grid |
| Open natively (engine only) | 6.7 s | 2.1 s |
| Save `.xlsx` natively | 4.0 s | 2.0 s |
| Autosave after an edit | 2.7 s freeze per save, and the save failed (5 MB `localStorage` limit) | **4 ms** (journal), no freeze |
| Reopen with unsaved edits | — | 3.2 s (file + journal replay) |
| Paint one frame while scrolling | — | 2.8 ms (well under 16 ms for 60 fps) |
| Switch sheet | — | 57 ms |

## What made the difference

- **Release WebAssembly in development.** `pnpm dev` builds the engine with
  optimisations; the debug build is 10–50× slower. Use `pnpm dev:debug` only
  when you need Rust debug info.
- **A zero-allocation XML reader** (`crates/formats/src/xml.rs`) instead of
  allocating names and attribute lists for every element, and **one pass** over
  each worksheet for cells and layout (it used to be read twice).
- **Typed loading** (`Sheet::load_cells`): values from a file arrive as numbers,
  text or booleans, so they skip the typed-input parsing used for keyboard input.
- **Inputs stored only when needed**: most literal cells keep just their value;
  `Sheet::input` derives the text when asked.
- **FxHash** instead of SipHash for cell maps.
- **A faster writer**: shared strings like Excel, no per-cell temporary strings,
  hashed style lookup, deflate level 3.
- **Autosave journal** (`apps/web/src/sheets/journal.ts`): each edit is appended
  to IndexedDB in milliseconds; the whole `.xlsx` is rewritten after every edit
  only when that is cheap (under 150 ms), otherwise after a pause once 200 edits
  have piled up. Opening replays the journal over the last saved file.
- **IndexedDB** for document bytes (no size limit, no base64).
- Opening a file stores the original bytes rather than writing them out again,
  and React's development double-mount shares one load instead of parsing twice.

## Pitfalls

- `std::time::Instant` panics in `wasm32-unknown-unknown`; the error shows as
  "unreachable" in the browser. Timing code in library crates must be
  `#[cfg(not(target_arch = "wasm32"))]` (see `Timer` in `xlsx.rs`).
- Headless Chrome runs `requestAnimationFrame` at about 30 fps, so frame times
  measured there bottom out near 33 ms; measure paint time itself instead.

## Next steps

- Run the engine in a Web Worker so even a multi-second open never blocks the page.
- Parse sheets in parallel (worker threads) for workbooks with many large sheets.
- Keep `rows_to_measure` incrementally instead of scanning every cell after each edit.
