# sheets-engine

The spreadsheet engine: cell storage, formulas, functions and recalculation.
Pure Rust with no I/O — it runs natively in tests and as WebAssembly in the browser.

## Tour

| File | Responsibility |
|---|---|
| `cell.rs` | `CellRef` (`B12`), `Range` (`A1:C3`), Excel limits |
| `value.rs` | `Value` (number/text/bool/error/empty), Excel coercion and comparison rules, number display |
| `formula/lexer.rs` | Formula text → tokens |
| `formula/parser.rs` | Tokens → `Expr` AST, with Excel operator precedence |
| `formula/eval.rs` | `Expr` → `Value`, reading cells through the `CellSource` trait |
| `formula/refs.rs` | Rewrites references in formula text: shift on copy/fill, remap on insert/delete |
| `functions/` | Built-in functions (with syntax and summary for autocomplete) and their registry |
| `style.rs` | `Style` (font flags, colours, alignment, `NumberFormat`) and `StylePatch` for toolbar edits |
| `format.rs` | Number formatting for display (₹ Indian grouping, %, dates) and parsing typed `18%`, `₹1,500`, dates |
| `date.rs` | Excel date serial numbers, including the 1900 leap-year quirk; pluggable clock for `TODAY()` |
| `fill.rs` | Fill-handle series detection |
| `sheet.rs` | One grid of cells + styles + column widths + dependency graph + recalculation; structural edits and sort |
| `workbook.rs` | Named sheets |
| `ops.rs` | Undoable edits (`SetCells`, `SetStyles`, `SetColWidths`, `ReplaceSheet`, sheet ops, `Batch`) |

## How an edit is processed

1. `Sheet::set_inputs` stores the raw input. Inputs starting with `=` are parsed
   into an `Expr`; anything else becomes a literal `Value` (`Value::from_input`).
2. Each formula records the ranges it reads (its *precedents*). The sheet keeps
   the reverse map, *dependents*: cell → formulas that read it.
3. Starting from the changed cells, the sheet follows `dependents` to collect
   every stale formula, then evaluates them in topological order (Kahn's
   algorithm). This is iterative, so long chains can't overflow the stack.
4. Formulas still waiting after the sort are part of a cycle and become `#CIRC!`.

Only affected cells are recalculated, so editing one cell in a big sheet is cheap.

## Errors are values

`#DIV/0!`, `#N/A` and friends are `Value::Error(ErrorKind)`, not Rust errors.
They flow through formulas the way they do in Excel, until something like
`IFERROR` handles them. Inside function implementations, returning
`Err(ErrorKind::X)` is a shortcut for producing `Value::Error(X)`.

## Structural edits

Inserting/deleting rows or columns and sorting build a *new* `Sheet`
(`Sheet::with_inserted`, `with_deleted`, `with_sorted_rows`) with every formula
rewritten by `formula/refs.rs`. The UI applies it with `ops::ReplaceSheet`, whose
inverse is simply the old sheet, so undo is exact.

## Known limitations

- No cross-sheet references (`Sheet2!A1`), whole-column references (`A:A`) or named ranges yet.
- Large ranges in formulas (e.g. `SUM(A1:A100000)`) are expanded cell by cell in the dependency graph.
- Volatile functions (`TODAY`, `NOW`, `RAND`) only recalculate when something they depend on changes.
- `ReplaceSheet` keeps a full copy of the sheet per structural edit in the undo history.

## Testing

```sh
cargo test -p sheets-engine
```

Function tests use a fixed fixture (A1:A5 = 1..5, B1:B5 = a..e, see
`functions/mod.rs`) so each test reads like a spreadsheet formula.
