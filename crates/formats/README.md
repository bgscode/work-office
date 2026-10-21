# office-formats

File import/export for the suite.

| Module | Format | Status |
|---|---|---|
| `csv` | CSV (RFC 4180) | read + write (exports displayed, formatted values) |
| `xlsx` | Excel workbook | read + write: values, formulas, formatting, number formats, column widths, multiple sheets |
| `docx`, `pptx`, `pdf` | | planned |

`xlsx.rs` starts with a map of the parts inside an `.xlsx` ZIP. All XML parsing
goes through one small helper (`walk`), so the rest of the code doesn't depend
on the `quick-xml` API.

Debug a file from the command line:

```sh
cargo run -p office-formats --example xlsx_roundtrip -- input.xlsx [resaved.xlsx]
```

Not yet supported in `.xlsx`: borders, font names/sizes, merged cells, row
heights, charts. Theme colours are ignored (only explicit RGB colours are read).
Excel *shared formulas* are imported as values except for the first cell of each
group. Functions newer than Excel 2007 are written with Excel's `_xlfn.` prefix
(`NEWER_FUNCTIONS` in `xlsx.rs`) — add new ones there when you add such functions.
