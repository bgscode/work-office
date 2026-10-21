# Changelog

All notable changes are recorded here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## Unreleased

### Added

- Sheets: canvas grid with virtual scrolling, formula bar, multiple sheets, status bar (sum/count/average).
- Sheets: formula engine with incremental recalculation, cycle detection and 46 functions.
- Sheets: open/save `.xlsx` and `.csv`; copy/paste as tab-separated text; undo/redo.
- Sheets: IME-friendly typing (Hindi, Chinese, Japanese…).
- Sheets: formula autocomplete with syntax hints; point-and-click and arrow-key references with coloured highlights; open parentheses are closed on Enter.
- Sheets: fill handle with series detection (numbers, numbered text, day and month names, formulas); Ctrl+D / Ctrl+R; double-click to fill to the end of the data.
- Sheets: cell formatting (bold, italic, underline, strikethrough, colours, alignment) and number formats (₹ with Indian grouping, $, %, date, time, text); typed `18%`, `₹1,500`, `15-08-2025` set the format automatically.
- Sheets: insert/delete rows and columns with formula rewriting, sort, find, right-click menu, column resize/autofit, sheet rename/delete.
- Sheets: 81 new functions (127 total): conditional aggregates (`SUMIFS`…), `XLOOKUP`, statistics, text (`TEXT`, `SUBSTITUTE`…), date & time.
- Sheets: copy/paste within Office keeps formulas (shifted) and formatting.
- Sheets: autosave to browser storage.
- Files: `.xlsx` import/export now includes formatting and column widths.
- Sheets (M1): Alt+Enter line breaks, Ctrl+Enter fills the selection, Paste Special (values/formatting/formulas/transpose, Ctrl+Shift+V), Find & Replace with match-case and whole-cell options, Ctrl+arrow / Ctrl+Shift+arrow, Ctrl+Home/End, Go To via the Name Box (Ctrl+G / F5), Shift+Space / Ctrl+Space, zoom 50–200%.
