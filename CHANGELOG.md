# Changelog

All notable changes are recorded here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## Unreleased

### Added

- Cell formatting: font family and size boxes, increase/decrease font size, Borders menu with Excel's presets, line styles and colours, Wrap Text, top/middle/bottom alignment, Format Painter; Alt+Enter turns on Wrap Text. All saved in .xlsx (fonts, borders, alignment).

- Home page in the style of the Microsoft 365 start page: app rail, Start new tiles (blank workbook; document and presentation marked coming soon), open or drag-and-drop .xlsx/.csv, and Recent files with search and app filters. Every workbook is autosaved separately in the browser (up to 20) and listed in Recent. Apps open at `#/sheets/<id>`, so Back and refresh work. The title bar now has a Home button instead of the app switcher.

- Excel look and feel: green title bar with file name, save status and Quick Access (save, undo, redo); ribbon with Excel's tabs and labelled groups, icons, split buttons, tooltips and collapse (Ctrl+F1); formula bar with Name Box dropdown, ✕/✓ and fx (Insert Function dialog); Excel grid defaults (64px columns, 20px rows, Calibri 11, green selection); Excel-style sheet tabs, status bar (mode, statistics, zoom slider), menus and dialogs.

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
