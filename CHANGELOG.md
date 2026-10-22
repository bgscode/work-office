# Changelog

All notable changes are recorded here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## Unreleased

### Fixed

- Clicking or dragging a row border in the row headers crashed the app; dragging a column border lost the other columns' custom widths in the live preview.
- Opening a file from Home could open an empty workbook instead (a race between navigation and the pending file).
- Ctrl+S remembers the file on disk across reloads (the browser asks once to allow saving again).
- Big printouts (hundreds or thousands of pages) are prepared in the background with progress and Cancel instead of freezing the app.
- Freeze Panes works like Excel when the sheet is scrolled: it freezes the rows and columns on screen instead of the first ones, so the grid no longer gets stuck when the frozen part is bigger than the window; scrolled frozen panes in Excel files (`<sheetView topLeftCell>`) open correctly.

### Added

- Notes: Review → New Note (Shift+F2) opens Excel's yellow note box beside the cell; cells with notes show a red triangle and their note on hover. Review → Delete, Previous Note, Next Note and Show All Notes; Home → Clear → Clear Notes; right-click → New/Edit/Delete Note. Notes move with their cells (insert, delete, sort) and are read from and saved to .xlsx (shown by Excel).
- Hyperlinks: Insert → Link (Ctrl+K) to a web page, an e-mail address or a place in the workbook, with text to display and ScreenTip; the cell gets Excel's blue underlined link style. A click follows the link (click and hold selects the cell); right-click → Edit/Open/Remove Hyperlink; Home → Clear → Clear/Remove Hyperlinks. Saved in .xlsx. Script links (`javascript:`) never run.

- Sheet tabs like Excel: drag a tab to move the sheet (Ctrl+drag copies it), Move or Copy… (with Create a copy), Tab Color…, Hide and Unhide… from the tab's right-click menu or Home → Format; Ctrl+PgUp/PgDn switch sheets. Copies are named `Sales (2)` and get their own copies of sheet-scoped names. Tab colours and hidden sheets are saved in .xlsx; hidden sheets don't print with the whole workbook.
- Adding a sheet no longer copies the whole workbook in memory first (faster on big files).

- AutoSum (Home → Editing, Alt+=): Sum, Average, Count Numbers, Max or Min of the numbers above the cell (else to its left), ready to adjust before Enter; with a range selected, a total goes under each column.

- Defined names: type a name in the Name Box to name the selected cells, or use Formulas → Define Name / Name Manager (Ctrl+F3) with workbook or sheet scope and comments. Formulas use names (`=SUM(Sales)*(1+GST)`), autocomplete suggests them, the Name Box goes to them, and they follow renamed or deleted sheets and inserted or deleted rows. Saved in .xlsx. Formulas → Show Formulas (Ctrl+`).

- Dynamic arrays (Excel 365): formulas that return several values spill into the cells around them, with `#SPILL!` when something is in the way (it spills again once cleared); spilled cells show their formula greyed in the formula bar and the area is outlined. Operators and functions like `IF`, `UPPER`, `ROUND` work element by element (`=A2:A10*B2:B10`, `=SUM(A2:A10*B2:B10)`). New functions: `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`, `TRANSPOSE`, `TAKE`, `DROP`, `CHOOSEROWS`, `CHOOSECOLS`, `VSTACK`, `HSTACK`, `TOCOL`, `TOROW`, `EXPAND`, `TEXTSPLIT`, `XMATCH`, `LET`; omitted arguments (`SORT(A1:A9,,-1)`). Saved to .xlsx as array formulas with `_xlfn` names; Excel's dynamic arrays open correctly.

- Number formats: every Excel format code now displays exactly as in Excel — sections (positive;negative;zero;text), colours (`[Red]`), conditions (`[>=1000]`), thousands scaling (`#,##0,"K"`), literals, dates and times (`dd-mmm-yyyy`, `dddd`, `h:mm AM/PM`, `[h]:mm:ss`, `.000`), fractions, scientific and accounting formats. Codes in .xlsx files are kept as they are instead of being simplified. Format Cells dialog (Ctrl+1, right-click, Number box → More Number Formats…) with Number (General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Custom with a live sample), Alignment, Font, Border and Fill tabs; one undo step.

- Charts: double-click a chart title to edit it in place; Change Colors now recolours pie and doughnut slices (slice colours are saved in .xlsx).
- Ribbon: on narrow windows groups fold into one button each, from the right, like Excel; the button opens the group's commands.

- Quality (M7): browser tests for the main flows (typing and formulas, formatting and undo, filter and validation, charts and shapes, freeze panes, print, reload, damaged files, crash recovery) run in CI with Playwright; damaged, truncated or unsupported files (.xls, password-protected) give a clear message instead of failing, covered by fuzz-style tests; a recovery screen replaces a blank page if the app or the WebAssembly engine crashes.

- Shapes: Insert → Shapes with Excel's gallery (lines and arrows, rectangles, basic shapes, block arrows, stars, callouts, flowchart) and Insert → Text Box. Drag on the sheet to draw (Shift keeps it square, a click inserts the default size); type to add text. Shape Format tab: Edit Shape, fill, outline colour and weight, arrowheads, font size, bold/italic, alignment, text colour, arrange. Shapes move, resize, undo, print, and are read from and saved to .xlsx (including Excel's theme-styled shapes).

- Files & print (M6): File → Print (Ctrl+P) with a page preview and Excel's settings (active sheet, whole workbook or selection; orientation; paper size; margins; scaling) and Export PDF through the browser's "Save as PDF"; a Page Layout tab (Margins, Orientation, Size, Print Area, Print Titles, Scale to Fit, print gridlines/headings) and the Page Setup dialog (Page, Margins, Header/Footer, Sheet). Printouts keep formatting, conditional formatting, merged cells, charts and pictures, and repeat title rows/columns. Save (Ctrl+S) writes the .xlsx to a file on disk and later saves overwrite it; Save As (F12). Page setup is saved in .xlsx. Opening .csv/.tsv/.txt detects the delimiter. A round-trip test covers every feature .xlsx can hold.

- Charts (M5): Insert → Charts (clustered/stacked/100% column and bar, line, line with markers, area, stacked area, pie, doughnut, scatter) from the selection or the table around the active cell; charts read their cells live and follow row/column changes; move, resize (Shift keeps the shape), nudge with arrows, delete, bring to front/send to back, all undoable. Chart Design tab: chart and axis titles, legend position, data labels, colour sets, switch row/column, select data, change chart type. Insert → Pictures. Big series draw one min/max per pixel, so 100,000-point charts stay fast. Charts and pictures are read from and saved to .xlsx.

- One Excel-style dropdown (`Select`) everywhere instead of the browser's: ribbon, dialogs, status bar. Font, size and other dropdown lists open above the ribbon instead of being clipped. The duplicate "Paste special" box is gone (it is in the Paste menu, as in Excel), and the active ribbon tab has a single underline.

- Data tools (M4): AutoFilter (Ctrl+Shift+L) with Excel's header dropdowns: value checklist with search, Text/Number Filters (custom conditions with wildcards), Top 10, above/below average, sort from the header, Clear and Reapply, blue row numbers and "N of M records found"; Custom Sort with several levels; Remove Duplicates; Data Validation (lists with an in-cell dropdown and Alt+↓, number/date/time/text-length limits, custom formulas, input messages, Stop/Warning/Information alerts, Circle Invalid Data); Conditional Formatting (highlight rules, top/bottom, above/below average, duplicates, dates, formulas, data bars, colour scales, Manage Rules). All saved in .xlsx. Enter after a run of Tab entries returns to the starting column, like Excel.

- Formulas (M3): references to other sheets with workbook-wide recalculation, kept correct when sheets are renamed or deleted and when rows/columns move; pointing at cells on another sheet while typing a formula; whole column/row references (`A:A`, `1:1`) without per-cell cost; F4; volatile functions; `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `OFFSET`, `INDIRECT`, `ADDRESS`, `SUBTOTAL`, `AGGREGATE`, financial functions (`PMT`, `NPV`, `IRR`, `XIRR`, `RATE`…), percentiles and quartiles (158 functions). Date, time and loan functions get Excel's default number formats.

- Performance: a 3-sheet × 100,000-row workbook (3M cells) opens in about 2.6 s in the browser instead of hanging; autosave no longer freezes (edits are journaled to IndexedDB in milliseconds); faster `.xlsx` reading and writing (single-pass zero-allocation XML reader, shared strings, typed loading). See docs/PERFORMANCE.md.

- Grid layout: row heights (drag the row border, double-click to autofit, Row Height / Column Width dialogs in Excel's units), automatic row height for wrapped text and bigger fonts, hide/unhide rows and columns, Merge & Center / Across / Cells / Unmerge (selection and arrow keys treat merged cells as one), Freeze Panes (at selection, top row, first column). All saved in .xlsx.

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
