# Excel feature checklist

What Office Sheets has compared with Microsoft Excel, and what's left.
Use it to pick work: each row says how important it is and which milestone it belongs to.

**Status:** ✅ done · 🟡 partly done · ⬜ not started
**Priority:**
- **P1** — needed for everyday use. All P1 items done ≈ 80% of what most people use Excel for.
- **P2** — commonly used; do after P1.
- **P3** — specialised or rarely used.

**Milestones** (M1–M7, see [the bottom of this page](#milestones-to-reach-80)) group the P1 items into shippable chunks.

When you finish an item, update its row in the same pull request.

---

## 1. Entering and editing data

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Type values, formulas; edit with F2 / double-click / formula bar | ✅ | P1 | – |
| Undo / redo for every action | ✅ | P1 | – |
| Copy, cut, paste (with Excel/Google Sheets; formulas shift inside Office) | ✅ | P1 | – |
| Fill handle: series, day/month names, formulas; Ctrl+D / Ctrl+R | ✅ | P1 | – |
| Typed formats detected (`18%`, `₹1,500`, `15-08-2025`) | ✅ | P1 | – |
| Hindi and other IME input | ✅ | P1 | – |
| New line inside a cell (Alt+Enter) | ✅ | P1 | M1 |
| Fill the whole selection with one entry (Ctrl+Enter) | ✅ | P1 | M1 |
| Paste Special: values only (also Ctrl+Shift+V), formatting only, formulas only, transpose | ✅ | P1 | M1 |
| Find & Replace (Ctrl+F / Ctrl+H, match case, whole cell, replace all) | ✅ | P1 | M1 |
| Move cells by dragging the selection border | ⬜ | P2 | – |
| AutoComplete text from the same column while typing | ⬜ | P2 | – |
| Notes / comments on cells | ⬜ | P2 | – |
| Hyperlinks in cells | ⬜ | P2 | – |
| Flash Fill (Ctrl+E) | ⬜ | P3 | – |
| Spell check | ⬜ | P3 | – |

## 2. Moving around and selecting

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Arrows, Shift+arrows, Page Up/Down, Home, Tab/Enter (Enter after Tabs returns to the first column, like Excel) | ✅ | P1 | – |
| Click/drag select; whole row/column/sheet select | ✅ | P1 | – |
| Ctrl+arrow (jump to edge of data), Ctrl+Shift+arrow (select to edge) | ✅ | P1 | M1 |
| Ctrl+Home / Ctrl+End | ✅ | P1 | M1 |
| Type a cell address in the Name Box to jump (Go To, Ctrl+G / F5) | ✅ | P1 | M1 |
| Shift+Space / Ctrl+Space (select row / column) | ✅ | P1 | M1 |
| Zoom (50%–200%, status bar or Ctrl + mouse wheel) | ✅ | P1 | M1 |
| Several separate selections (Ctrl+click) | ⬜ | P3 | – |

## 3. Formatting

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Bold, italic, underline, strikethrough | ✅ | P1 | – |
| Text colour, fill colour | ✅ | P1 | – |
| Horizontal alignment | ✅ | P1 | – |
| Number formats: number, ₹/$ currency, %, date, time, text; decimals ± | ✅ | P1 | – |
| Column width: drag, autofit | ✅ | P1 | – |
| Font family and font size (incl. Hindi fonts Mangal / Nirmala UI), increase/decrease size | ✅ | P1 | M2 |
| Borders (Excel's presets, line styles, line colours) | ✅ | P1 | M2 |
| Merge & Center, Merge Across, Merge Cells, Unmerge | ✅ | P1 | M2 |
| Wrap text (rows grow to fit, like Excel's automatic row height) | ✅ | P1 | M2 |
| Vertical alignment (top / middle / bottom) | ✅ | P1 | M2 |
| Row height: drag, double-click to autofit, Row Height dialog (points); automatic height for big fonts and wrapped text | ✅ | P1 | M2 |
| Hide / unhide rows and columns | ✅ | P1 | M2 |
| Format Painter (click once, double-click to keep painting, Esc to stop) | ✅ | P1 | M2 |
| Conditional formatting: highlight rules (>, <, between, equal, text contains, date occurring, duplicates, top/bottom, above/below average, formula), Manage Rules | ✅ | P1 | M4 |
| Conditional formatting: colour scales, data bars | ✅ | P2 | M4 |
| Conditional formatting: icon sets, custom format picker | ⬜ | P2 | – |
| Format as Table (banded rows, header style) | ⬜ | P2 | – |
| Custom number format codes (`0.0"kg"`, `[Red]`…) | ⬜ | P2 | – |
| Accounting, fraction, scientific formats | ⬜ | P2 | – |
| Indent, text rotation | ⬜ | P3 | – |
| Cell styles and document themes | ⬜ | P3 | – |

## 4. Formulas and calculation

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Formula engine, operators, error values, incremental recalculation, cycle detection | ✅ | P1 | – |
| 158 functions: math, statistics, text, logical, lookup (`XLOOKUP`…), date & time, financial, reference, `SUMIFS` family | ✅ | P1 | – |
| Autocomplete, syntax hints, click/arrow references with colour highlights | ✅ | P1 | – |
| References rewritten on insert/delete/sort/copy | ✅ | P1 | – |
| References to other sheets (`Sheet2!A1`, `'My Sheet'!A1:B5`), updated on rename/delete; click another sheet while typing a formula | ✅ | P1 | M3 |
| Whole column / row references (`A:A`, `1:1`), cheap to recalculate | ✅ | P1 | M3 |
| F4 to cycle `$A$1` → `A$1` → `$A1` → `A1` while editing | ✅ | P1 | M3 |
| `TODAY`, `NOW`, `RAND`, `OFFSET`, `INDIRECT` recalculate on every change (volatile) | ✅ | P1 | M3 |
| `SUBTOTAL`, `AGGREGATE`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `OFFSET`, `INDIRECT`, `ADDRESS` | ✅ | P1 | M3 |
| Financial functions: `PMT`, `FV`, `PV`, `NPV`, `XNPV`, `IRR`, `XIRR`, `RATE`, `NPER`, `IPMT`, `PPMT`, `SLN`, `DB`, `DDB`, `EFFECT`, `NOMINAL` | ✅ | P1 | M3 |
| Named ranges (Name Manager, use names in formulas) | ⬜ | P2 | – |
| Dynamic arrays that spill: `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE` | ⬜ | P2 | – |
| Show formulas (Ctrl+`) | ⬜ | P2 | – |
| Error checking, trace precedents/dependents | ⬜ | P3 | – |
| `LET`, `LAMBDA` | ⬜ | P3 | – |
| Iterative calculation, R1C1 style, links to other files | ⬜ | P3 | – |

## 5. Data tools

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Sort by one column, A→Z / Z→A | ✅ | P1 | – |
| Freeze panes (top row, first column, at selection) | ✅ | P1 | M4 |
| Filter (AutoFilter dropdowns: by value with search, text/number conditions, top 10, above/below average, sort from header, Ctrl+Shift+L, Clear/Reapply, "N of M records found") | ✅ | P1 | M4 |
| Data validation: dropdown lists (Alt+↓), number/date/time/text-length limits, custom formula, input message, Stop/Warning/Information alerts, Circle Invalid Data | ✅ | P1 | M4 |
| Remove duplicates | ✅ | P1 | M4 |
| Custom Sort dialog (several levels, headers) | ✅ | P1 | M4 |
| Sort by several columns, sort with header row option | ⬜ | P2 | – |
| Text to Columns | ⬜ | P2 | – |
| Pivot tables | ⬜ | P2 | – |
| Excel Tables (structured, auto-expanding ranges) | ⬜ | P2 | – |
| Group / outline rows and columns | ⬜ | P3 | – |
| Goal Seek, Scenario Manager, Solver | ⬜ | P3 | – |
| Power Query, data connections | ⬜ | P3 | – |

## 6. Charts and pictures

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Charts: column, bar, line, pie (insert from selection, move, resize, title, legend) | ✅ | P1 | M5 |
| Charts: stacked / 100% variants, area, doughnut, scatter; data labels, axis titles, colour sets, switch row/column, select data, change type | ✅ | P2 | M5 |
| Charts: combo, secondary axis, trendlines, per-point formatting, in-place title editing | ⬜ | P2 | – |
| Insert images (PNG, JPEG, GIF, BMP), move, resize, saved in .xlsx | ✅ | P2 | M5 |
| Sparklines | ⬜ | P3 | – |
| Shapes, text boxes | ⬜ | P3 | – |

## 7. Sheets and workbook

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Several sheets; add, rename, delete | ✅ | P1 | – |
| Autosave in the browser | ✅ | P1 | – |
| Reorder sheets by dragging tabs | ⬜ | P2 | – |
| Duplicate sheet, tab colours, hide sheet | ⬜ | P2 | – |
| Protect sheet / workbook with password | ⬜ | P3 | – |

## 8. Files, saving and printing

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Open/save `.xlsx`: values, formulas, formatting, number formats, column widths | ✅ | P1 | – |
| Open/save `.csv` | ✅ | P1 | – |
| Save back to the same file on disk (Ctrl+S, no re-download), recent files | ⬜ | P1 | M6 |
| Print and export to PDF (print area, fit to page width, portrait/landscape) | ⬜ | P1 | M6 |
| `.xlsx` round-trip for every M1–M5 feature (borders, merges, fonts, row heights, freeze panes, filters, validation, conditional formatting, charts) | ⬜ | P1 | M6 |
| Page setup: margins, headers/footers, repeat header rows | ⬜ | P2 | – |
| CSV options: delimiter, encoding | ⬜ | P2 | – |
| `.ods` (LibreOffice) | ⬜ | P2 | – |
| Old `.xls` (Excel 97–2003) | ⬜ | P3 | – |

## 9. Quality, performance and reach

| Feature | Status | Priority | Milestone |
|---|---|---|---|
| Smooth with 100,000+ rows (3 × 100k rows opens in ~2.6 s, 2.8 ms paints, no autosave freezes) | ✅ | P1 | M7 |
| Browser tests for the main flows, run in CI | ⬜ | P1 | M7 |
| Keyboard shortcuts help (Ctrl+/) | ⬜ | P2 | – |
| Dark mode | ⬜ | P2 | – |
| Hindi UI (and other languages) | ⬜ | P2 | – |
| Desktop app (Tauri) | ⬜ | P2 | – |
| Touch / mobile | ⬜ | P2 | – |
| Screen reader support | ⬜ | P2 | – |

## 10. Later (beyond Excel desktop parity)

| Feature | Status | Priority |
|---|---|---|
| Real-time co-editing, sharing links, version history | ⬜ | P3 |
| Scripting / macros (JavaScript, not VBA) | ⬜ | P3 |
| AI assistant (write formulas, explain data) | ⬜ | P3 |

---

## Milestones to reach 80%

Done in this order; each one is usable on its own.

| # | Milestone | Contents |
|---|---|---|
| M1 ✅ | Editing & navigation | Alt+Enter, Ctrl+Enter, Paste Special, Replace, Ctrl+arrows, Ctrl+Home/End, Go To / Name Box, row/column select shortcuts, zoom |
| M2 ✅ | Formatting | Font family/size, borders, merge, wrap, vertical alignment, row height, hide/unhide, Format Painter |
| M3 ✅ | Formulas | Cross-sheet references, `A:A` references, F4, volatile functions, `SUBTOTAL`/`AGGREGATE`/`ROW`/`COLUMN`, financial functions |
| M4 | Data tools | Freeze panes, filter, data validation dropdowns, remove duplicates, conditional formatting (highlight rules) |
| M5 | Charts | Column, bar, line, pie |
| M6 | Files & print | Save to disk (Ctrl+S), print / PDF, `.xlsx` round-trip for everything above |
| M7 | Quality | 100k-row performance, browser tests in CI |
