//! JavaScript bindings for the engines, built with `wasm-bindgen`.
//!
//! Build: `pnpm build:wasm` (runs wasm-pack, output goes to `packages/engine/pkg`).
//! The UI imports these through `@office/engine`.
//!
//! Keep this layer thin: convert types and forward calls. Logic belongs in the
//! engine crates where it can be unit-tested without a browser.
//!
//! Conventions: rectangles are passed as `top, left, bottom, right` (inclusive,
//! zero-based); every method that changes the document is one undo step.

// Bindings take flat arguments (rectangles as four numbers) to keep the JS side simple.
#![allow(clippy::too_many_arguments)]

use office_core::history::History;
use office_core::Op;
use sheets_engine::fill::FillDirection;
use sheets_engine::formula::refs::{self, Axis};
use sheets_engine::ops::{
    Batch, InsertSheet, RemoveSheet, RenameSheet, ReplaceSheet, SetCells, SetColWidths, SetStyles,
};
use sheets_engine::{Align, CellRef, FindOptions, NumberFormat, Range, Sheet, Style, StylePatch, Value, Workbook};
use wasm_bindgen::prelude::*;

mod files;

fn local_now_ms() -> f64 {
    let d = js_sys::Date::new_0();
    d.get_time() - d.get_timezone_offset() * 60_000.0
}

/// A snapshot taken by `copy`, so paste can bring formulas and formatting along.
struct Clip {
    origin: CellRef,
    rows: u32,
    cols: u32,
    inputs: Vec<String>,
    /// Computed values as inputs, for "paste values".
    values: Vec<String>,
    styles: Vec<Style>,
}

/// An open spreadsheet document plus its undo history.
#[wasm_bindgen]
pub struct SheetsDocument {
    book: Workbook,
    history: History<Workbook>,
    clip: Option<Clip>,
}

impl Default for SheetsDocument {
    fn default() -> Self {
        Self::new()
    }
}

impl SheetsDocument {
    fn from_book(book: Workbook) -> SheetsDocument {
        sheets_engine::date::set_clock(local_now_ms);
        SheetsDocument { book, history: History::default(), clip: None }
    }

    fn sheet(&self, sheet: usize) -> &Sheet {
        &self.book.sheets[sheet].sheet
    }

    fn apply(&mut self, op: impl Op<Workbook> + 'static) {
        self.history.apply(&mut self.book, Box::new(op));
    }
}

/// Cells of a block, row-major.
fn block(row: u32, col: u32, rows: u32, cols: u32) -> impl Iterator<Item = CellRef> {
    (row..row + rows).flat_map(move |r| (col..col + cols).map(move |c| CellRef::new(r, c)))
}

fn rect(top: u32, left: u32, bottom: u32, right: u32) -> Range {
    Range::new(CellRef::new(top, left), CellRef::new(bottom, right))
}

#[wasm_bindgen]
impl SheetsDocument {
    #[wasm_bindgen(constructor)]
    pub fn new() -> SheetsDocument {
        Self::from_book(Workbook::new())
    }

    #[wasm_bindgen(js_name = fromXlsx)]
    pub fn from_xlsx(bytes: &[u8]) -> Result<SheetsDocument, JsError> {
        let book = office_formats::xlsx::read(bytes).map_err(|e| JsError::new(&e.to_string()))?;
        Ok(Self::from_book(book))
    }

    #[wasm_bindgen(js_name = fromCsv)]
    pub fn from_csv(text: &str) -> SheetsDocument {
        let mut book = Workbook::empty();
        book.push_sheet("Sheet1".into(), office_formats::csv::read(text));
        Self::from_book(book)
    }

    #[wasm_bindgen(js_name = toXlsx)]
    pub fn to_xlsx(&self) -> Result<Vec<u8>, JsError> {
        office_formats::xlsx::write(&self.book).map_err(|e| JsError::new(&e.to_string()))
    }

    #[wasm_bindgen(js_name = toCsv)]
    pub fn to_csv(&self, sheet: usize) -> String {
        office_formats::csv::write(self.sheet(sheet))
    }

    // ---------------------------------------------------------------- sheets

    #[wasm_bindgen(js_name = sheetNames)]
    pub fn sheet_names(&self) -> Vec<String> {
        self.book.sheets.iter().map(|s| s.name.clone()).collect()
    }

    /// Append a new sheet and return its index.
    #[wasm_bindgen(js_name = addSheet)]
    pub fn add_sheet(&mut self) -> usize {
        let mut probe = self.book.clone();
        let index = probe.add_sheet();
        let sheet = probe.sheets.swap_remove(index);
        self.apply(InsertSheet { index, sheet });
        index
    }

    /// Returns false if the name is empty or already used.
    #[wasm_bindgen(js_name = renameSheet)]
    pub fn rename_sheet(&mut self, index: usize, name: &str) -> bool {
        let name = name.trim();
        let taken = self.book.sheets.iter().enumerate().any(|(i, s)| i != index && s.name.eq_ignore_ascii_case(name));
        if name.is_empty() || taken {
            return false;
        }
        self.apply(RenameSheet { index, name: name.to_string() });
        true
    }

    /// Returns false when it is the only sheet.
    #[wasm_bindgen(js_name = deleteSheet)]
    pub fn delete_sheet(&mut self, index: usize) -> bool {
        if self.book.sheets.len() <= 1 {
            return false;
        }
        self.apply(RemoveSheet { index });
        true
    }

    // ---------------------------------------------------------------- cells

    /// Raw input of a cell, e.g. "=SUM(A1:A3)".
    pub fn input(&self, sheet: usize, row: u32, col: u32) -> String {
        self.sheet(sheet).input(CellRef::new(row, col)).to_string()
    }

    /// Set many cells as one undoable step. The three arrays are parallel.
    #[wasm_bindgen(js_name = setInputs)]
    pub fn set_inputs(&mut self, sheet: usize, rows: Vec<u32>, cols: Vec<u32>, inputs: Vec<String>) {
        let edits = rows.into_iter().zip(cols).zip(inputs).map(|((r, c), s)| (CellRef::new(r, c), s)).collect();
        self.apply(SetCells { sheet, edits });
    }

    pub fn undo(&mut self) -> bool {
        self.history.undo(&mut self.book)
    }

    pub fn redo(&mut self) -> bool {
        self.history.redo(&mut self.book)
    }

    /// Display text (number formats applied) for a block of cells, row-major.
    #[wasm_bindgen(js_name = displayBlock)]
    pub fn display_block(&self, sheet: usize, row: u32, col: u32, rows: u32, cols: u32) -> Vec<String> {
        let s = self.sheet(sheet);
        block(row, col, rows, cols).map(|at| s.display(at)).collect()
    }

    /// Value kind for a block of cells (see `ValueKind` in `@office/engine`), row-major.
    #[wasm_bindgen(js_name = kindBlock)]
    pub fn kind_block(&self, sheet: usize, row: u32, col: u32, rows: u32, cols: u32) -> Vec<u8> {
        let s = self.sheet(sheet);
        block(row, col, rows, cols)
            .map(|at| match s.value(at) {
                Value::Empty => 0,
                Value::Number(_) => 1,
                Value::Text(_) => 2,
                Value::Bool(_) => 3,
                Value::Error(_) => 4,
            })
            .collect()
    }

    /// Style of each cell as JSON (see `CellStyle` in `@office/engine`), "" when unformatted.
    #[wasm_bindgen(js_name = styleBlock)]
    pub fn style_block(&self, sheet: usize, row: u32, col: u32, rows: u32, cols: u32) -> Vec<String> {
        let s = self.sheet(sheet);
        block(row, col, rows, cols).map(|at| style_json(s.style(at))).collect()
    }

    /// [count of non-empty cells, count of numbers, sum of numbers] for the status bar.
    #[wasm_bindgen(js_name = rangeStats)]
    pub fn range_stats(&self, sheet: usize, row: u32, col: u32, rows: u32, cols: u32) -> Vec<f64> {
        let s = self.sheet(sheet);
        let (mut count, mut numbers, mut sum) = (0.0, 0.0, 0.0);
        for at in block(row, col, rows, cols) {
            let v = s.value(at);
            if !v.is_empty() {
                count += 1.0;
            }
            if let Value::Number(n) = v {
                numbers += 1.0;
                sum += n;
            }
        }
        vec![count, numbers, sum]
    }

    /// [rows, cols] that contain data.
    pub fn extent(&self, sheet: usize) -> Vec<u32> {
        let (r, c) = self.sheet(sheet).extent();
        vec![r, c]
    }

    /// Next cell matching `query`: [row, col], or empty if none.
    pub fn find(
        &self,
        sheet: usize,
        query: &str,
        row: u32,
        col: u32,
        forward: bool,
        match_case: bool,
        whole_cell: bool,
    ) -> Vec<u32> {
        let opts = FindOptions { match_case, whole_cell };
        self.sheet(sheet)
            .find(query, CellRef::new(row, col), forward, opts)
            .map_or(Vec::new(), |at| vec![at.row, at.col])
    }

    /// Replace in one cell. Returns false if the cell doesn't match.
    #[wasm_bindgen(js_name = replaceIn)]
    pub fn replace_in(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        query: &str,
        replacement: &str,
        match_case: bool,
        whole_cell: bool,
    ) -> bool {
        let at = CellRef::new(row, col);
        let opts = FindOptions { match_case, whole_cell };
        let Some(input) = self.sheet(sheet).replace_in(at, query, replacement, opts) else { return false };
        self.apply(SetCells { sheet, edits: vec![(at, input)] });
        true
    }

    /// Replace in every matching cell as one undo step; returns how many cells changed.
    #[wasm_bindgen(js_name = replaceAll)]
    pub fn replace_all(
        &mut self,
        sheet: usize,
        query: &str,
        replacement: &str,
        match_case: bool,
        whole_cell: bool,
    ) -> u32 {
        let edits = self.sheet(sheet).replace_all(query, replacement, FindOptions { match_case, whole_cell });
        let count = edits.len() as u32;
        if count > 0 {
            self.apply(SetCells { sheet, edits });
        }
        count
    }

    /// Ctrl+arrow target from (row, col) in direction (d_row, d_col): [row, col].
    pub fn jump(&self, sheet: usize, row: u32, col: u32, d_row: i32, d_col: i32) -> Vec<u32> {
        let at = self.sheet(sheet).jump(CellRef::new(row, col), d_row, d_col);
        vec![at.row, at.col]
    }

    // ---------------------------------------------------------------- formatting

    /// Apply one formatting change to a rectangle. `key`/`value` pairs:
    /// bold|italic|underline|strike: "1"/"0"; color|fill: "#RRGGBB" or "" to reset;
    /// align: "left"|"center"|"right"|""; format: "general"|"number"|"currency"|
    /// "percent"|"date"|"time"|"datetime"|"text"; decimals: "+1"/"-1"; clear: "".
    #[wasm_bindgen(js_name = formatRange)]
    pub fn format_range(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        key: &str,
        value: &str,
    ) -> Result<(), JsError> {
        let patch = parse_patch(key, value).ok_or_else(|| JsError::new(&format!("unknown format {key}={value}")))?;
        let s = self.sheet(sheet);
        let styles = rect(top, left, bottom, right).cells().map(|at| (at, s.style(at).patched(&patch))).collect();
        self.apply(SetStyles { sheet, styles });
        Ok(())
    }

    /// Column widths that differ from the default, flattened as [col, px, col, px, ...].
    #[wasm_bindgen(js_name = columnWidths)]
    pub fn column_widths(&self, sheet: usize) -> Vec<f64> {
        self.sheet(sheet).col_widths().flat_map(|(c, w)| [c as f64, w]).collect()
    }

    #[wasm_bindgen(js_name = setColumnWidth)]
    pub fn set_column_width(&mut self, sheet: usize, first: u32, last: u32, px: f64) {
        self.apply(SetColWidths { sheet, widths: (first..=last).map(|c| (c, Some(px))).collect() });
    }

    // ---------------------------------------------------------------- structure

    /// `axis`: "rows" or "cols".
    #[wasm_bindgen(js_name = insertLines)]
    pub fn insert_lines(&mut self, sheet: usize, axis: &str, at: u32, count: u32) {
        let new = self.sheet(sheet).with_inserted(parse_axis(axis), at, count);
        self.apply(ReplaceSheet { sheet, with: Box::new(new) });
    }

    #[wasm_bindgen(js_name = deleteLines)]
    pub fn delete_lines(&mut self, sheet: usize, axis: &str, at: u32, count: u32) {
        let new = self.sheet(sheet).with_deleted(parse_axis(axis), at, count);
        self.apply(ReplaceSheet { sheet, with: Box::new(new) });
    }

    /// Sort the rows of a rectangle by the values in column `key_col`.
    #[wasm_bindgen(js_name = sortRange)]
    pub fn sort_range(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        key_col: u32,
        ascending: bool,
    ) {
        let new = self.sheet(sheet).with_sorted_rows(rect(top, left, bottom, right), key_col, ascending);
        self.apply(ReplaceSheet { sheet, with: Box::new(new) });
    }

    /// Fill-handle drag: extend the source rectangle `count` cells in
    /// `direction` ("down"|"up"|"right"|"left").
    pub fn fill(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        direction: &str,
        count: u32,
        increment_single: bool,
    ) {
        let dir = match direction {
            "up" => FillDirection::Up,
            "left" => FillDirection::Left,
            "right" => FillDirection::Right,
            _ => FillDirection::Down,
        };
        let (edits, styles) = self.sheet(sheet).fill(rect(top, left, bottom, right), dir, count, increment_single);
        self.apply(Batch(vec![Box::new(SetCells { sheet, edits }), Box::new(SetStyles { sheet, styles })]));
    }

    // ---------------------------------------------------------------- clipboard

    /// Remember a rectangle's inputs and formatting for `pasteCopied`, and
    /// return its displayed values as tab-separated rows for the system clipboard.
    pub fn copy(&mut self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32) -> Vec<String> {
        let r = rect(top, left, bottom, right);
        let s = self.sheet(sheet);
        let clip = Clip {
            origin: r.start,
            rows: r.rows(),
            cols: r.cols(),
            inputs: r.cells().map(|at| s.input(at).to_string()).collect(),
            values: r.cells().map(|at| s.value(at).to_input()).collect(),
            styles: r.cells().map(|at| s.style(at).clone()).collect(),
        };
        let display = r.cells().map(|at| s.display(at)).collect();
        self.clip = Some(clip);
        display
    }

    /// Paste the last `copy` at (row, col). `mode`: "all" (formulas shifted +
    /// formatting), "values" (computed values only), "formulas" (no formatting),
    /// "formats" (formatting only). `transpose` swaps rows and columns.
    /// Returns [rows, cols] pasted, or empty if nothing was copied.
    #[wasm_bindgen(js_name = pasteCopied)]
    pub fn paste_copied(&mut self, sheet: usize, row: u32, col: u32, mode: &str, transpose: bool) -> Vec<u32> {
        let Some(clip) = &self.clip else { return Vec::new() };
        let (out_rows, out_cols) = if transpose { (clip.cols, clip.rows) } else { (clip.rows, clip.cols) };
        let mut edits = Vec::new();
        let mut styles = Vec::new();
        for i in 0..out_rows {
            for j in 0..out_cols {
                let at = CellRef::new(row + i, col + j);
                if at.row >= sheets_engine::MAX_ROWS || at.col >= sheets_engine::MAX_COLS {
                    continue;
                }
                let (src_i, src_j) = if transpose { (j, i) } else { (i, j) };
                let k = (src_i * clip.cols + src_j) as usize;
                let source = CellRef::new(clip.origin.row + src_i, clip.origin.col + src_j);
                let (dr, dc) = (at.row as i64 - source.row as i64, at.col as i64 - source.col as i64);
                match mode {
                    "values" => edits.push((at, clip.values[k].clone())),
                    "formats" => {}
                    _ => edits.push((at, refs::shift_input(&clip.inputs[k], dr, dc))),
                }
                if mode == "all" || mode == "formats" {
                    styles.push((at, clip.styles[k].clone()));
                }
            }
        }
        self.apply(Batch(vec![Box::new(SetCells { sheet, edits }), Box::new(SetStyles { sheet, styles })]));
        vec![out_rows, out_cols]
    }

    /// True after `copy`, so the UI can enable Paste Special.
    #[wasm_bindgen(js_name = hasCopied)]
    pub fn has_copied(&self) -> bool {
        self.clip.is_some()
    }

    /// Shift a formula's relative references, as when it is copied.
    #[wasm_bindgen(js_name = shiftFormula)]
    pub fn shift_formula(input: &str, d_row: i32, d_col: i32) -> String {
        refs::shift_input(input, d_row as i64, d_col as i64)
    }

    // ---------------------------------------------------------------- functions

    /// Every built-in function as "NAME\tsyntax\tsummary", alphabetical.
    #[wasm_bindgen(js_name = functionList)]
    pub fn function_list() -> Vec<String> {
        sheets_engine::functions::FUNCTIONS.iter().map(|f| format!("{}\t{}\t{}", f.name, f.syntax, f.summary)).collect()
    }
}

fn parse_axis(axis: &str) -> Axis {
    if axis == "cols" {
        Axis::Col
    } else {
        Axis::Row
    }
}

fn parse_patch(key: &str, value: &str) -> Option<StylePatch> {
    let on = value == "1";
    let color = || (!value.is_empty()).then(|| value.to_string());
    Some(match key {
        "bold" => StylePatch::Bold(on),
        "italic" => StylePatch::Italic(on),
        "underline" => StylePatch::Underline(on),
        "strike" => StylePatch::Strike(on),
        "color" => StylePatch::Color(color()),
        "fill" => StylePatch::Fill(color()),
        "align" => StylePatch::Align(match value {
            "left" => Some(Align::Left),
            "center" => Some(Align::Center),
            "right" => Some(Align::Right),
            _ => None,
        }),
        "decimals" => StylePatch::Decimals(value.parse().ok()?),
        "clear" => StylePatch::Clear,
        "format" => StylePatch::NumberFormat(match value {
            "general" => NumberFormat::General,
            "number" => NumberFormat::Number { decimals: 2, thousands: true },
            "currency" => NumberFormat::Currency { symbol: "₹".into(), decimals: 2 },
            "dollar" => NumberFormat::Currency { symbol: "$".into(), decimals: 2 },
            "percent" => NumberFormat::Percent { decimals: 0 },
            "date" => NumberFormat::Date,
            "time" => NumberFormat::Time,
            "datetime" => NumberFormat::DateTime,
            "text" => NumberFormat::Text,
            _ => return None,
        }),
        _ => return None,
    })
}

/// Compact JSON for the UI. Keys: b i u s (flags), color, fill, align, fmt.
fn style_json(s: &Style) -> String {
    if s.is_default() {
        return String::new();
    }
    let mut parts = Vec::new();
    for (key, on) in [("b", s.bold), ("i", s.italic), ("u", s.underline), ("s", s.strike)] {
        if on {
            parts.push(format!("\"{key}\":1"));
        }
    }
    if let Some(c) = &s.color {
        parts.push(format!("\"color\":\"{c}\""));
    }
    if let Some(c) = &s.fill {
        parts.push(format!("\"fill\":\"{c}\""));
    }
    if let Some(a) = s.align {
        let a = match a {
            Align::Left => "left",
            Align::Center => "center",
            Align::Right => "right",
        };
        parts.push(format!("\"align\":\"{a}\""));
    }
    let fmt = match &s.number_format {
        NumberFormat::General => None,
        NumberFormat::Number { .. } => Some("number"),
        NumberFormat::Currency { .. } => Some("currency"),
        NumberFormat::Percent { .. } => Some("percent"),
        NumberFormat::Date => Some("date"),
        NumberFormat::Time => Some("time"),
        NumberFormat::DateTime => Some("datetime"),
        NumberFormat::Text => Some("text"),
    };
    if let Some(f) = fmt {
        parts.push(format!("\"fmt\":\"{f}\""));
    }
    format!("{{{}}}", parts.join(","))
}
