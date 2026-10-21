//! Sheet layout in `.xlsx` worksheets: row heights and hidden rows
//! (`<row ht customHeight hidden>`), hidden columns (`<col hidden>`), merged
//! cells (`<mergeCells>`) and frozen panes (`<sheetViews><pane state="frozen">`).
//!
//! A child module of `xlsx.rs` so it can share its XML helpers.
//!
//! Units: Excel row heights are points (1pt = 4/3 px); column widths are
//! characters (see `chars_to_px`).
//!
//! When the rows from some row to the end of the sheet are hidden (a common
//! way to "tidy" a sheet), Excel writes `<sheetFormatPr zeroHeight="1">` and
//! lists only the visible rows. We read that, and write it when it is smaller.

use sheets_engine::layout::{Freeze, LineSet, SheetLayout, DEFAULT_ROW_HEIGHT};
use sheets_engine::{CellRef, Range, Sheet, MAX_COLS, MAX_ROWS};

use super::px_to_chars;
use crate::xml::{attr, Attrs, XmlEvent};

const PX_PER_PT: f64 = 4.0 / 3.0;

fn on(attrs: &Attrs, key: &str) -> bool {
    matches!(&*attr(attrs, key), "1" | "true")
}

/// Reads row heights, hidden lines, merges and frozen panes while
/// `xlsx::sheet_data` walks the worksheet, so a big sheet is read in one pass.
pub(super) struct LayoutReader {
    layout: SheetLayout,
    default_ht: f64,
    zero_height: bool,
    /// Rows listed in <sheetData>, for zeroHeight sheets.
    listed: Vec<u32>,
    merges: Vec<Range>,
    next_row: u32,
}

impl Default for LayoutReader {
    fn default() -> Self {
        LayoutReader {
            layout: SheetLayout::default(),
            default_ht: 15.0,
            zero_height: false,
            listed: Vec::new(),
            merges: Vec::new(),
            next_row: 0,
        }
    }
}

impl LayoutReader {
    pub(super) fn event(&mut self, ev: &XmlEvent) {
        let XmlEvent::Start { name, attrs } = ev else { return };
        match *name {
            "sheetFormatPr" => {
                self.default_ht = attr(attrs, "defaultRowHeight").parse().unwrap_or(15.0);
                self.zero_height = on(attrs, "zeroHeight");
            }
            "row" => {
                let row = attr(attrs, "r").parse::<u32>().map_or(self.next_row, |r| r.saturating_sub(1));
                self.next_row = row + 1;
                if row >= MAX_ROWS {
                    return;
                }
                self.listed.push(row);
                if on(attrs, "hidden") {
                    self.layout.hidden_rows.set(row, row, Some(()));
                }
                if let Ok(ht) = attr(attrs, "ht").parse::<f64>() {
                    // Custom heights, plus rows Excel grew to fit (wrapped text, big fonts).
                    if on(attrs, "customHeight") || ht > self.default_ht + 0.5 {
                        self.layout.set_row_height(row, row, Some(ht * PX_PER_PT));
                    }
                }
            }
            "col" if on(attrs, "hidden") => {
                let (Ok(min), Ok(max)) = (attr(attrs, "min").parse::<u32>(), attr(attrs, "max").parse::<u32>()) else {
                    return;
                };
                if min >= 1 && max >= min && min <= MAX_COLS {
                    self.layout.hidden_cols.set(min - 1, max.min(MAX_COLS) - 1, Some(()));
                }
            }
            "mergeCell" => {
                if let Some((a, b)) = attr(attrs, "ref").split_once(':') {
                    if let (Some(a), Some(b)) = (CellRef::parse(a), CellRef::parse(b)) {
                        self.merges.push(Range::new(a, b));
                    }
                }
            }
            "pane" if matches!(&*attr(attrs, "state"), "frozen" | "frozenSplit") => {
                let split = |key: &str| attr(attrs, key).parse::<f64>().map_or(0, |n| n.max(0.0) as u32);
                let freeze =
                    Freeze { rows: split("ySplit").min(MAX_ROWS - 1), cols: split("xSplit").min(MAX_COLS - 1) };
                self.layout.freeze = (freeze.rows > 0 || freeze.cols > 0).then_some(freeze);
            }
            _ => {}
        }
    }

    pub(super) fn finish(self) -> SheetLayout {
        let Self { mut layout, zero_height, listed, merges, .. } = self;
        Self::finish_parts(&mut layout, zero_height, listed, merges);
        layout
    }

    fn finish_parts(layout: &mut SheetLayout, zero_height: bool, mut listed: Vec<u32>, merges: Vec<Range>) {
        if zero_height {
            // Every row that isn't listed is hidden.
            listed.sort_unstable();
            listed.dedup();
            let mut from = 0;
            for row in listed.into_iter().chain([MAX_ROWS]) {
                if row > from {
                    layout.hidden_rows.set(from, row - 1, Some(()));
                }
                from = row + 1;
            }
        }
        layout.set_merges(merges);
    }
}

/// `<sheetViews>` with the frozen pane, or "" when nothing is frozen.
/// Goes right after `<worksheet>` (before `<sheetFormatPr>` and `<cols>`).
pub(super) fn sheet_views_xml(sheet: &Sheet) -> String {
    let Some(Freeze { rows, cols }) = sheet.layout().freeze else { return String::new() };
    let pane = match (rows > 0, cols > 0) {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        _ => "topRight",
    };
    let x = if cols > 0 { format!(r#" xSplit="{cols}""#) } else { String::new() };
    let y = if rows > 0 { format!(r#" ySplit="{rows}""#) } else { String::new() };
    let top_left = CellRef::new(rows, cols);
    format!(
        r#"<sheetViews><sheetView workbookViewId="0"><pane{x}{y} topLeftCell="{top_left}" activePane="{pane}" state="frozen"/><selection pane="{pane}" activeCell="{top_left}" sqref="{top_left}"/></sheetView></sheetViews>"#
    )
}

/// Where the zeroHeight trick starts: the first row of a hidden run that
/// reaches the last row, when listing the rows above it is shorter than
/// listing the hidden ones.
fn hidden_tail(sheet: &Sheet) -> Option<u32> {
    let (start, end, _) = sheet.layout().hidden_rows.iter().last()?;
    (end == MAX_ROWS - 1 && MAX_ROWS - start > start).then_some(start)
}

/// `<sheetFormatPr>` when the sheet needs one, or "". Goes before `<cols>`.
pub(super) fn sheet_format_xml(sheet: &Sheet) -> String {
    match hidden_tail(sheet) {
        Some(_) => r#"<sheetFormatPr defaultRowHeight="15" zeroHeight="1"/>"#.to_string(),
        None => String::new(),
    }
}

/// `<cols>` with custom widths and hidden columns, or "".
pub(super) fn cols_xml(sheet: &Sheet) -> String {
    let hidden = &sheet.layout().hidden_cols;
    // Every column with a custom width or hidden, then runs of equal (width, hidden).
    let mut lines: Vec<u32> = sheet.col_widths().map(|(c, _)| c).collect();
    for (first, last, _) in hidden.iter() {
        lines.extend(first..=last);
    }
    lines.sort_unstable();
    lines.dedup();
    let mut xml = String::new();
    let mut i = 0;
    while i < lines.len() {
        let col = lines[i];
        let key = (sheet.col_width(col), hidden.contains(col));
        let mut last = col;
        while i + 1 < lines.len()
            && lines[i + 1] == last + 1
            && (sheet.col_width(last + 1), hidden.contains(last + 1)) == key
        {
            i += 1;
            last += 1;
        }
        let custom = if key.0 != sheets_engine::sheet::DEFAULT_COL_WIDTH { r#" customWidth="1""# } else { "" };
        let hide = if key.1 { r#" hidden="1""# } else { "" };
        xml.push_str(&format!(
            r#"<col min="{}" max="{}" width="{}"{hide}{custom}/>"#,
            col + 1,
            last + 1,
            px_to_chars(key.0)
        ));
        i += 1;
    }
    if xml.is_empty() {
        xml
    } else {
        format!("<cols>{xml}</cols>")
    }
}

/// Emits `<row>` elements in order: rows with cells (opened by the caller)
/// and rows that only carry a height or hidden flag.
pub(super) struct RowWriter<'a> {
    sheet: &'a Sheet,
    /// Rows that need a `<row>` for layout alone, ascending, after `next`.
    pending: Vec<u32>,
    next: usize,
}

impl<'a> RowWriter<'a> {
    pub(super) fn new(sheet: &'a Sheet) -> Self {
        let layout = sheet.layout();
        let tail = hidden_tail(sheet);
        let mut pending: Vec<u32> = Vec::new();
        let mut add = |set: &LineSet| {
            for (first, last, _) in set.iter() {
                pending.extend(first..=last.min(tail.map_or(u32::MAX, |t| t.saturating_sub(1))));
            }
        };
        add(&layout.hidden_rows);
        for (first, last, _) in layout.row_heights.iter() {
            pending.extend(first..=last.min(tail.map_or(u32::MAX, |t| t.saturating_sub(1))));
        }
        if let Some(t) = tail {
            // With zeroHeight, every row above the hidden tail must be listed.
            pending.extend(0..t);
        }
        pending.sort_unstable();
        pending.dedup();
        RowWriter { sheet, pending, next: 0 }
    }

    /// Write layout-only rows before `row`.
    pub(super) fn before(&mut self, row: u32, xml: &mut String) {
        while self.next < self.pending.len() && self.pending[self.next] < row {
            let r = self.pending[self.next];
            xml.push_str(&format!("<row {}/>", self.attrs(r)));
            self.next += 1;
        }
        if self.pending.get(self.next) == Some(&row) {
            self.next += 1;
        }
    }

    /// Write the remaining layout-only rows.
    pub(super) fn rest(&mut self, xml: &mut String) {
        self.before(u32::MAX, xml);
    }

    /// `<row r="N" ...>` with height and hidden attributes.
    pub(super) fn open(&self, row: u32) -> String {
        format!("<row {}>", self.attrs(row))
    }

    fn attrs(&self, row: u32) -> String {
        let layout = self.sheet.layout();
        let mut attrs = format!(r#"r="{}""#, row + 1);
        let ht = layout.row_height(row);
        if ht != DEFAULT_ROW_HEIGHT {
            attrs.push_str(&format!(r#" ht="{}" customHeight="1""#, (ht * 0.75 * 100.0).round() / 100.0));
        }
        if layout.hidden_rows.contains(row) {
            attrs.push_str(r#" hidden="1""#);
        }
        attrs
    }
}

/// `<mergeCells>`, or "". Goes right after `</sheetData>`.
pub(super) fn merge_cells_xml(sheet: &Sheet) -> String {
    let merges = sheet.layout().merges();
    if merges.is_empty() {
        return String::new();
    }
    let mut xml = format!(r#"<mergeCells count="{}">"#, merges.len());
    for m in merges {
        xml.push_str(&format!(r#"<mergeCell ref="{}:{}"/>"#, m.start, m.end));
    }
    xml.push_str("</mergeCells>");
    xml
}
