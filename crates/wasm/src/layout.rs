//! Bindings for the sheet layout: row heights, hidden rows/columns, merged
//! cells and frozen panes (`sheets_engine::layout`). Same conventions as `lib.rs`.

use sheets_engine::layout::{Freeze, MergeKind, SetFreeze, SetHidden, SetMerges, SetRowHeights};
use sheets_engine::ops::{Batch, SetCells, SetStyles};
use sheets_engine::{CellRef, Range, MAX_COLS, MAX_ROWS};
use wasm_bindgen::prelude::*;

use crate::{rect, SheetsDocument};

fn merge_kind(kind: &str) -> MergeKind {
    match kind {
        "across" => MergeKind::Across,
        "cells" => MergeKind::Cells,
        _ => MergeKind::Center,
    }
}

#[wasm_bindgen]
impl SheetsDocument {
    /// Row layout for the grid, flattened: `[n, (first, last, px) × n, (first, last) × rest]`.
    /// The first `n` runs are custom heights; the rest are hidden rows (hidden
    /// by the user or by a filter, see `Sheet::is_row_hidden`).
    #[wasm_bindgen(js_name = rowLayout)]
    pub fn row_layout(&self, sheet: usize) -> Vec<f64> {
        let s = self.sheet(sheet);
        let layout = s.layout();
        let mut out = vec![layout.row_heights.iter().count() as f64];
        for (first, last, px) in layout.row_heights.iter() {
            out.extend([first as f64, last as f64, *px]);
        }
        let mut hidden = layout.hidden_rows.clone();
        for &row in s.filter_hidden_rows() {
            hidden.set(row, row, Some(()));
        }
        for (first, last, _) in hidden.iter() {
            out.extend([first as f64, last as f64]);
        }
        out
    }

    /// Rows whose automatic height the UI should measure (wrapped text, big
    /// fonts, line breaks); see `Sheet::rows_to_measure`.
    #[wasm_bindgen(js_name = rowsToMeasure)]
    pub fn rows_to_measure(&self, sheet: usize) -> Vec<u32> {
        self.sheet(sheet).rows_to_measure()
    }

    /// Hidden columns as `[first, last, first, last, ...]`.
    #[wasm_bindgen(js_name = hiddenColumns)]
    pub fn hidden_columns(&self, sheet: usize) -> Vec<u32> {
        self.sheet(sheet).layout().hidden_cols.iter().flat_map(|(f, l, _)| [f, l]).collect()
    }

    /// Set the height of rows `first..=last` in pixels; a negative `px` restores the default height.
    #[wasm_bindgen(js_name = setRowHeight)]
    pub fn set_row_height(&mut self, sheet: usize, first: u32, last: u32, px: f64) {
        let px = (px >= 0.0).then_some(px);
        self.apply(SetRowHeights { sheet, heights: vec![(first, last.min(MAX_ROWS - 1), px)] });
    }

    /// Set several rows' heights as one undo step (AutoFit). Parallel arrays;
    /// a negative height restores the default.
    #[wasm_bindgen(js_name = setRowHeights)]
    pub fn set_row_heights(&mut self, sheet: usize, rows: Vec<u32>, px: Vec<f64>) {
        let heights = rows.into_iter().zip(px).map(|(r, h)| (r, r, (h >= 0.0).then_some(h))).collect();
        self.apply(SetRowHeights { sheet, heights });
    }

    /// Hide or unhide rows (`axis` "rows") or columns ("cols") `first..=last`.
    #[wasm_bindgen(js_name = setHidden)]
    pub fn set_hidden(&mut self, sheet: usize, axis: &str, first: u32, last: u32, hidden: bool) {
        let rows = axis != "cols";
        let last = last.min(if rows { MAX_ROWS } else { MAX_COLS } - 1);
        self.apply(SetHidden { sheet, rows, lines: vec![(first, last, hidden)] });
    }

    /// Merged areas as `[top, left, bottom, right, ...]`.
    pub fn merges(&self, sheet: usize) -> Vec<u32> {
        self.sheet(sheet)
            .layout()
            .merges()
            .iter()
            .flat_map(|m| [m.start.row, m.start.col, m.end.row, m.end.col])
            .collect()
    }

    /// True when merging would keep only the first value of an area (Excel warns first).
    /// `kind`: "center" | "across" | "cells".
    #[wasm_bindgen(js_name = mergeDiscardsData)]
    pub fn merge_discards_data(&self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32, kind: &str) -> bool {
        self.sheet(sheet).merge_discards_data(rect(top, left, bottom, right), merge_kind(kind))
    }

    /// Merge & Center ("center"), Merge Across ("across") or Merge Cells ("cells").
    #[wasm_bindgen(js_name = mergeCells)]
    pub fn merge_cells(&mut self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32, kind: &str) {
        let edits = self.sheet(sheet).merge_edits(rect(top, left, bottom, right), merge_kind(kind));
        self.apply(Batch(vec![
            Box::new(SetCells { sheet, edits: edits.cells }),
            Box::new(SetStyles { sheet, styles: edits.styles }),
            Box::new(SetMerges { sheet, merges: edits.merges }),
        ]));
    }

    /// Unmerge every merged area that overlaps the rectangle.
    #[wasm_bindgen(js_name = unmergeCells)]
    pub fn unmerge_cells(&mut self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32) {
        let merges = self.sheet(sheet).layout().without_merges_in(rect(top, left, bottom, right));
        self.apply(SetMerges { sheet, merges });
    }

    /// False when the rectangle overlaps merged cells, which can't be sorted.
    #[wasm_bindgen(js_name = canSort)]
    pub fn can_sort(&self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32) -> bool {
        self.sheet(sheet).can_sort(rect(top, left, bottom, right))
    }

    /// Frozen `[rows, cols]`, or empty when nothing is frozen.
    /// Frozen panes as [first scrolling row, first scrolling column, first frozen row,
    /// first frozen column], or empty when nothing is frozen (see `Freeze`).
    #[wasm_bindgen(js_name = freezePanes)]
    pub fn freeze_panes(&self, sheet: usize) -> Vec<u32> {
        self.sheet(sheet).layout().freeze.map_or(Vec::new(), |f| {
            let (rows, cols) = f.split();
            vec![rows, cols, f.top, f.left]
        })
    }

    /// Freeze rows `top..rows` and columns `left..cols` (`rows`/`cols`: the first
    /// line that scrolls; `top`/`left`: the first visible line when freezing).
    #[wasm_bindgen(js_name = setFreezePanes)]
    pub fn set_freeze_panes(&mut self, sheet: usize, rows: u32, cols: u32, top: u32, left: u32) {
        let top = if rows > top { top.min(MAX_ROWS - 1) } else { 0 };
        let left = if cols > left { left.min(MAX_COLS - 1) } else { 0 };
        let freeze = Freeze {
            rows: rows.min(MAX_ROWS - 1).saturating_sub(top),
            cols: cols.min(MAX_COLS - 1).saturating_sub(left),
            top,
            left,
        };
        self.apply(SetFreeze { sheet, freeze: (freeze.rows > 0 || freeze.cols > 0).then_some(freeze) });
    }
}

/// Merged areas fully inside `copied`, relative to its top-left, for the clipboard.
pub(crate) fn copied_merges(merges: &[Range], copied: Range) -> Vec<Range> {
    merges
        .iter()
        .filter(|m| {
            m.start.row >= copied.start.row
                && m.end.row <= copied.end.row
                && m.start.col >= copied.start.col
                && m.end.col <= copied.end.col
        })
        .map(|m| {
            let rel = |c: CellRef| CellRef::new(c.row - copied.start.row, c.col - copied.start.col);
            Range::new(rel(m.start), rel(m.end))
        })
        .collect()
}

/// The sheet's merges after pasting `clip` merges (relative) at `at`, covering `size` cells:
/// merges overlapping the pasted area go, the copied ones are added.
pub(crate) fn pasted_merges(
    existing: &[Range],
    clip: &[Range],
    at: CellRef,
    size: (u32, u32),
    transpose: bool,
) -> Vec<Range> {
    let area = Range::new(at, CellRef::new(at.row + size.0.max(1) - 1, at.col + size.1.max(1) - 1));
    let overlaps = |m: &Range| {
        m.start.row <= area.end.row
            && area.start.row <= m.end.row
            && m.start.col <= area.end.col
            && area.start.col <= m.end.col
    };
    let mut out: Vec<Range> = existing.iter().filter(|m| !overlaps(m)).copied().collect();
    for m in clip {
        let place = |c: CellRef| {
            let (r, k) = if transpose { (c.col, c.row) } else { (c.row, c.col) };
            CellRef::new((at.row + r).min(MAX_ROWS - 1), (at.col + k).min(MAX_COLS - 1))
        };
        out.push(Range::new(place(m.start), place(m.end)));
    }
    out
}
