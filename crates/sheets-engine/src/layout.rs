//! Sheet layout: row heights, hidden rows and columns, merged cells and frozen
//! panes. (Column widths live in `sheet.rs`, next to the cells.)
//!
//! Row heights and hidden lines are stored as sorted runs of lines
//! (`Runs`), so "hide rows 100 to 1,048,576" or "make every row 30px tall"
//! costs one entry, not a million.
//!
//! Every change goes through an op at the bottom of this file
//! (`SetRowHeights`, `SetHidden`, `SetMerges`, `SetFreeze`) so it can be undone.

use office_core::Op;

use crate::cell::{CellRef, Range};
use crate::sheet::Sheet;
use crate::style::{Align, Style, StylePatch};
use crate::workbook::Workbook;

/// Row height in CSS pixels when none is set: Excel's default 15pt row.
pub const DEFAULT_ROW_HEIGHT: f64 = 20.0;

/// Tallest row Excel allows: 409 points.
pub const MAX_ROW_HEIGHT: f64 = 409.0 / 0.75;

/// Sorted, non-overlapping, inclusive runs of lines (rows or columns) that
/// share a value. Adjacent runs with equal values are merged.
#[derive(Debug, Clone, PartialEq)]
pub struct Runs<T> {
    runs: Vec<(u32, u32, T)>,
}

impl<T> Default for Runs<T> {
    fn default() -> Self {
        Runs { runs: Vec::new() }
    }
}

/// A set of lines, e.g. hidden rows.
pub type LineSet = Runs<()>;

impl<T> Runs<T> {
    /// An empty set, usable in `static`s.
    pub const fn new_const() -> Self {
        Runs { runs: Vec::new() }
    }
}

impl<T: Clone + PartialEq> Runs<T> {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn is_empty(&self) -> bool {
        self.runs.is_empty()
    }

    /// The value for line `i`, if it is inside a run.
    pub fn get(&self, i: u32) -> Option<&T> {
        let idx = self.runs.partition_point(|(start, _, _)| *start <= i);
        let (_, end, value) = self.runs.get(idx.checked_sub(1)?)?;
        (i <= *end).then_some(value)
    }

    pub fn contains(&self, i: u32) -> bool {
        self.get(i).is_some()
    }

    /// (first, last, value) for every run, ascending.
    pub fn iter(&self) -> impl Iterator<Item = (u32, u32, &T)> + '_ {
        self.runs.iter().map(|(s, e, v)| (*s, *e, v))
    }

    /// Set lines `first..=last` to `value`, or remove them with `None`.
    pub fn set(&mut self, first: u32, last: u32, value: Option<T>) {
        let (first, last) = (first.min(last), first.max(last));
        let mut out = Vec::with_capacity(self.runs.len() + 2);
        for (s, e, v) in self.runs.drain(..) {
            if e < first || s > last {
                out.push((s, e, v));
                continue;
            }
            if s < first {
                out.push((s, first - 1, v.clone()));
            }
            if e > last {
                out.push((last + 1, e, v));
            }
        }
        if let Some(v) = value {
            out.push((first, last, v));
        }
        self.runs = out;
        self.normalize();
    }

    /// Lines in `first..=last` that are in a run, as (first, last, value) pieces.
    pub fn within(&self, first: u32, last: u32) -> Vec<(u32, u32, T)> {
        self.runs
            .iter()
            .filter(|(s, e, _)| *e >= first && *s <= last)
            .map(|(s, e, v)| ((*s).max(first), (*e).min(last), v.clone()))
            .collect()
    }

    /// The runs after inserting `count` lines before line `at`. Lines past
    /// `limit` fall off the end. Inserted lines get no value.
    pub fn inserted(&self, at: u32, count: u32, limit: u32) -> Self {
        let mut out = Vec::new();
        let shift = |i: u32| i.saturating_add(count);
        for (s, e, v) in &self.runs {
            if *e < at {
                out.push((*s, *e, v.clone()));
            } else if *s >= at {
                out.push((shift(*s), shift(*e), v.clone()));
            } else {
                out.push((*s, at - 1, v.clone()));
                out.push((shift(at), shift(*e), v.clone()));
            }
        }
        let runs = out.into_iter().filter(|(s, _, _)| *s < limit).map(|(s, e, v)| (s, e.min(limit - 1), v)).collect();
        let mut r = Runs { runs };
        r.normalize();
        r
    }

    /// The runs after deleting lines `at..at + count`.
    pub fn deleted(&self, at: u32, count: u32) -> Self {
        let end = at.saturating_add(count).saturating_sub(1);
        let mut out = Vec::new();
        for (s, e, v) in &self.runs {
            if *s < at {
                out.push((*s, (*e).min(at - 1), v.clone()));
            }
            if *e > end {
                out.push(((*s).max(end + 1) - count, *e - count, v.clone()));
            }
        }
        let mut r = Runs { runs: out };
        r.normalize();
        r
    }

    fn normalize(&mut self) {
        self.runs.sort_by_key(|(s, _, _)| *s);
        let mut out: Vec<(u32, u32, T)> = Vec::with_capacity(self.runs.len());
        for (s, e, v) in self.runs.drain(..) {
            match out.last_mut() {
                Some(last) if last.1.checked_add(1) == Some(s) && last.2 == v => last.1 = e,
                _ => out.push((s, e, v)),
            }
        }
        self.runs = out;
    }
}

/// Frozen panes: the first `rows` rows and `cols` columns stay put while the rest scrolls.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Freeze {
    pub rows: u32,
    pub cols: u32,
}

/// How merged cells are created (Home → Merge & Center menu).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeKind {
    /// One merged cell for the whole range, text centred.
    Center,
    /// One merged cell per row of the range.
    Across,
    /// One merged cell for the whole range, alignment unchanged.
    Cells,
}

/// Row heights, hidden lines, merged cells and frozen panes of one sheet.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct SheetLayout {
    /// Rows whose height isn't `DEFAULT_ROW_HEIGHT`, in pixels.
    pub row_heights: Runs<f64>,
    /// Rows hidden by the user (Hide). Rows hidden by a filter are separate:
    /// see `Sheet::is_row_hidden`.
    pub hidden_rows: LineSet,
    pub hidden_cols: LineSet,
    /// Non-overlapping merged areas, each at least two cells, sorted.
    merges: Vec<Range>,
    pub freeze: Option<Freeze>,
}

fn intersects(a: &Range, b: &Range) -> bool {
    a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col
}

fn contains(r: &Range, at: CellRef) -> bool {
    at.row >= r.start.row && at.row <= r.end.row && at.col >= r.start.col && at.col <= r.end.col
}

/// Where line `i` goes after inserting `count` lines at `at` (merge edges grow).
fn grow(i: u32, at: u32, count: u32) -> u32 {
    if i >= at {
        i.saturating_add(count)
    } else {
        i
    }
}

impl SheetLayout {
    pub fn row_height(&self, row: u32) -> f64 {
        self.row_heights.get(row).copied().unwrap_or(DEFAULT_ROW_HEIGHT)
    }

    /// `None` (or the default height) resets the rows to the default height.
    pub fn set_row_height(&mut self, first: u32, last: u32, px: Option<f64>) {
        let px = px.map(|h| h.clamp(0.0, MAX_ROW_HEIGHT)).filter(|h| *h != DEFAULT_ROW_HEIGHT);
        self.row_heights.set(first, last, px);
    }

    pub fn merges(&self) -> &[Range] {
        &self.merges
    }

    /// The merged area containing `at`, if any.
    pub fn merge_at(&self, at: CellRef) -> Option<Range> {
        self.merges.iter().find(|m| contains(m, at)).copied()
    }

    /// Merged areas that overlap `range`.
    pub fn merges_in(&self, range: Range) -> impl Iterator<Item = &Range> + '_ {
        self.merges.iter().filter(move |m| intersects(m, &range))
    }

    /// Grow `range` until no merged area sticks out of it, like Excel's selection.
    pub fn expand_to_merges(&self, mut range: Range) -> Range {
        loop {
            let mut grown = range;
            for m in self.merges_in(range) {
                grown = Range::new(
                    CellRef::new(grown.start.row.min(m.start.row), grown.start.col.min(m.start.col)),
                    CellRef::new(grown.end.row.max(m.end.row), grown.end.col.max(m.end.col)),
                );
            }
            if grown == range {
                return range;
            }
            range = grown;
        }
    }

    /// Replace every merged area. Single cells and overlaps (later ones lose) are dropped.
    pub fn set_merges(&mut self, merges: impl IntoIterator<Item = Range>) {
        let mut out: Vec<Range> = Vec::new();
        for m in merges {
            if m.rows() * m.cols() > 1 && !out.iter().any(|o| intersects(o, &m)) {
                out.push(m);
            }
        }
        out.sort_by_key(|m| (m.start, m.end));
        self.merges = out;
    }

    /// The merged areas after merging `range` (replacing merges inside it).
    pub fn with_merge(&self, range: Range, kind: MergeKind) -> Vec<Range> {
        let mut merges: Vec<Range> = self.merges.iter().filter(|m| !intersects(m, &range)).copied().collect();
        match kind {
            MergeKind::Across => {
                for row in range.start.row..=range.end.row {
                    merges.push(Range::new(CellRef::new(row, range.start.col), CellRef::new(row, range.end.col)));
                }
            }
            MergeKind::Center | MergeKind::Cells => merges.push(range),
        }
        merges
    }

    /// The merged areas after unmerging everything that overlaps `range`.
    pub fn without_merges_in(&self, range: Range) -> Vec<Range> {
        self.merges.iter().filter(|m| !intersects(m, &range)).copied().collect()
    }

    /// Layout after inserting `count` rows (`rows = true`) or columns before `at`.
    pub fn inserted(&self, rows: bool, at: u32, count: u32, limit: u32) -> SheetLayout {
        let mut out = self.clone();
        let fit = |i: u32| i.min(limit - 1);
        if rows {
            out.row_heights = self.row_heights.inserted(at, count, limit);
            out.hidden_rows = self.hidden_rows.inserted(at, count, limit);
        } else {
            out.hidden_cols = self.hidden_cols.inserted(at, count, limit);
        }
        let merges = self.merges.iter().filter_map(|m| {
            let (s, e) = if rows { (m.start.row, m.end.row) } else { (m.start.col, m.end.col) };
            // An insert strictly inside a merge widens it; at its start, it moves it.
            let (ns, ne) = (grow(s, at, count), if e >= at { e.saturating_add(count) } else { e });
            if ns >= limit {
                return None;
            }
            let (ns, ne) = (ns, fit(ne));
            Some(if rows {
                Range::new(CellRef::new(ns, m.start.col), CellRef::new(ne, m.end.col))
            } else {
                Range::new(CellRef::new(m.start.row, ns), CellRef::new(m.end.row, ne))
            })
        });
        out.set_merges(merges.collect::<Vec<_>>());
        if let Some(f) = self.freeze {
            let n = if rows { f.rows } else { f.cols };
            let n = if at < n { fit(n.saturating_add(count)) } else { n };
            out.freeze = Some(if rows { Freeze { rows: n, ..f } } else { Freeze { cols: n, ..f } });
        }
        out
    }

    /// Layout after deleting `count` rows (`rows = true`) or columns starting at `at`.
    pub fn deleted(&self, rows: bool, at: u32, count: u32) -> SheetLayout {
        let mut out = self.clone();
        let end = at.saturating_add(count).saturating_sub(1);
        // New position of line `i`, or of the first line kept after it.
        let start_of = |i: u32| {
            if i < at {
                i
            } else if i > end {
                i - count
            } else {
                at
            }
        };
        if rows {
            out.row_heights = self.row_heights.deleted(at, count);
            out.hidden_rows = self.hidden_rows.deleted(at, count);
        } else {
            out.hidden_cols = self.hidden_cols.deleted(at, count);
        }
        let merges = self.merges.iter().filter_map(|m| {
            let (s, e) = if rows { (m.start.row, m.end.row) } else { (m.start.col, m.end.col) };
            if s >= at && e <= end {
                return None;
            }
            let ns = start_of(s);
            let ne = if e < at {
                e
            } else if e > end {
                e - count
            } else {
                at - 1
            };
            Some(if rows {
                Range::new(CellRef::new(ns, m.start.col), CellRef::new(ne, m.end.col))
            } else {
                Range::new(CellRef::new(m.start.row, ns), CellRef::new(m.end.row, ne))
            })
        });
        out.set_merges(merges.collect::<Vec<_>>());
        if let Some(f) = self.freeze {
            let n = if rows { f.rows } else { f.cols };
            let removed = if at < n { n.min(end + 1) - at } else { 0 };
            let n = n - removed;
            out.freeze = Some(if rows { Freeze { rows: n, ..f } } else { Freeze { cols: n, ..f } });
            if out.freeze.is_some_and(|f| f.rows == 0 && f.cols == 0) {
                out.freeze = None;
            }
        }
        out
    }
}

/// What merging a range changes: cell inputs, the kept cell's style, and the new merge list.
pub struct MergeEdits {
    pub cells: Vec<(CellRef, String)>,
    pub styles: Vec<(CellRef, Style)>,
    pub merges: Vec<Range>,
}

impl Sheet {
    /// The areas `range` would become with `kind` (one per row for Merge Across).
    fn merge_areas(range: Range, kind: MergeKind) -> Vec<Range> {
        match kind {
            MergeKind::Across => (range.start.row..=range.end.row)
                .map(|row| Range::new(CellRef::new(row, range.start.col), CellRef::new(row, range.end.col)))
                .collect(),
            _ => vec![range],
        }
    }

    /// Non-empty cells of `area` in reading order (only cells that exist, so huge areas are cheap).
    fn filled_in(&self, area: Range) -> Vec<CellRef> {
        let mut out: Vec<CellRef> = self
            .cells()
            .into_iter()
            .map(|(at, _, _)| at)
            .filter(|at| {
                at.row >= area.start.row && at.row <= area.end.row && at.col >= area.start.col && at.col <= area.end.col
            })
            .collect();
        out.sort();
        out
    }

    /// True when merging would throw away data: some area has more than one
    /// non-empty cell. Excel warns before doing that.
    pub fn merge_discards_data(&self, range: Range, kind: MergeKind) -> bool {
        Self::merge_areas(range, kind).into_iter().any(|a| a.cols() * a.rows() > 1 && self.filled_in(a).len() > 1)
    }

    /// Edits for Merge & Center / Merge Across / Merge Cells over `range`.
    /// Like Excel, each merged area keeps only its first non-empty value (in
    /// reading order), moved to the top-left cell; Merge & Center also centres it.
    pub fn merge_edits(&self, range: Range, kind: MergeKind) -> MergeEdits {
        let (mut cells, mut styles) = (Vec::new(), Vec::new());
        let areas = Self::merge_areas(range, kind);
        for area in areas.iter().filter(|a| a.rows() * a.cols() > 1) {
            let filled = self.filled_in(*area);
            if let Some(first) = filled.first() {
                if *first != area.start {
                    cells.push((area.start, self.input(*first).to_string()));
                }
                cells.extend(filled.iter().filter(|at| **at != area.start).map(|at| (*at, String::new())));
            }
            if kind == MergeKind::Center {
                styles.push((area.start, self.style(area.start).patched(&StylePatch::Align(Some(Align::Center)))));
            }
        }
        // Clearing must not wipe the value just moved to a top-left cell.
        let moved: std::collections::HashSet<CellRef> = areas.iter().map(|a| a.start).collect();
        cells.retain(|(at, input)| !(input.is_empty() && moved.contains(at)));
        let mut merges: Vec<Range> =
            self.layout().merges().iter().filter(|m| !intersects(m, &range)).copied().collect();
        merges.extend(areas.into_iter().filter(|a| a.rows() * a.cols() > 1));
        MergeEdits { cells, styles, merges }
    }

    /// Sorting can't move parts of merged cells, so Excel refuses to sort a
    /// range that overlaps a merged area.
    pub fn can_sort(&self, range: Range) -> bool {
        self.layout().merges_in(range).next().is_none()
    }
}

// ---------------------------------------------------------------- ops

/// Set row heights for runs of rows; `None` restores the default height.
#[derive(Debug, Clone)]
pub struct SetRowHeights {
    pub sheet: usize,
    /// (first row, last row, height)
    pub heights: Vec<(u32, u32, Option<f64>)>,
}

impl Op<Workbook> for SetRowHeights {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            for (first, last, px) in &self.heights {
                ws.sheet.layout_mut().set_row_height(*first, *last, *px);
            }
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let layout = doc.sheets[self.sheet].sheet.layout();
        // Reset each run, then put back the custom heights that were there.
        let mut heights: Vec<(u32, u32, Option<f64>)> = self.heights.iter().map(|(f, l, _)| (*f, *l, None)).collect();
        for (f, l, _) in &self.heights {
            heights.extend(layout.row_heights.within(*f, *l).into_iter().map(|(a, b, h)| (a, b, Some(h))));
        }
        Box::new(SetRowHeights { sheet: self.sheet, heights })
    }
}

/// Hide or unhide runs of rows or columns.
#[derive(Debug, Clone)]
pub struct SetHidden {
    pub sheet: usize,
    pub rows: bool,
    /// (first, last, hidden)
    pub lines: Vec<(u32, u32, bool)>,
}

impl Op<Workbook> for SetHidden {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            let layout = ws.sheet.layout_mut();
            let set = if self.rows { &mut layout.hidden_rows } else { &mut layout.hidden_cols };
            for (first, last, hidden) in &self.lines {
                set.set(*first, *last, hidden.then_some(()));
            }
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let layout = doc.sheets[self.sheet].sheet.layout();
        let set = if self.rows { &layout.hidden_rows } else { &layout.hidden_cols };
        let mut lines: Vec<(u32, u32, bool)> = self.lines.iter().map(|(f, l, _)| (*f, *l, false)).collect();
        for (f, l, _) in &self.lines {
            lines.extend(set.within(*f, *l).into_iter().map(|(a, b, _)| (a, b, true)));
        }
        Box::new(SetHidden { sheet: self.sheet, rows: self.rows, lines })
    }
}

/// Replace a sheet's merged areas.
#[derive(Debug, Clone)]
pub struct SetMerges {
    pub sheet: usize,
    pub merges: Vec<Range>,
}

impl Op<Workbook> for SetMerges {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            ws.sheet.layout_mut().set_merges(self.merges.iter().copied());
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(SetMerges { sheet: self.sheet, merges: doc.sheets[self.sheet].sheet.layout().merges().to_vec() })
    }
}

/// Freeze panes (`None` unfreezes).
#[derive(Debug, Clone)]
pub struct SetFreeze {
    pub sheet: usize,
    pub freeze: Option<Freeze>,
}

impl Op<Workbook> for SetFreeze {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            ws.sheet.layout_mut().freeze = self.freeze.filter(|f| f.rows > 0 || f.cols > 0);
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(SetFreeze { sheet: self.sheet, freeze: doc.sheets[self.sheet].sheet.layout().freeze })
    }
}

#[cfg(test)]
mod tests {
    use office_core::history::History;

    use super::*;

    fn r(s: &str) -> Range {
        let (a, b) = s.split_once(':').unwrap();
        Range::new(CellRef::parse(a).unwrap(), CellRef::parse(b).unwrap())
    }

    #[test]
    fn runs_set_split_and_merge() {
        let mut runs: Runs<f64> = Runs::new();
        runs.set(10, 20, Some(30.0));
        runs.set(21, 25, Some(30.0));
        assert_eq!(runs.iter().count(), 1, "adjacent equal runs merge");
        runs.set(15, 16, Some(50.0));
        assert_eq!(runs.get(14), Some(&30.0));
        assert_eq!(runs.get(15), Some(&50.0));
        assert_eq!(runs.get(17), Some(&30.0));
        assert_eq!(runs.get(26), None);
        assert_eq!(runs.get(9), None);
        runs.set(0, 100, None);
        assert!(runs.is_empty());
    }

    #[test]
    fn runs_cover_a_million_rows_cheaply() {
        let mut hidden = LineSet::new();
        hidden.set(100, crate::MAX_ROWS - 1, Some(()));
        assert!(hidden.contains(crate::MAX_ROWS - 1));
        assert!(!hidden.contains(99));
        assert_eq!(hidden.iter().count(), 1);
    }

    #[test]
    fn runs_follow_inserts_and_deletes() {
        let mut runs = LineSet::new();
        runs.set(5, 9, Some(()));
        let ins = runs.inserted(7, 3, 1000);
        assert_eq!(ins.iter().map(|(s, e, _)| (s, e)).collect::<Vec<_>>(), vec![(5, 6), (10, 12)]);
        let del = runs.deleted(3, 4);
        assert_eq!(del.iter().map(|(s, e, _)| (s, e)).collect::<Vec<_>>(), vec![(3, 5)]);
        let del = runs.deleted(5, 5);
        assert!(del.is_empty());
        assert!(runs.inserted(0, 995, 1000).is_empty(), "pushed past the last row");
    }

    #[test]
    fn merges_expand_selection_and_follow_structure() {
        let mut l = SheetLayout::default();
        l.set_merges([r("B2:C3"), r("C4:D4"), r("A1:A1")]);
        assert_eq!(l.merges().len(), 2, "single-cell merges are dropped");
        assert_eq!(l.merge_at(CellRef::parse("C3").unwrap()), Some(r("B2:C3")));
        assert_eq!(l.expand_to_merges(r("B3:B4")), r("B2:D4"));

        let ins = l.inserted(true, 2, 2, crate::MAX_ROWS);
        assert_eq!(ins.merges(), &[r("B2:C5"), r("C6:D6")]);
        let del = l.deleted(true, 1, 1);
        assert_eq!(del.merges(), &[r("B2:C2"), r("C3:D3")]);
        let del = l.deleted(false, 2, 1);
        assert_eq!(del.merges(), &[r("B2:B3")], "C4:D4 shrinks to one cell and goes");
    }

    #[test]
    fn merge_kinds() {
        let l = SheetLayout::default();
        assert_eq!(l.with_merge(r("A1:C2"), MergeKind::Across), vec![r("A1:C1"), r("A2:C2")]);
        assert_eq!(l.with_merge(r("A1:C2"), MergeKind::Center), vec![r("A1:C2")]);
    }

    #[test]
    fn freeze_follows_structure() {
        let l = SheetLayout { freeze: Some(Freeze { rows: 2, cols: 1 }), ..Default::default() };
        assert_eq!(l.inserted(true, 0, 3, 100).freeze, Some(Freeze { rows: 5, cols: 1 }));
        assert_eq!(l.inserted(true, 2, 3, 100).freeze, Some(Freeze { rows: 2, cols: 1 }));
        assert_eq!(l.deleted(false, 0, 5).freeze, Some(Freeze { rows: 2, cols: 0 }));
        assert_eq!(l.deleted(true, 0, 9).deleted(false, 0, 1).freeze, None);
    }

    #[test]
    fn merging_keeps_the_first_value() {
        let mut s = Sheet::new();
        s.set_inputs([(CellRef::new(0, 1), "b".to_string()), (CellRef::new(1, 0), "c".to_string())]);
        let range = r("A1:B2");
        assert!(s.merge_discards_data(range, MergeKind::Center));
        assert!(!s.merge_discards_data(r("A1:B1"), MergeKind::Center));
        let edits = s.merge_edits(range, MergeKind::Center);
        assert_eq!(
            edits.cells,
            vec![
                (CellRef::new(0, 0), "b".to_string()),
                (CellRef::new(0, 1), String::new()),
                (CellRef::new(1, 0), String::new())
            ]
        );
        assert_eq!(edits.styles[0].1.align, Some(Align::Center));
        assert_eq!(edits.merges, vec![range]);
        let across = s.merge_edits(range, MergeKind::Across);
        assert_eq!(across.merges, vec![r("A1:B1"), r("A2:B2")]);
        assert_eq!(across.cells, vec![(CellRef::new(0, 0), "b".to_string()), (CellRef::new(0, 1), String::new())]);
        s.layout_mut().set_merges(edits.merges);
        assert!(!s.can_sort(r("A1:A9")));
        assert!(s.can_sort(r("C1:D9")));
    }

    #[test]
    fn layout_ops_undo() {
        let mut wb = Workbook::new();
        let mut h = History::default();
        h.apply(&mut wb, Box::new(SetRowHeights { sheet: 0, heights: vec![(2, 4, Some(40.0))] }));
        h.apply(&mut wb, Box::new(SetRowHeights { sheet: 0, heights: vec![(3, 10, Some(60.0))] }));
        h.apply(&mut wb, Box::new(SetHidden { sheet: 0, rows: false, lines: vec![(1, 2, true)] }));
        h.apply(&mut wb, Box::new(SetMerges { sheet: 0, merges: vec![r("A1:B2")] }));
        h.apply(&mut wb, Box::new(SetFreeze { sheet: 0, freeze: Some(Freeze { rows: 1, cols: 0 }) }));
        let l = wb.sheets[0].sheet.layout();
        assert_eq!((l.row_height(2), l.row_height(3), l.row_height(11)), (40.0, 60.0, DEFAULT_ROW_HEIGHT));
        assert!(l.hidden_cols.contains(2));
        for _ in 0..4 {
            h.undo(&mut wb);
        }
        let l = wb.sheets[0].sheet.layout();
        assert_eq!((l.row_height(2), l.row_height(3), l.row_height(5)), (40.0, 40.0, DEFAULT_ROW_HEIGHT));
        assert!(!l.hidden_cols.contains(2) && l.merges().is_empty() && l.freeze.is_none());
        h.undo(&mut wb);
        assert_eq!(wb.sheets[0].sheet.layout(), &SheetLayout::default());
        h.redo(&mut wb);
        h.redo(&mut wb);
        assert_eq!(wb.sheets[0].sheet.layout().row_height(10), 60.0);
    }
}
