//! Per-sheet data-tool settings: AutoFilter, data validation and conditional
//! formatting, kept together in one [`DataTools`] value on each [`Sheet`].
//!
//! The settings are document state (saved in `.xlsx`, undoable) but not cell
//! content, so they change through one op, [`SetDataTools`], whose inverse is
//! simply the previous value. Structural edits (insert/delete rows and
//! columns) move the ranges with the cells via [`DataTools::remapped`].
//!
//! The logic lives in the sibling modules: `filter.rs`, `validation.rs`,
//! `conditional_format.rs`, plus `sort.rs` (multi-level sort) and `dedupe.rs`
//! (Remove Duplicates).

use std::collections::BTreeSet;

use office_core::Op;

use crate::cell::{CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::conditional_format::CfRule;
use crate::filter::AutoFilter;
use crate::formula::refs::Axis;
use crate::sheet::Sheet;
use crate::validation::Validation;
use crate::workbook::Workbook;

/// Everything the Data tools store for one sheet.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DataTools {
    /// The AutoFilter range and its column criteria (at most one per sheet, as in Excel).
    pub auto_filter: Option<AutoFilter>,
    /// Rows hidden by the AutoFilter. Like Excel, this is only recomputed when a
    /// filter is applied or reapplied, not on every edit.
    pub filter_hidden: BTreeSet<u32>,
    /// Data validation rules; earlier rules win where ranges overlap.
    pub validations: Vec<Validation>,
    /// Conditional formatting rules in priority order (index 0 = highest priority).
    pub cond_formats: Vec<CfRule>,
}

impl DataTools {
    pub fn is_empty(&self) -> bool {
        *self == DataTools::default()
    }

    /// A copy with every range moved by an insert/delete of rows or columns.
    /// `map` gives each old index its new one, or `None` when it was deleted.
    pub fn remapped(&self, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> DataTools {
        let auto_filter = self.auto_filter.as_ref().and_then(|f| f.remapped(axis, map));
        let filter_hidden = match axis {
            Axis::Row => self.filter_hidden.iter().filter_map(|r| map(*r).filter(|r| *r < MAX_ROWS)).collect(),
            Axis::Col => self.filter_hidden.clone(),
        };
        let validations = self
            .validations
            .iter()
            .filter_map(|v| {
                let ranges = remap_ranges(&v.ranges, axis, map);
                (!ranges.is_empty()).then(|| Validation { ranges, ..v.clone() })
            })
            .collect();
        let cond_formats = self
            .cond_formats
            .iter()
            .filter_map(|r| {
                let ranges = remap_ranges(&r.ranges, axis, map);
                (!ranges.is_empty()).then(|| CfRule { ranges, ..r.clone() })
            })
            .collect();
        DataTools { auto_filter, filter_hidden, validations, cond_formats }
    }
}

// ------------------------------------------------------------ sheet access

impl Sheet {
    pub fn data_tools(&self) -> &DataTools {
        &self.tools
    }

    pub fn set_data_tools(&mut self, tools: DataTools) {
        self.tools = tools;
    }

    /// Rows hidden by the AutoFilter, ascending. Manual hiding (Hide rows) is
    /// separate; `Sheet::is_row_hidden` combines both.
    pub fn filter_hidden_rows(&self) -> &BTreeSet<u32> {
        &self.tools.filter_hidden
    }

    pub fn is_filter_hidden(&self, row: u32) -> bool {
        self.tools.filter_hidden.contains(&row)
    }
}

/// Replace a sheet's data-tool settings (filter, validation, conditional formats).
#[derive(Debug, Clone)]
pub struct SetDataTools {
    pub sheet: usize,
    pub tools: DataTools,
}

impl Op<Workbook> for SetDataTools {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            ws.sheet.set_data_tools(self.tools.clone());
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(SetDataTools { sheet: self.sheet, tools: doc.sheets[self.sheet].sheet.data_tools().clone() })
    }
}

// ------------------------------------------------------------ range helpers

/// Move one range through an insert/delete map. Deleted edges snap to the
/// nearest surviving row/column inside the range; `None` if nothing survives.
pub fn remap_range(r: Range, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Option<Range> {
    let (lo, hi, limit) = match axis {
        Axis::Row => (r.start.row, r.end.row, MAX_ROWS),
        Axis::Col => (r.start.col, r.end.col, MAX_COLS),
    };
    // Deleted indices form one contiguous block, so these loops are short.
    let mut a = lo;
    let new_lo = loop {
        if a > hi {
            return None;
        }
        if let Some(n) = map(a) {
            break n;
        }
        a += 1;
    };
    let mut b = hi;
    let new_hi = loop {
        if let Some(n) = map(b) {
            break n;
        }
        b -= 1;
    };
    if new_lo >= limit {
        return None;
    }
    let new_hi = new_hi.min(limit - 1);
    Some(match axis {
        Axis::Row => Range::new(CellRef::new(new_lo, r.start.col), CellRef::new(new_hi, r.end.col)),
        Axis::Col => Range::new(CellRef::new(r.start.row, new_lo), CellRef::new(r.end.row, new_hi)),
    })
}

pub fn remap_ranges(ranges: &[Range], axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Vec<Range> {
    ranges.iter().filter_map(|r| remap_range(*r, axis, map)).collect()
}

pub fn contains(r: &Range, at: CellRef) -> bool {
    at.row >= r.start.row && at.row <= r.end.row && at.col >= r.start.col && at.col <= r.end.col
}

pub fn intersects(a: &Range, b: &Range) -> bool {
    a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col
}

/// `r` minus `cut`, as up to four rectangles (above, below, left, right).
pub fn subtract(r: Range, cut: &Range) -> Vec<Range> {
    if !intersects(&r, cut) {
        return vec![r];
    }
    let mut out = Vec::new();
    let (top, bottom) = (cut.start.row.max(r.start.row), cut.end.row.min(r.end.row));
    if r.start.row < top {
        out.push(Range::new(r.start, CellRef::new(top - 1, r.end.col)));
    }
    if r.end.row > bottom {
        out.push(Range::new(CellRef::new(bottom + 1, r.start.col), r.end));
    }
    if r.start.col < cut.start.col {
        out.push(Range::new(CellRef::new(top, r.start.col), CellRef::new(bottom, cut.start.col - 1)));
    }
    if r.end.col > cut.end.col {
        out.push(Range::new(CellRef::new(top, cut.end.col + 1), CellRef::new(bottom, r.end.col)));
    }
    out
}

/// Parse Excel's space-separated range list (`sqref="A1:B5 D7"`).
pub fn parse_sqref(s: &str) -> Vec<Range> {
    s.split_whitespace()
        .filter_map(|part| match part.split_once(':') {
            Some((a, b)) => Some(Range::new(CellRef::parse(a)?, CellRef::parse(b)?)),
            None => CellRef::parse(part).map(Range::single),
        })
        .collect()
}

pub fn format_sqref(ranges: &[Range]) -> String {
    ranges.iter().map(|r| r.to_string()).collect::<Vec<_>>().join(" ")
}

/// Excel's "current region": the block of data around `at`, bounded by empty
/// rows and columns (Ctrl+A / Ctrl+Shift+* in Excel). A lone cell if nothing
/// around it holds data.
pub fn current_region(sheet: &Sheet, at: CellRef) -> Range {
    let filled =
        |r: u32, c: u32| !sheet.value(CellRef::new(r, c)).is_empty() || !sheet.input(CellRef::new(r, c)).is_empty();
    // Nothing lies beyond the used extent, so scans stop there.
    let (rows, cols) = sheet.extent();
    let (last_row, last_col) = (rows.max(at.row + 1).min(MAX_ROWS) - 1, cols.max(at.col + 1).min(MAX_COLS) - 1);
    let (mut top, mut left, mut bottom, mut right) = (at.row, at.col, at.row, at.col);
    // Each side grows as far as it can before the others are checked again, so
    // a 100k-row table takes a few passes rather than one pass per row.
    loop {
        let mut grew = false;
        let (lo_c, hi_c) = (left.saturating_sub(1), (right + 1).min(last_col));
        while top > 0 && (lo_c..=hi_c).any(|c| filled(top - 1, c)) {
            top -= 1;
            grew = true;
        }
        while bottom < last_row && (lo_c..=hi_c).any(|c| filled(bottom + 1, c)) {
            bottom += 1;
            grew = true;
        }
        let (lo_r, hi_r) = (top.saturating_sub(1), (bottom + 1).min(last_row));
        while left > 0 && (lo_r..=hi_r).any(|r| filled(r, left - 1)) {
            left -= 1;
            grew = true;
        }
        while right < last_col && (lo_r..=hi_r).any(|r| filled(r, right + 1)) {
            right += 1;
            grew = true;
        }
        if !grew {
            break;
        }
    }
    Range::new(CellRef::new(top, left), CellRef::new(bottom, right))
}

/// Excel wildcard match: `*` any run of characters, `?` one character, `~`
/// escapes the next one. Case-insensitive, works on characters (any script).
pub fn wildcard_match(pattern: &str, text: &str) -> bool {
    #[derive(Clone, Copy, PartialEq)]
    enum Tok {
        Any,
        One,
        Ch(char),
    }
    let mut toks = Vec::new();
    let mut chars = pattern.chars();
    while let Some(c) = chars.next() {
        toks.push(match c {
            '*' => Tok::Any,
            '?' => Tok::One,
            '~' => Tok::Ch(chars.next().unwrap_or('~')),
            c => Tok::Ch(c),
        });
    }
    let fold = |c: char| c.to_lowercase().next().unwrap_or(c);
    let text: Vec<char> = text.chars().map(fold).collect();
    // Classic two-pointer glob match with backtracking to the last `*`.
    let (mut t, mut p) = (0usize, 0usize);
    let (mut star, mut mark) = (None::<usize>, 0usize);
    while t < text.len() {
        match toks.get(p) {
            Some(Tok::One) => {
                t += 1;
                p += 1;
            }
            Some(Tok::Ch(c)) if fold(*c) == text[t] => {
                t += 1;
                p += 1;
            }
            Some(Tok::Any) => {
                star = Some(p);
                mark = t;
                p += 1;
            }
            _ => match star {
                Some(s) => {
                    p = s + 1;
                    mark += 1;
                    t = mark;
                }
                None => return false,
            },
        }
    }
    toks[p..].iter().all(|t| *t == Tok::Any)
}

pub fn has_wildcards(s: &str) -> bool {
    s.contains(['*', '?'])
}

/// A JSON string literal.
pub fn json_str(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::refs;

    fn r(s: &str) -> Range {
        parse_sqref(s)[0]
    }

    #[test]
    fn remaps_ranges() {
        let ins = refs::insert_map(2, 3);
        assert_eq!(remap_range(r("A1:B5"), Axis::Row, &ins), Some(r("A1:B8")));
        assert_eq!(remap_range(r("A4:B5"), Axis::Row, &ins), Some(r("A7:B8")));
        let del = refs::delete_map(0, 2);
        assert_eq!(remap_range(r("A1:B5"), Axis::Row, &del), Some(r("A1:B3")));
        assert_eq!(remap_range(r("A1:B2"), Axis::Row, &del), None);
        assert_eq!(remap_range(r("C1:D1"), Axis::Col, &del), Some(r("A1:B1")));
    }

    #[test]
    fn subtracts_rectangles() {
        let parts = subtract(r("A1:C3"), &r("B2"));
        let cells: usize = parts.iter().map(|p| p.cells().count()).sum();
        assert_eq!(cells, 8);
        assert!(parts.iter().all(|p| !contains(p, CellRef::new(1, 1))));
        assert_eq!(subtract(r("A1:A3"), &r("A1:A3")), Vec::<Range>::new());
        assert_eq!(subtract(r("A1"), &r("C3")), vec![r("A1")]);
    }

    #[test]
    fn sqref_round_trip() {
        let ranges = parse_sqref("A1:B5 D7  $E$1:$E$3");
        assert_eq!(format_sqref(&ranges), "A1:B5 D7 E1:E3");
    }

    #[test]
    fn finds_current_region() {
        let mut s = Sheet::new();
        for (a, v) in [("B2", "h1"), ("C2", "h2"), ("B3", "1"), ("C4", "2"), ("F9", "far")] {
            s.set_input(CellRef::parse(a).unwrap(), v);
        }
        assert_eq!(current_region(&s, CellRef::parse("B2").unwrap()), r("B2:C4"));
        assert_eq!(current_region(&s, CellRef::parse("C4").unwrap()), r("B2:C4"));
        assert_eq!(current_region(&s, CellRef::parse("A10").unwrap()), r("A10"));
    }

    #[test]
    fn wildcards() {
        assert!(wildcard_match("ab*", "ABC"));
        assert!(wildcard_match("*b*", "abc"));
        assert!(!wildcard_match("*z*", "abc"));
        assert!(wildcard_match("a?c", "abc"));
        assert!(!wildcard_match("a?c", "ac"));
        assert!(wildcard_match("*", ""));
        assert!(wildcard_match("a~*", "a*"));
        assert!(!wildcard_match("a~*", "ab"));
        assert!(wildcard_match("*सेब*", "लाल सेब"));
    }

    #[test]
    fn escapes_json() {
        assert_eq!(json_str("a\"b\\c\nd"), r#""a\"b\\c\nd""#);
    }
}
