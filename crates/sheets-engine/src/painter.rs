//! Format Painter: copy the formatting of a range and paint it onto another.
//!
//! Like Excel, a single target cell receives the whole copied block (the
//! painted area grows to the source size), and a larger target is tiled with
//! the source pattern.

use crate::cell::{CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::sheet::Sheet;
use crate::style::Style;

/// The formatting of a block of cells, captured when Format Painter is clicked.
#[derive(Debug, Clone, PartialEq)]
pub struct StyleSnapshot {
    rows: u32,
    cols: u32,
    /// Row-major.
    styles: Vec<Style>,
}

impl StyleSnapshot {
    pub fn capture(sheet: &Sheet, range: Range) -> StyleSnapshot {
        StyleSnapshot {
            rows: range.rows(),
            cols: range.cols(),
            styles: range.cells().map(|at| sheet.style(at).clone()).collect(),
        }
    }

    /// The area painted when the user selects `target`: the source size at a
    /// single cell, otherwise the target itself. Clamped to the sheet.
    pub fn paint_area(&self, target: Range) -> Range {
        let single = target.start == target.end;
        let (rows, cols) = if single { (self.rows, self.cols) } else { (target.rows(), target.cols()) };
        let end = CellRef::new(
            (target.start.row + rows - 1).min(MAX_ROWS - 1),
            (target.start.col + cols - 1).min(MAX_COLS - 1),
        );
        Range::new(target.start, end)
    }

    /// New styles for `area` (from `paint_area`), tiling the snapshot.
    pub fn styles_for(&self, area: Range) -> Vec<(CellRef, Style)> {
        area.cells()
            .map(|at| {
                let i = (at.row - area.start.row) % self.rows;
                let j = (at.col - area.start.col) % self.cols;
                (at, self.styles[(i * self.cols + j) as usize].clone())
            })
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    #[test]
    fn single_target_gets_the_source_size() {
        let mut s = Sheet::new();
        s.set_style(at("A1"), Style { bold: true, ..Style::default() });
        s.set_style(at("A2"), Style { italic: true, ..Style::default() });
        let snap = StyleSnapshot::capture(&s, Range::new(at("A1"), at("A2")));
        let area = snap.paint_area(Range::single(at("C5")));
        assert_eq!(area, Range::new(at("C5"), at("C6")));
        let styles = snap.styles_for(area);
        assert!(styles[0].1.bold && !styles[0].1.italic);
        assert!(styles[1].1.italic);
    }

    #[test]
    fn larger_target_is_tiled() {
        let mut s = Sheet::new();
        s.set_style(at("A1"), Style { fill: Some("#FFFF00".into()), ..Style::default() });
        // A1 yellow, B1 plain: a striped pattern.
        let snap = StyleSnapshot::capture(&s, Range::new(at("A1"), at("B1")));
        let area = snap.paint_area(Range::new(at("D1"), at("H2")));
        assert_eq!(area, Range::new(at("D1"), at("H2")));
        let styles = snap.styles_for(area);
        let yellow: Vec<String> =
            styles.iter().filter(|(_, st)| st.fill.is_some()).map(|(c, _)| c.to_string()).collect();
        assert_eq!(yellow, ["D1", "F1", "H1", "D2", "F2", "H2"]);
    }

    #[test]
    fn painting_plain_cells_clears_formatting() {
        let mut s = Sheet::new();
        s.set_style(at("B2"), Style { bold: true, ..Style::default() });
        let snap = StyleSnapshot::capture(&s, Range::single(at("A1")));
        let styles = snap.styles_for(snap.paint_area(Range::single(at("B2"))));
        assert_eq!(styles, vec![(at("B2"), Style::default())]);
    }

    #[test]
    fn clamped_at_the_sheet_edge() {
        let s = Sheet::new();
        let snap = StyleSnapshot::capture(&s, Range::new(at("A1"), at("C3")));
        let corner = CellRef::new(MAX_ROWS - 1, MAX_COLS - 1);
        assert_eq!(snap.paint_area(Range::single(corner)), Range::single(corner));
    }
}
