//! Multi-level sort (Data → Sort): sort the rows of a range by several
//! columns, each ascending or descending. Empty cells always go last, ties
//! keep their order (stable), and formulas move with their rows with their
//! relative references shifted, as `Sheet::with_sorted_rows` does for the
//! one-column sort buttons.

use std::cmp::Ordering;

use crate::cell::{CellRef, Range};
use crate::formula::refs;
use crate::sheet::Sheet;
use crate::value::compare;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SortKey {
    pub col: u32,
    pub ascending: bool,
}

impl Sheet {
    /// A copy with the rows of `range` (data only, no header) sorted by `keys`, first key first.
    pub fn with_sorted_rows_by(&self, range: Range, keys: &[SortKey]) -> Sheet {
        let mut rows: Vec<u32> = (range.start.row..=range.end.row).collect();
        rows.sort_by(|a, b| {
            for key in keys {
                let (va, vb) = (self.value(CellRef::new(*a, key.col)), self.value(CellRef::new(*b, key.col)));
                let ord = match (va.is_empty(), vb.is_empty()) {
                    (true, true) => Ordering::Equal,
                    (true, false) => Ordering::Greater,
                    (false, true) => Ordering::Less,
                    _ if key.ascending => compare(va, vb),
                    _ => compare(vb, va),
                };
                if ord != Ordering::Equal {
                    return ord;
                }
            }
            Ordering::Equal
        });
        self.with_rows_moved(range, &rows, range.rows())
    }

    /// A copy where row `start + i` of `range` takes the content and formatting
    /// of row `order[i]`; rows of `range` past `order.len()` are emptied.
    /// `span` is how many rows of `range` to rewrite.
    pub(crate) fn with_rows_moved(&self, range: Range, order: &[u32], span: u32) -> Sheet {
        let mut out = self.clone();
        let mut edits = Vec::new();
        for i in 0..span {
            let to_row = range.start.row + i;
            let from_row = order.get(i as usize).copied();
            if from_row == Some(to_row) {
                continue;
            }
            for col in range.start.col..=range.end.col {
                let to = CellRef::new(to_row, col);
                match from_row {
                    Some(r) => {
                        let from = CellRef::new(r, col);
                        out.set_style(to, self.style(from).clone());
                        edits.push((to, refs::shift_input(&self.input(from), to_row as i64 - r as i64, 0)));
                    }
                    None => {
                        out.set_style(to, crate::Style::default());
                        if !self.input(to).is_empty() {
                            edits.push((to, String::new()));
                        }
                    }
                }
            }
        }
        out.set_inputs(edits);
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Style, Value};

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    #[test]
    fn sorts_by_several_columns() {
        let mut s = Sheet::new();
        let data = [["Delhi", "3", "a"], ["Agra", "5", "b"], ["Delhi", "1", "c"], ["", "9", "d"], ["Agra", "5", "e"]];
        for (r, row) in data.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                s.set_input(CellRef::new(r as u32, c as u32), v);
            }
        }
        s.set_input(at("D1"), "=B1*2");
        s.set_style(at("C3"), Style { bold: true, ..Style::default() });
        let keys = [SortKey { col: 0, ascending: true }, SortKey { col: 1, ascending: false }];
        let sorted = s.with_sorted_rows_by(Range::new(at("A1"), at("D5")), &keys);
        let col = |c: u32| (0..5).map(|r| sorted.input(CellRef::new(r, c)).to_string()).collect::<Vec<_>>();
        assert_eq!(col(0), ["Agra", "Agra", "Delhi", "Delhi", ""]);
        // Equal keys keep their order (b before e).
        assert_eq!(col(2), ["b", "e", "a", "c", "d"]);
        assert_eq!(sorted.input(at("D3")), "=B3*2");
        assert_eq!(sorted.value(at("D3")), &Value::Number(6.0));
        assert!(sorted.style(at("C4")).bold);
        assert!(!sorted.style(at("C3")).bold);
    }
}
