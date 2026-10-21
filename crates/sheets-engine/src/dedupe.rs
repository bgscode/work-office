//! Remove Duplicates (Data → Remove Duplicates).
//!
//! Rows of a range are duplicates when the chosen columns show the same text
//! (compared case-insensitively, like Excel). The first occurrence is kept;
//! later ones are removed and the rows below move up *inside the range only*,
//! leaving the rest of the sheet where it is.

use std::collections::HashSet;

use crate::cell::{CellRef, Range};
use crate::sheet::Sheet;

/// Result of Remove Duplicates.
pub struct Deduped {
    pub sheet: Sheet,
    pub removed: u32,
    pub remaining: u32,
}

impl Sheet {
    /// `cols` are absolute column indices to compare; `has_header` keeps the first row out of it.
    pub fn with_duplicates_removed(&self, range: Range, cols: &[u32], has_header: bool) -> Deduped {
        let first = range.start.row + has_header as u32;
        if first > range.end.row || cols.is_empty() {
            let remaining = (range.end.row + 1).saturating_sub(first);
            return Deduped { sheet: self.clone(), removed: 0, remaining };
        }
        let mut seen: HashSet<Vec<String>> = HashSet::new();
        let kept: Vec<u32> = (first..=range.end.row)
            .filter(|r| seen.insert(cols.iter().map(|c| self.display(CellRef::new(*r, *c)).to_lowercase()).collect()))
            .collect();
        let total = range.end.row - first + 1;
        let removed = total - kept.len() as u32;
        let data = Range::new(CellRef::new(first, range.start.col), range.end);
        let sheet = if removed == 0 { self.clone() } else { self.with_rows_moved(data, &kept, total) };
        Deduped { sheet, removed, remaining: kept.len() as u32 }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn sheet() -> Sheet {
        let mut s = Sheet::new();
        let rows = [
            ["Name", "City"],
            ["Asha", "Delhi"],
            ["asha", "DELHI"],
            ["Ravi", "Delhi"],
            ["Asha", "Pune"],
            ["Ravi", "Delhi"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                s.set_input(CellRef::new(r as u32, c as u32), v);
            }
        }
        s.set_input(at("A8"), "below the table");
        s
    }

    #[test]
    fn removes_duplicate_rows() {
        let s = sheet();
        let d = s.with_duplicates_removed(Range::new(at("A1"), at("B6")), &[0, 1], true);
        assert_eq!((d.removed, d.remaining), (2, 3));
        let names: Vec<&str> = (0..6).map(|r| d.sheet.input(CellRef::new(r, 0))).collect();
        assert_eq!(names, ["Name", "Asha", "Ravi", "Asha", "", ""]);
        assert_eq!(d.sheet.input(at("B4")), "Pune");
        assert_eq!(d.sheet.input(at("A8")), "below the table");

        let by_name = s.with_duplicates_removed(Range::new(at("A1"), at("B6")), &[0], true);
        assert_eq!((by_name.removed, by_name.remaining), (3, 2));
        let none = s.with_duplicates_removed(Range::new(at("A1"), at("B2")), &[0, 1], true);
        assert_eq!((none.removed, none.remaining), (0, 1));
        // Without a header row, "Name" is data too.
        let all = s.with_duplicates_removed(Range::new(at("A1"), at("B6")), &[1], false);
        assert_eq!((all.removed, all.remaining), (3, 3));
    }
}
