//! Cell addresses (`B12`) and rectangular ranges (`A1:C3`).

use std::fmt;

/// Excel's limits: 1,048,576 rows and 16,384 columns (A..XFD).
pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLS: u32 = 16_384;

/// Zero-based cell position. `B12` is `CellRef { row: 11, col: 1 }`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CellRef {
    pub row: u32,
    pub col: u32,
}

impl CellRef {
    pub const fn new(row: u32, col: u32) -> Self {
        CellRef { row, col }
    }

    /// Parse an A1-style reference. `$` markers (`$B$12`) are accepted and ignored.
    pub fn parse(s: &str) -> Option<Self> {
        if s.contains('$') {
            let plain: String = s.chars().filter(|&c| c != '$').collect();
            return Self::parse(&plain);
        }
        let split = s.find(|c: char| c.is_ascii_digit())?;
        let (letters, digits) = s.split_at(split);
        let col = parse_col(letters)?;
        if !digits.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
        let row: u32 = digits.parse().ok()?;
        if row == 0 || row > MAX_ROWS {
            return None;
        }
        Some(CellRef { row: row - 1, col })
    }
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}{}", col_name(self.col), self.row + 1)
    }
}

/// Column index to letters: 0 -> "A", 25 -> "Z", 26 -> "AA".
pub fn col_name(col: u32) -> String {
    let mut n = col + 1;
    let mut out = Vec::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap()
}

/// Column letters to index: "A" -> 0, "AA" -> 26. Case-insensitive.
pub fn parse_col(letters: &str) -> Option<u32> {
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let mut col: u32 = 0;
    for b in letters.bytes() {
        if !b.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (b.to_ascii_uppercase() - b'A' + 1) as u32;
    }
    (col <= MAX_COLS).then(|| col - 1)
}

/// Inclusive rectangle of cells, always stored with `start` top-left of `end`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Range {
    pub start: CellRef,
    pub end: CellRef,
}

impl Range {
    pub fn new(a: CellRef, b: CellRef) -> Self {
        Range {
            start: CellRef::new(a.row.min(b.row), a.col.min(b.col)),
            end: CellRef::new(a.row.max(b.row), a.col.max(b.col)),
        }
    }

    pub fn single(c: CellRef) -> Self {
        Range { start: c, end: c }
    }

    pub fn contains(&self, c: CellRef) -> bool {
        c.row >= self.start.row && c.row <= self.end.row && c.col >= self.start.col && c.col <= self.end.col
    }

    /// Number of cells, as u64 (a whole sheet is about 17 billion).
    pub fn area(&self) -> u64 {
        self.rows() as u64 * self.cols() as u64
    }

    pub fn rows(&self) -> u32 {
        self.end.row - self.start.row + 1
    }

    pub fn cols(&self) -> u32 {
        self.end.col - self.start.col + 1
    }

    /// Cells in row-major order.
    pub fn cells(&self) -> impl Iterator<Item = CellRef> + '_ {
        (self.start.row..=self.end.row)
            .flat_map(move |row| (self.start.col..=self.end.col).map(move |col| CellRef { row, col }))
    }
}

impl fmt::Display for Range {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.start == self.end {
            write!(f, "{}", self.start)
        } else {
            write!(f, "{}:{}", self.start, self.end)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_formats_refs() {
        assert_eq!(CellRef::parse("A1"), Some(CellRef::new(0, 0)));
        assert_eq!(CellRef::parse("$AB$10"), Some(CellRef::new(9, 27)));
        assert_eq!(CellRef::parse("xfd1048576"), Some(CellRef::new(MAX_ROWS - 1, MAX_COLS - 1)));
        assert_eq!(CellRef::parse("XFE1"), None);
        assert_eq!(CellRef::parse("A0"), None);
        assert_eq!(CellRef::parse("10"), None);
        assert_eq!(CellRef::parse("SUM"), None);
        assert_eq!(CellRef::new(9, 27).to_string(), "AB10");
    }

    #[test]
    fn column_names_round_trip() {
        for col in [0, 25, 26, 701, 702, MAX_COLS - 1] {
            assert_eq!(parse_col(&col_name(col)), Some(col));
        }
        assert_eq!(col_name(MAX_COLS - 1), "XFD");
    }

    #[test]
    fn range_normalizes() {
        let r = Range::new(CellRef::new(5, 3), CellRef::new(1, 0));
        assert_eq!(r.to_string(), "A2:D6");
        assert_eq!(r.cells().count(), 20);
    }
}
