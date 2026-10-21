//! Rewrite the cell references inside formula text, leaving everything else
//! (strings, function names, spacing) untouched. Used when formulas move:
//!
//! - copy/paste and the fill handle shift relative references (`A1` -> `A2`)
//!   but keep absolute ones (`$A$1`);
//! - inserting/deleting rows or columns moves every reference that points
//!   past the change, and turns references to deleted cells into `#REF!`.

use crate::cell::{col_name, parse_col, CellRef, MAX_COLS, MAX_ROWS};

/// A reference as written, including its `$` markers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RefText {
    pub cell: CellRef,
    pub abs_col: bool,
    pub abs_row: bool,
}

impl RefText {
    pub fn parse(s: &str) -> Option<RefText> {
        let b = s.as_bytes();
        let mut i = 0;
        let abs_col = b.first() == Some(&b'$');
        if abs_col {
            i += 1;
        }
        let col_start = i;
        while i < b.len() && b[i].is_ascii_alphabetic() {
            i += 1;
        }
        let col = parse_col(&s[col_start..i])?;
        let abs_row = b.get(i) == Some(&b'$');
        if abs_row {
            i += 1;
        }
        let digits = &s[i..];
        if digits.is_empty() || !digits.bytes().all(|c| c.is_ascii_digit()) {
            return None;
        }
        let row: u32 = digits.parse().ok()?;
        if row == 0 || row > MAX_ROWS {
            return None;
        }
        Some(RefText { cell: CellRef::new(row - 1, col), abs_col, abs_row })
    }
}

impl std::fmt::Display for RefText {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let dollar = |abs: bool| if abs { "$" } else { "" };
        write!(f, "{}{}{}{}", dollar(self.abs_col), col_name(self.cell.col), dollar(self.abs_row), self.cell.row + 1)
    }
}

/// A single reference or a range, as found in the text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RefSpan {
    Cell(RefText),
    Range(RefText, RefText),
}

/// Call `f` for every reference in `formula` (text after `=`) and splice in
/// what it returns; `None` becomes `#REF!`.
pub fn rewrite(formula: &str, mut f: impl FnMut(RefSpan) -> Option<RefSpan>) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    let is_ident_start = |c: char| c.is_alphabetic() || c == '_' || c == '$';
    let is_ident = |c: char| c.is_alphanumeric() || matches!(c, '_' | '$' | '.');
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            // Copy a string literal verbatim, including "" escapes.
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '"' {
                    if chars.get(i + 1) == Some(&'"') {
                        out.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if !is_ident_start(c) || (i > 0 && is_ident(chars[i - 1])) {
            out.push(c);
            i += 1;
            continue;
        }
        let start = i;
        while i < chars.len() && is_ident(chars[i]) {
            i += 1;
        }
        let word: String = chars[start..i].iter().collect();
        let next_non_space = chars[i..].iter().find(|c| !c.is_whitespace());
        let Some(first) = RefText::parse(&word).filter(|_| next_non_space != Some(&'(')) else {
            out.push_str(&word);
            continue;
        };
        // Range: A1:B2
        let mut span = RefSpan::Cell(first);
        if chars.get(i) == Some(&':') {
            let mut j = i + 1;
            while j < chars.len() && is_ident(chars[j]) {
                j += 1;
            }
            let second: String = chars[i + 1..j].iter().collect();
            if let Some(end) = RefText::parse(&second) {
                span = RefSpan::Range(first, end);
                i = j;
            }
        }
        match f(span) {
            Some(RefSpan::Cell(r)) => out.push_str(&r.to_string()),
            Some(RefSpan::Range(a, b)) => out.push_str(&format!("{a}:{b}")),
            None => out.push_str("#REF!"),
        }
    }
    out
}

/// Shift relative references by (d_row, d_col), as when copying a formula.
/// `input` is a full cell input; non-formulas are returned unchanged.
pub fn shift_input(input: &str, d_row: i64, d_col: i64) -> String {
    let Some(formula) = input.strip_prefix('=') else { return input.to_string() };
    let shift = |r: RefText| -> Option<RefText> {
        let row = if r.abs_row { r.cell.row as i64 } else { r.cell.row as i64 + d_row };
        let col = if r.abs_col { r.cell.col as i64 } else { r.cell.col as i64 + d_col };
        if row < 0 || col < 0 || row >= MAX_ROWS as i64 || col >= MAX_COLS as i64 {
            return None;
        }
        Some(RefText { cell: CellRef::new(row as u32, col as u32), ..r })
    };
    let body = rewrite(formula, |span| match span {
        RefSpan::Cell(r) => shift(r).map(RefSpan::Cell),
        RefSpan::Range(a, b) => Some(RefSpan::Range(shift(a)?, shift(b)?)),
    });
    format!("={body}")
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Axis {
    Row,
    Col,
}

/// Rewrite references after rows or columns move. `map` takes an old index on
/// `axis` and returns its new index, or `None` if it was deleted. Ranges shrink
/// to their surviving part and become `#REF!` only when nothing survives.
pub fn remap_input(input: &str, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> String {
    let Some(formula) = input.strip_prefix('=') else { return input.to_string() };
    let get = |r: &RefText| if axis == Axis::Row { r.cell.row } else { r.cell.col };
    let set = |r: RefText, v: u32| {
        let cell = if axis == Axis::Row { CellRef::new(v, r.cell.col) } else { CellRef::new(r.cell.row, v) };
        RefText { cell, ..r }
    };
    let body = rewrite(formula, |span| match span {
        RefSpan::Cell(r) => map(get(&r)).map(|v| RefSpan::Cell(set(r, v))),
        RefSpan::Range(a, b) => {
            let (lo, hi) = (get(&a).min(get(&b)), get(&a).max(get(&b)));
            let new_lo = (lo..=hi).find_map(map)?;
            let new_hi = (lo..=hi).rev().find_map(map)?;
            Some(RefSpan::Range(set(a, new_lo), set(b, new_hi)))
        }
    });
    format!("={body}")
}

/// Index mapping for inserting `count` rows/columns before `at`.
pub fn insert_map(at: u32, count: u32) -> impl Fn(u32) -> Option<u32> {
    move |i| if i >= at { i.checked_add(count) } else { Some(i) }
}

/// Index mapping for deleting `count` rows/columns starting at `at`.
pub fn delete_map(at: u32, count: u32) -> impl Fn(u32) -> Option<u32> {
    move |i| {
        if i < at {
            Some(i)
        } else if i >= at + count {
            Some(i - count)
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shifts_relative_refs_only() {
        assert_eq!(shift_input("=A1+$B$2+C$3+$D4", 1, 1), "=B2+$B$2+D$3+$D5");
        assert_eq!(shift_input("=SUM(A1:A3)*2", 2, 0), "=SUM(A3:A5)*2");
        assert_eq!(shift_input("=A1", -1, 0), "=#REF!");
        assert_eq!(shift_input("plain A1", 1, 0), "plain A1");
    }

    #[test]
    fn leaves_strings_and_functions_alone() {
        assert_eq!(shift_input(r#"=CONCAT("A1",A1)&LOG10(A1)"#, 1, 0), r#"=CONCAT("A1",A2)&LOG10(A2)"#);
        assert_eq!(shift_input(r#"="say ""A1""" & B1"#, 0, 1), r#"="say ""A1""" & C1"#);
    }

    #[test]
    fn remaps_for_insert_and_delete() {
        let ins = insert_map(2, 2); // insert 2 rows before row 3
        assert_eq!(remap_input("=A1+A3+SUM(A1:A5)", Axis::Row, &ins), "=A1+A5+SUM(A1:A7)");
        let del = delete_map(2, 2); // delete rows 3-4
        assert_eq!(remap_input("=A3+A5+SUM(A1:A10)", Axis::Row, &del), "=#REF!+A3+SUM(A1:A8)");
        assert_eq!(remap_input("=SUM(A3:A4)", Axis::Row, &del), "=SUM(#REF!)");
        assert_eq!(remap_input("=SUM(A3:A6)", Axis::Row, &del), "=SUM(A3:A4)");
        let del_col = delete_map(0, 1);
        assert_eq!(remap_input("=$B$1*C2", Axis::Col, &del_col), "=$A$1*B2");
    }
}
