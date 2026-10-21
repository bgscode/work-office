//! Rewrite the cell references inside formula text, leaving everything else
//! (strings, function names, spacing) untouched. Used when formulas move:
//!
//! - copy/paste and the fill handle shift relative references (`A1` -> `A2`)
//!   but keep absolute ones (`$A$1`);
//! - inserting/deleting rows or columns moves every reference that points
//!   past the change, and turns references to deleted cells into `#REF!`;
//! - renaming a sheet rewrites `Old!A1` to `New!A1`, deleting it turns
//!   `Old!A1` into `#REF!`.
//!
//! References may carry a sheet prefix (`Sheet2!A1`, `'My Sheet'!A1:B5`) and
//! may be whole columns (`A:C`) or whole rows (`1:3`).

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

/// A whole column (`$C` in `A:$C`) or whole row (`$3` in `1:$3`) as written.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LineText {
    /// Zero-based column or row.
    pub index: u32,
    pub abs: bool,
}

impl LineText {
    pub fn parse_col(s: &str) -> Option<LineText> {
        let (abs, letters) = match s.strip_prefix('$') {
            Some(rest) => (true, rest),
            None => (false, s),
        };
        if letters.is_empty() || !letters.bytes().all(|b| b.is_ascii_alphabetic()) {
            return None;
        }
        Some(LineText { index: parse_col(letters)?, abs })
    }

    pub fn parse_row(s: &str) -> Option<LineText> {
        let (abs, digits) = match s.strip_prefix('$') {
            Some(rest) => (true, rest),
            None => (false, s),
        };
        if digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
        let n: u32 = digits.parse().ok()?;
        (1..=MAX_ROWS).contains(&n).then(|| LineText { index: n - 1, abs })
    }

    fn col_text(self) -> String {
        format!("{}{}", if self.abs { "$" } else { "" }, col_name(self.index))
    }

    fn row_text(self) -> String {
        format!("{}{}", if self.abs { "$" } else { "" }, self.index + 1)
    }
}

/// A single reference or a range, as found in the text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RefSpan {
    Cell(RefText),
    Range(RefText, RefText),
    /// Whole columns: `A:C`.
    Cols(LineText, LineText),
    /// Whole rows: `1:3`.
    Rows(LineText, LineText),
}

impl std::fmt::Display for RefSpan {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            RefSpan::Cell(r) => write!(f, "{r}"),
            RefSpan::Range(a, b) => write!(f, "{a}:{b}"),
            RefSpan::Cols(a, b) => write!(f, "{}:{}", a.col_text(), b.col_text()),
            RefSpan::Rows(a, b) => write!(f, "{}:{}", a.row_text(), b.row_text()),
        }
    }
}

/// A sheet name as it must be written in a formula: `Sheet2`, or quoted
/// `'My Sheet'` when Excel requires it (spaces or punctuation, a leading
/// digit, or a name that looks like a cell reference). `'` is doubled.
pub fn quote_sheet_name(name: &str) -> String {
    // Letters and digits of any script are fine unquoted, including combining
    // marks such as Devanagari vowel signs (बिक्री); ASCII punctuation and spaces need quotes.
    let plain =
        name.chars().all(|c| c.is_alphanumeric() || c == '_' || c == '.' || (!c.is_ascii() && !c.is_whitespace()))
            && name.chars().next().is_some_and(|c| c.is_alphabetic() || c == '_')
            && CellRef::parse(name).is_none()
            && !looks_like_r1c1(name)
            && !name.eq_ignore_ascii_case("TRUE")
            && !name.eq_ignore_ascii_case("FALSE");
    if plain {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

/// `R`, `C`, `R1`, `C2`, `R1C1`... (R1C1-style addresses need quotes too).
fn looks_like_r1c1(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    let digits = |s: &str| s.trim_start_matches(|c: char| c.is_ascii_digit()).len();
    if let Some(rest) = upper.strip_prefix('R') {
        let after = &rest[rest.len() - digits(rest)..];
        return after.is_empty() || after.strip_prefix('C').is_some_and(|c| digits(c) == 0);
    }
    upper.strip_prefix('C').is_some_and(|c| digits(c) == 0)
}

/// A reference found in formula text.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FoundRef {
    /// Character offsets in the formula; `start` includes the sheet prefix.
    pub start: usize,
    pub end: usize,
    /// The sheet prefix, unquoted, if any.
    pub sheet: Option<String>,
    pub span: RefSpan,
}

/// Every reference in `formula` (text after `=`), skipping strings and function names.
pub fn find_refs(formula: &str) -> Vec<FoundRef> {
    let chars: Vec<char> = formula.chars().collect();
    let is_ident_start = |c: char| c.is_alphabetic() || c == '_' || c == '$';
    let is_ident = |c: char| c.is_alphanumeric() || matches!(c, '_' | '$' | '.');
    let word_end = |from: usize| {
        let mut j = from;
        while j < chars.len() && is_ident(chars[j]) {
            j += 1;
        }
        j
    };
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            // Skip a string literal, including "" escapes.
            i += 1;
            while i < chars.len() {
                if chars[i] == '"' {
                    if chars.get(i + 1) == Some(&'"') {
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
        let after_ident = i > 0 && is_ident(chars[i - 1]);
        let start = i;
        let mut sheet = None;
        if c == '\'' && !after_ident {
            // 'Quoted name'!
            let mut name = String::new();
            let mut j = i + 1;
            let mut closed = false;
            while j < chars.len() {
                if chars[j] == '\'' {
                    if chars.get(j + 1) == Some(&'\'') {
                        name.push('\'');
                        j += 2;
                        continue;
                    }
                    closed = true;
                    j += 1;
                    break;
                }
                name.push(chars[j]);
                j += 1;
            }
            if !closed || chars.get(j) != Some(&'!') {
                i = j;
                continue;
            }
            sheet = Some(name);
            i = j + 1;
        } else if (is_ident_start(c) || c.is_ascii_digit()) && !after_ident {
            let j = word_end(i);
            if chars.get(j) == Some(&'!') && !c.is_ascii_digit() {
                sheet = Some(chars[i..j].iter().collect::<String>());
                i = j + 1;
            }
        } else {
            i += 1;
            continue;
        }
        // The reference itself.
        let j = word_end(i);
        let word: String = chars[i..j].iter().collect();
        let is_call = sheet.is_none() && chars[j..].iter().find(|c| !c.is_whitespace()) == Some(&'(');
        let second = |parse: &dyn Fn(&str) -> Option<LineOrRef>| -> Option<(LineOrRef, usize)> {
            if chars.get(j) != Some(&':') {
                return None;
            }
            let k = word_end(j + 1);
            let w: String = chars[j + 1..k].iter().collect();
            Some((parse(&w)?, k))
        };
        let mut found: Option<(RefSpan, usize)> = None;
        if !is_call && !word.is_empty() {
            if let Some(a) = RefText::parse(&word) {
                found = Some(match second(&|w| RefText::parse(w).map(LineOrRef::Ref)) {
                    Some((LineOrRef::Ref(b), k)) => (RefSpan::Range(a, b), k),
                    _ => (RefSpan::Cell(a), j),
                });
            } else if let Some(a) = LineText::parse_col(&word) {
                if let Some((LineOrRef::Line(b), k)) = second(&|w| LineText::parse_col(w).map(LineOrRef::Line)) {
                    found = Some((RefSpan::Cols(a, b), k));
                }
            } else if let Some(a) = LineText::parse_row(&word) {
                if let Some((LineOrRef::Line(b), k)) = second(&|w| LineText::parse_row(w).map(LineOrRef::Line)) {
                    found = Some((RefSpan::Rows(a, b), k));
                }
            }
        }
        match found {
            Some((span, k)) => {
                out.push(FoundRef { start, end: k, sheet, span });
                i = k;
            }
            None => i = j.max(i + 1),
        }
    }
    out
}

enum LineOrRef {
    Ref(RefText),
    Line(LineText),
}

/// Call `f` for every reference in `formula` (text after `=`) with its sheet
/// prefix (unquoted) and splice in what it returns: the sheet prefix to write
/// (`None` for none; quoted as needed) and the reference, or `None` for `#REF!`.
pub fn rewrite_refs(
    formula: &str,
    mut f: impl FnMut(Option<&str>, RefSpan) -> Option<(Option<String>, RefSpan)>,
) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut pos = 0;
    for r in find_refs(formula) {
        out.extend(&chars[pos..r.start]);
        match f(r.sheet.as_deref(), r.span) {
            // Unchanged references keep their text exactly as typed.
            Some((sheet, span)) if sheet == r.sheet && span == r.span => out.extend(&chars[r.start..r.end]),
            Some((Some(sheet), span)) => out.push_str(&format!("{}!{span}", quote_sheet_name(&sheet))),
            Some((None, span)) => out.push_str(&span.to_string()),
            None => out.push_str("#REF!"),
        }
        pos = r.end;
    }
    out.extend(&chars[pos..]);
    out
}

/// Call `f` for every reference in `formula` (text after `=`) and splice in
/// what it returns, keeping any sheet prefix; `None` becomes `#REF!`.
pub fn rewrite(formula: &str, mut f: impl FnMut(RefSpan) -> Option<RefSpan>) -> String {
    rewrite_refs(formula, |sheet, span| Some((sheet.map(str::to_string), f(span)?)))
}

/// Apply `f` to the formula part of a cell input; other inputs are returned unchanged.
fn with_formula(input: &str, f: impl FnOnce(&str) -> String) -> String {
    match input.strip_prefix('=') {
        Some(formula) => format!("={}", f(formula)),
        None => input.to_string(),
    }
}

/// Shift relative references by (d_row, d_col), as when copying a formula.
/// `input` is a full cell input; non-formulas are returned unchanged.
pub fn shift_input(input: &str, d_row: i64, d_col: i64) -> String {
    let shift_by = |v: u32, abs: bool, d: i64, limit: u32| -> Option<u32> {
        let n = if abs { v as i64 } else { v as i64 + d };
        (n >= 0 && n < limit as i64).then_some(n as u32)
    };
    let shift = |r: RefText| -> Option<RefText> {
        let row = shift_by(r.cell.row, r.abs_row, d_row, MAX_ROWS)?;
        let col = shift_by(r.cell.col, r.abs_col, d_col, MAX_COLS)?;
        Some(RefText { cell: CellRef::new(row, col), ..r })
    };
    let line = |l: LineText, d: i64, limit: u32| Some(LineText { index: shift_by(l.index, l.abs, d, limit)?, ..l });
    with_formula(input, |formula| {
        rewrite(formula, |span| match span {
            RefSpan::Cell(r) => shift(r).map(RefSpan::Cell),
            RefSpan::Range(a, b) => Some(RefSpan::Range(shift(a)?, shift(b)?)),
            RefSpan::Cols(a, b) => Some(RefSpan::Cols(line(a, d_col, MAX_COLS)?, line(b, d_col, MAX_COLS)?)),
            RefSpan::Rows(a, b) => Some(RefSpan::Rows(line(a, d_row, MAX_ROWS)?, line(b, d_row, MAX_ROWS)?)),
        })
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Axis {
    Row,
    Col,
}

/// Remap one reference after rows/columns move (see [`remap_input`]).
fn remap_span(span: RefSpan, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Option<RefSpan> {
    let limit = if axis == Axis::Row { MAX_ROWS } else { MAX_COLS };
    let map = |i: u32| map(i).filter(|n| *n < limit);
    // A span lo..=hi keeps its surviving part; it is #REF! only when nothing survives.
    let span_map = |lo: u32, hi: u32| -> Option<(u32, u32)> {
        let new_lo = (lo..=hi).find_map(map)?;
        // Cells pushed past the sheet edge are dropped from the end of the range.
        let new_hi = (lo..=hi).rev().find_map(map)?;
        Some((new_lo, new_hi))
    };
    let get = |r: &RefText| if axis == Axis::Row { r.cell.row } else { r.cell.col };
    let set = |r: RefText, v: u32| {
        let cell = if axis == Axis::Row { CellRef::new(v, r.cell.col) } else { CellRef::new(r.cell.row, v) };
        RefText { cell, ..r }
    };
    match span {
        RefSpan::Cell(r) => map(get(&r)).map(|v| RefSpan::Cell(set(r, v))),
        RefSpan::Range(a, b) => {
            let (lo, hi) = span_map(get(&a).min(get(&b)), get(&a).max(get(&b)))?;
            let (a_lo, b_lo) = (get(&a) <= get(&b), get(&b) < get(&a));
            Some(RefSpan::Range(set(a, if a_lo { lo } else { hi }), set(b, if b_lo { lo } else { hi })))
        }
        RefSpan::Cols(a, b) if axis == Axis::Col => {
            let (lo, hi) = span_map(a.index.min(b.index), a.index.max(b.index))?;
            Some(RefSpan::Cols(LineText { index: lo, ..a }, LineText { index: hi, ..b }))
        }
        RefSpan::Rows(a, b) if axis == Axis::Row => {
            let (lo, hi) = span_map(a.index.min(b.index), a.index.max(b.index))?;
            Some(RefSpan::Rows(LineText { index: lo, ..a }, LineText { index: hi, ..b }))
        }
        // Whole columns don't move when rows change, and vice versa.
        other => Some(other),
    }
}

/// Rewrite the references *without* a sheet prefix after rows or columns of
/// the formula's own sheet move. `map` takes an old index on `axis` and
/// returns its new index, or `None` if it was deleted. Ranges shrink to their
/// surviving part and become `#REF!` only when nothing survives.
pub fn remap_input(input: &str, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> String {
    with_formula(input, |formula| {
        rewrite_refs(formula, |sheet, span| match sheet {
            None => Some((None, remap_span(span, axis, map)?)),
            Some(s) => Some((Some(s.to_string()), span)),
        })
    })
}

/// Like [`remap_input`], but for references *with* a prefix naming `sheet`
/// (any case): used for formulas anywhere in the workbook after rows or
/// columns of `sheet` move.
pub fn remap_sheet_input(input: &str, sheet: &str, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> String {
    with_formula(input, |formula| {
        rewrite_refs(formula, |prefix, span| match prefix {
            Some(s) if s.eq_ignore_ascii_case(sheet) || s.to_lowercase() == sheet.to_lowercase() => {
                Some((Some(s.to_string()), remap_span(span, axis, map)?))
            }
            other => Some((other.map(str::to_string), span)),
        })
    })
}

/// Rewrite `Old!A1` to `New!A1` (quoting `new` if needed) after a sheet rename.
pub fn rename_sheet_input(input: &str, old: &str, new: &str) -> String {
    with_formula(input, |formula| {
        rewrite_refs(formula, |prefix, span| match prefix {
            Some(s) if same_sheet(s, old) => Some((Some(new.to_string()), span)),
            other => Some((other.map(str::to_string), span)),
        })
    })
}

/// Turn every reference to `sheet` into `#REF!`, as when the sheet is deleted.
pub fn delete_sheet_input(input: &str, sheet: &str) -> String {
    with_formula(input, |formula| {
        rewrite_refs(formula, |prefix, span| match prefix {
            Some(s) if same_sheet(s, sheet) => None,
            other => Some((other.map(str::to_string), span)),
        })
    })
}

/// Sheet names compare case-insensitively, as in Excel.
pub fn same_sheet(a: &str, b: &str) -> bool {
    a.eq_ignore_ascii_case(b) || a.to_lowercase() == b.to_lowercase()
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
        assert_eq!(shift_input("=1.5E+3+A1", 1, 0), "=1.5E+3+A2");
    }

    #[test]
    fn shifts_sheet_refs_and_whole_lines() {
        assert_eq!(shift_input("=Sheet2!A1+'My Sheet'!$B$1", 1, 0), "=Sheet2!A2+'My Sheet'!$B$1");
        assert_eq!(shift_input("=SUM(A:A)+SUM($B:C)", 5, 1), "=SUM(B:B)+SUM($B:D)");
        assert_eq!(shift_input("=SUM(1:1)+SUM($2:3)", 1, 5), "=SUM(2:2)+SUM($2:4)");
        assert_eq!(shift_input("=SUM(A:A)", 0, -1), "=SUM(#REF!)");
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

    #[test]
    fn remaps_whole_lines_and_leaves_other_sheets() {
        let ins = insert_map(0, 1);
        assert_eq!(remap_input("=SUM(A:A)+SUM(2:3)+Data!A1", Axis::Row, &ins), "=SUM(A:A)+SUM(3:4)+Data!A1");
        assert_eq!(remap_input("=SUM(A:B)", Axis::Col, &ins), "=SUM(B:C)");
        // A range reaching the last row keeps ending there.
        assert_eq!(remap_input("=SUM(A1:A1048576)", Axis::Row, &ins), "=SUM(A2:A1048576)");
        assert_eq!(remap_input("=A1048576", Axis::Row, &ins), "=#REF!");
        assert_eq!(remap_sheet_input("=data!A5+A5+'Other'!A5", "Data", Axis::Row, &ins), "=data!A6+A5+'Other'!A5");
        assert_eq!(remap_sheet_input("='My Sheet'!B2:B4", "my sheet", Axis::Row, &delete_map(1, 3)), "=#REF!");
    }

    #[test]
    fn renames_and_deletes_sheets() {
        assert_eq!(
            rename_sheet_input("=Sheet2!A1+sheet2!B:B+A1", "Sheet2", "My Data"),
            "='My Data'!A1+'My Data'!B:B+A1"
        );
        // "Q1" is also a cell address, so Excel quotes it.
        assert_eq!(rename_sheet_input("='My Data'!A1", "my data", "Q1"), "='Q1'!A1");
        assert_eq!(rename_sheet_input("='My Data'!A1", "my data", "Sales"), "=Sales!A1");
        assert_eq!(rename_sheet_input(r#"="Sheet2!A1""#, "Sheet2", "X"), r#"="Sheet2!A1""#);
        assert_eq!(delete_sheet_input("=Sheet2!A1+Sheet3!A1", "Sheet2"), "=#REF!+Sheet3!A1");
    }

    #[test]
    fn finds_refs_with_prefixes() {
        let refs = find_refs("SUM('It''s'!A1:B2, Data!C:C, 3:3, x1, LOG10(2))");
        let spans: Vec<_> = refs.iter().map(|r| (r.sheet.clone(), r.span.to_string())).collect();
        assert_eq!(
            spans,
            vec![
                (Some("It's".to_string()), "A1:B2".to_string()),
                (Some("Data".to_string()), "C:C".to_string()),
                (None, "3:3".to_string()),
                (None, "X1".to_string()),
            ]
        );
        assert_eq!((refs[0].start, refs[0].end), (4, 17));
    }

    #[test]
    fn quotes_sheet_names_like_excel() {
        assert_eq!(quote_sheet_name("Sheet2"), "Sheet2");
        assert_eq!(quote_sheet_name("Q1_2025.v2"), "Q1_2025.v2");
        assert_eq!(quote_sheet_name("My Sheet"), "'My Sheet'");
        assert_eq!(quote_sheet_name("2024"), "'2024'");
        assert_eq!(quote_sheet_name("A1"), "'A1'");
        assert_eq!(quote_sheet_name("R1C1"), "'R1C1'");
        assert_eq!(quote_sheet_name("R"), "'R'");
        assert_eq!(quote_sheet_name("Rates"), "Rates");
        assert_eq!(quote_sheet_name("Bob's"), "'Bob''s'");
        assert_eq!(quote_sheet_name("बिक्री"), "बिक्री");
        assert_eq!(quote_sheet_name("a-b"), "'a-b'");
    }
}
