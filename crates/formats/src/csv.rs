//! CSV import/export (RFC 4180: quoted fields, `""` escapes, CRLF or LF).
//!
//! The delimiter is configurable (comma, semicolon as used where the comma is
//! the decimal separator, tab, pipe); [`detect_delimiter`] guesses it the way
//! Excel's text import does. Byte-level encoding (UTF-8 BOM, legacy code pages)
//! is handled by the caller, which passes and receives Rust strings.

use sheets_engine::{CellRef, Sheet};

/// Split CSV text into rows of fields.
pub fn parse(text: &str) -> Vec<Vec<String>> {
    parse_with(text, ',')
}

/// Split delimited text into rows of fields, with `delimiter` between fields.
pub fn parse_with(text: &str, delimiter: char) -> Vec<Vec<String>> {
    let text = text.strip_prefix('\u{feff}').unwrap_or(text);
    let mut rows = Vec::new();
    let mut row = Vec::new();
    let mut field = String::new();
    let mut in_quotes = false;
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        match (c, in_quotes) {
            ('"', true) if chars.peek() == Some(&'"') => {
                field.push('"');
                chars.next();
            }
            ('"', true) => in_quotes = false,
            ('"', false) if field.is_empty() => in_quotes = true,
            (d, false) if d == delimiter => row.push(std::mem::take(&mut field)),
            ('\r', false) if chars.peek() == Some(&'\n') => {}
            ('\n', false) | ('\r', false) => {
                row.push(std::mem::take(&mut field));
                rows.push(std::mem::take(&mut row));
            }
            (c, _) => field.push(c),
        }
    }
    if !field.is_empty() || !row.is_empty() {
        row.push(field);
        rows.push(row);
    }
    rows
}

/// Build a sheet from CSV. Fields are treated like typed input, so numbers
/// become numbers and `=...` becomes a formula, as in Excel.
pub fn read(text: &str) -> Sheet {
    read_with(text, ',')
}

/// [`read`] with another delimiter.
pub fn read_with(text: &str, delimiter: char) -> Sheet {
    let mut sheet = Sheet::new();
    let edits = parse_with(text, delimiter).into_iter().enumerate().flat_map(|(r, fields)| {
        fields
            .into_iter()
            .enumerate()
            .filter(|(_, f)| !f.is_empty())
            .map(move |(c, f)| (CellRef::new(r as u32, c as u32), f))
    });
    sheet.set_inputs(edits.collect::<Vec<_>>());
    sheet
}

/// Export the displayed values (not formulas) of a sheet, formatted as shown.
pub fn write(sheet: &Sheet) -> String {
    write_with(sheet, ',')
}

/// [`write`] with another delimiter.
pub fn write_with(sheet: &Sheet, delimiter: char) -> String {
    let (rows, cols) = sheet.extent();
    let mut out = String::new();
    let sep = delimiter.to_string();
    for row in 0..rows {
        let line: Vec<String> = (0..cols).map(|col| quote(&sheet.display(CellRef::new(row, col)), delimiter)).collect();
        out.push_str(line.join(&sep).trim_end_matches(delimiter));
        out.push_str("\r\n");
    }
    out
}

/// Guess the delimiter of delimited text: the candidate (tab, semicolon,
/// comma, pipe) that splits the first lines into the most consistent number of
/// fields (more than one). Falls back to comma.
pub fn detect_delimiter(text: &str) -> char {
    let sample: String = text.chars().take(64 * 1024).collect();
    let mut best = (',', 0usize, 0usize); // (delimiter, consistent lines, fields)
    for d in ['\t', ';', ',', '|'] {
        let rows = parse_with(&sample, d);
        let rows: Vec<_> = rows.iter().take(20).filter(|r| !(r.len() == 1 && r[0].is_empty())).collect();
        let Some(first) = rows.first() else { continue };
        let fields = first.len();
        if fields < 2 {
            continue;
        }
        let consistent = rows.iter().filter(|r| r.len() == fields).count();
        if (consistent, fields) > (best.1, best.2) {
            best = (d, consistent, fields);
        }
    }
    best.0
}

fn quote(s: &str, delimiter: char) -> String {
    if s.contains(delimiter) || s.contains(['"', '\n', '\r']) {
        format!("\"{}\"", s.replace('"', "\"\""))
    } else {
        s.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheets_engine::Value;

    #[test]
    fn parses_quotes_and_newlines() {
        let rows = parse("a,\"b,c\",\"say \"\"hi\"\"\"\r\n1,\"multi\nline\"\n");
        assert_eq!(rows, vec![vec!["a", "b,c", "say \"hi\""], vec!["1", "multi\nline"]]);
    }

    #[test]
    fn round_trips_values() {
        let sheet = read("name,qty\napple,3\npear,4\ntotal,=SUM(B2:B3)\n");
        assert_eq!(sheet.value(CellRef::new(3, 1)), &Value::Number(7.0));
        assert_eq!(write(&sheet), "name,qty\r\napple,3\r\npear,4\r\ntotal,7\r\n");
    }

    #[test]
    fn other_delimiters() {
        let sheet = read_with("name;note\r\n\"a;b\";x,y\r\n", ';');
        assert_eq!(sheet.input(CellRef::new(1, 0)), "a;b");
        assert_eq!(sheet.input(CellRef::new(1, 1)), "x,y");
        assert_eq!(write_with(&sheet, ';'), "name;note\r\n\"a;b\";x,y\r\n");
        let tsv = read_with("x\ty\n1\t2\n", '\t');
        assert_eq!(tsv.value(CellRef::new(1, 1)), &Value::Number(2.0));
        assert_eq!(write_with(&tsv, '\t'), "x\ty\r\n1\t2\r\n");
    }

    #[test]
    fn detects_delimiters() {
        assert_eq!(detect_delimiter("a,b,c\n1,2,3\n"), ',');
        assert_eq!(detect_delimiter("a;b;c\n1,5;2,5;3\n"), ';');
        assert_eq!(detect_delimiter("a\tb\n1\t2\n"), '\t');
        assert_eq!(detect_delimiter("a|b\n1|2\n"), '|');
        assert_eq!(detect_delimiter("just one column\nvalue\n"), ',');
        assert_eq!(detect_delimiter("\"x,y\";z\n\"1,2\";3\n"), ';');
    }
}
