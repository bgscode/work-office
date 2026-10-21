//! CSV import/export (RFC 4180: quoted fields, `""` escapes, CRLF or LF).

use sheets_engine::{CellRef, Sheet};

/// Split CSV text into rows of fields.
pub fn parse(text: &str) -> Vec<Vec<String>> {
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
            (',', false) => row.push(std::mem::take(&mut field)),
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
    let mut sheet = Sheet::new();
    let edits = parse(text).into_iter().enumerate().flat_map(|(r, fields)| {
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
    let (rows, cols) = sheet.extent();
    let mut out = String::new();
    for row in 0..rows {
        let line: Vec<String> = (0..cols).map(|col| quote(&sheet.display(CellRef::new(row, col)))).collect();
        out.push_str(line.join(",").trim_end_matches(','));
        out.push_str("\r\n");
    }
    out
}

fn quote(s: &str) -> String {
    if s.contains([',', '"', '\n', '\r']) {
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
}
