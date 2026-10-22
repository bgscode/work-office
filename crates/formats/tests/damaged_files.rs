//! Opening damaged or unexpected files must return a clear error (or the
//! readable part of the data), never panic: a panic in WebAssembly takes the
//! whole app down.

use std::io::{Cursor, Write};

use office_formats::{csv, xlsx};
use sheets_engine::{CellRef, Value, Workbook, MAX_COLS, MAX_ROWS};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

fn zip_of(parts: &[(&str, &str)]) -> Vec<u8> {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, body) in parts {
        zip.start_file(*name, SimpleFileOptions::default()).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

const RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;
const WORKBOOK: &str = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
const WB_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;

fn workbook_with_sheet(sheet_xml: &str, styles_xml: &str) -> Vec<u8> {
    zip_of(&[
        ("_rels/.rels", RELS),
        ("xl/workbook.xml", WORKBOOK),
        ("xl/_rels/workbook.xml.rels", WB_RELS),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/styles.xml", styles_xml),
    ])
}

fn message(bytes: &[u8]) -> String {
    match xlsx::read(bytes) {
        Ok(_) => panic!("expected an error"),
        Err(e) => e.to_string(),
    }
}

#[test]
fn rejects_files_that_are_not_zip_archives() {
    for bytes in [&b""[..], b"hello, world", b"PK\x03\x04 truncated", &[0u8; 1000]] {
        assert!(message(bytes).contains("isn't a valid .xlsx workbook"), "{bytes:?}");
    }
}

#[test]
fn explains_old_xls_and_encrypted_files() {
    let ole = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0];
    assert!(message(&ole).contains(".xls"));
}

#[test]
fn rejects_truncated_workbooks() {
    let mut book = Workbook::new();
    book.sheets[0].sheet.set_input(CellRef::new(0, 0), "hello");
    let bytes = xlsx::write(&book).unwrap();
    for len in [10, bytes.len() / 2, bytes.len() - 10] {
        // Either an error or (if only the zip directory survives) a readable book; never a panic.
        let _ = xlsx::read(&bytes[..len]);
    }
    assert!(xlsx::read(&bytes[..bytes.len() / 2]).is_err());
}

#[test]
fn zip_without_a_workbook_is_an_error() {
    let bytes = zip_of(&[("hello.txt", "not a spreadsheet")]);
    assert!(message(&bytes).contains("workbook"));
}

#[test]
fn malformed_xml_keeps_what_was_readable() {
    let sheet = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1" t="s"><v>99</v></c></row><row r="2"><c r="A2"><v>1"#;
    let book = xlsx::read(&workbook_with_sheet(sheet, "<styleSheet><cellXfs><xf")).unwrap();
    let s = &book.sheets[0].sheet;
    assert_eq!(s.value(CellRef::new(0, 0)), &Value::Number(42.0));
    // Shared string index 99 doesn't exist: the cell is left empty.
    assert_eq!(s.value(CellRef::new(0, 1)), &Value::Empty);
}

#[test]
fn ignores_cells_outside_the_sheet() {
    let sheet = format!(
        r#"<worksheet><sheetData>
            <row r="{max_row}"><c><v>1</v></c></row>
            <row r="4294967295"><c><v>2</v></c><c><v>3</v></c></row>
            <row r="3"><c r="XFD3"><v>4</v></c><c><v>5</v></c></row>
            <row r="4"><c r="A4" s="999999"><v>6</v></c></row>
        </sheetData></worksheet>"#,
        max_row = MAX_ROWS
    );
    let book = xlsx::read(&workbook_with_sheet(&sheet, "")).unwrap();
    let s = &book.sheets[0].sheet;
    let (rows, cols) = s.extent();
    assert!(rows <= MAX_ROWS && cols <= MAX_COLS, "extent {rows}x{cols}");
    assert_eq!(s.value(CellRef::new(MAX_ROWS - 1, 0)), &Value::Number(1.0));
    assert_eq!(s.value(CellRef::new(2, MAX_COLS - 1)), &Value::Number(4.0));
    assert_eq!(s.value(CellRef::new(3, 0)), &Value::Number(6.0));
}

#[test]
fn odd_values_do_not_panic() {
    let sheet = r#"<worksheet><sheetData><row r="1">
        <c r="A1" t="b"><v>maybe</v></c>
        <c r="B1" t="e"><v>#WHAT?</v></c>
        <c r="C1" t="s"><v>-1</v></c>
        <c r="D1"><f>SUM(</f><v>abc</v></c>
        <c r="E1"><f>_xlfn.</f></c>
        <c r="F1" t="inlineStr"><is><t>&#xD800;&amp;&bogus;</t></is></c>
        <c r="G1"><v>1e999</v></c>
    </row></sheetData><cols><col min="0" max="4294967295" width="-5"/><col min="2" max="2" width="NaN"/></cols></worksheet>"#;
    let book = xlsx::read(&workbook_with_sheet(sheet, "")).unwrap();
    let s = &book.sheets[0].sheet;
    assert_eq!(s.input(CellRef::new(0, 3)), "=SUM(");
    let _ = xlsx::write(&book).unwrap();
}

#[test]
fn csv_accepts_anything() {
    for text in ["", "\"unterminated", "a,b\r", "\u{feff}", "\"\"\"", ",,,\n,,,", "=1/0,=SUM(A1:A2"] {
        let sheet = csv::read(text);
        let _ = csv::write(&sheet);
    }
}

/// Every part of a real workbook (with a chart, a filter, validation and page
/// setup), cut off at many points: reading never panics.
#[test]
fn truncated_parts_never_panic() {
    use std::io::Read;
    use zip::ZipArchive;

    let mut book = Workbook::new();
    let s = &mut book.sheets[0].sheet;
    for (i, v) in ["Month", "Sales", "Jan", "10", "Feb", "20"].iter().enumerate() {
        s.set_input(CellRef::new(i as u32 / 2, i as u32 % 2), v);
    }
    let range = sheets_engine::Range::new(CellRef::new(0, 0), CellRef::new(2, 1));
    let tools = s.with_filter_applied(sheets_engine::filter::AutoFilter::new(range));
    s.set_data_tools(tools);
    let chart =
        sheets_engine::chart::chart_for_range(s, range, sheets_engine::chart::ChartType::ColumnClustered).unwrap();
    let mut drawings = s.drawings().clone();
    let anchor = sheets_engine::chart::Anchor {
        from: sheets_engine::chart::AnchorPoint::new(4, 0, 0.0, 0.0),
        to: sheets_engine::chart::AnchorPoint::new(14, 5, 0.0, 0.0),
    };
    drawings.push("Chart 1".into(), anchor, sheets_engine::chart::DrawingContent::Chart(Box::new(chart)));
    s.set_drawings(drawings);
    book.sheets[0].page_setup.set("printArea", "A1:B3").unwrap();
    let bytes = xlsx::write(&book).unwrap();

    let mut zip = ZipArchive::new(Cursor::new(&bytes[..])).unwrap();
    let parts: Vec<(String, String)> = (0..zip.len())
        .map(|i| {
            let mut f = zip.by_index(i).unwrap();
            let mut text = String::new();
            f.read_to_string(&mut text).unwrap();
            (f.name().unwrap().to_string(), text)
        })
        .collect();
    for (target, text) in &parts {
        let cuts: Vec<usize> = (1..12).map(|k| text.len() * k / 12).filter(|c| text.is_char_boundary(*c)).collect();
        for cut in cuts {
            let damaged: Vec<(&str, &str)> =
                parts.iter().map(|(n, t)| (n.as_str(), if n == target { &t[..cut] } else { t.as_str() })).collect();
            let _ = xlsx::read(&zip_of(&damaged));
        }
    }
}
