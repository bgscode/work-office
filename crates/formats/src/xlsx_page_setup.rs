//! Page setup in `.xlsx`: `<pageSetup>`, `<pageMargins>`, `<printOptions>`,
//! `<headerFooter>`, `<sheetPr><pageSetUpPr fitToPage>` in each worksheet, and
//! the `_xlnm.Print_Area` / `_xlnm.Print_Titles` defined names in workbook.xml.
//!
//! A child module of `xlsx` (it reuses its XML helpers). Writing works by
//! inserting into the finished XML strings at the position the OOXML schema
//! requires, so it keeps working whatever other elements the worksheet writer
//! emits (merged cells, validation, drawings...) and merges with defined names
//! written by other features.

use sheets_engine::page_setup::{
    format_area, format_col_span, format_row_span, parse_area, parse_col_span, parse_row_span, Orientation, PageSetup,
    Scaling,
};
use sheets_engine::{Workbook, MAX_COLS, MAX_ROWS};

use super::{attr, escape, walk, XmlEvent};

// ---------------------------------------------------------------- reading

/// Page setup stored in one worksheet part (everything except the defined names).
pub(super) fn read_sheet(xml: &str) -> PageSetup {
    let mut p = PageSetup::default();
    let mut fit_to_page = false;
    let (mut fit_w, mut fit_h) = (1u16, 1u16);
    let mut scale = 100u16;
    let mut in_text: Option<&'static str> = None;
    let on = |v: &str| v == "1" || v == "true";
    walk(xml, |ev| match ev {
        XmlEvent::Start { name: "pageSetUpPr", attrs } => fit_to_page = on(attr(attrs, "fitToPage")),
        XmlEvent::Start { name: "printOptions", attrs } => {
            p.gridlines = on(attr(attrs, "gridLines"));
            p.headings = on(attr(attrs, "headings"));
            p.center_horizontally = on(attr(attrs, "horizontalCentered"));
            p.center_vertically = on(attr(attrs, "verticalCentered"));
        }
        XmlEvent::Start { name: "pageMargins", attrs } => {
            let m = &mut p.margins;
            for (key, slot) in [
                ("left", &mut m.left),
                ("right", &mut m.right),
                ("top", &mut m.top),
                ("bottom", &mut m.bottom),
                ("header", &mut m.header),
                ("footer", &mut m.footer),
            ] {
                if let Ok(v) = attr(attrs, key).parse::<f64>() {
                    if v.is_finite() && v >= 0.0 {
                        *slot = v;
                    }
                }
            }
        }
        XmlEvent::Start { name: "pageSetup", attrs } => {
            // The schema default when the attribute is missing is Letter (1).
            p.paper_size = attr(attrs, "paperSize").parse().ok().filter(|n| *n > 0).unwrap_or(1);
            if attr(attrs, "orientation") == "landscape" {
                p.orientation = Orientation::Landscape;
            }
            if let Ok(s) = attr(attrs, "scale").parse::<u16>() {
                scale = s.clamp(10, 400);
            }
            fit_w = attr(attrs, "fitToWidth").parse().unwrap_or(1);
            fit_h = attr(attrs, "fitToHeight").parse().unwrap_or(1);
        }
        XmlEvent::Start { name: "oddHeader", .. } => in_text = Some("header"),
        XmlEvent::Start { name: "oddFooter", .. } => in_text = Some("footer"),
        XmlEvent::End { name: "oddHeader" | "oddFooter" } => in_text = None,
        XmlEvent::Text(t) => match in_text {
            Some("header") => p.header.push_str(t),
            Some("footer") => p.footer.push_str(t),
            _ => {}
        },
        _ => {}
    });
    p.scaling = if fit_to_page { Scaling::FitTo { width: fit_w, height: fit_h } } else { Scaling::Percent(scale) };
    p
}

/// Read `_xlnm.Print_Area` and `_xlnm.Print_Titles` from workbook.xml into the sheets.
/// `sheet_order` is the sheet names in workbook order (for `localSheetId`).
pub(super) fn apply_defined_names(workbook_xml: &str, sheet_order: &[String], book: &mut Workbook) {
    let mut names: Vec<(String, String, String)> = Vec::new(); // (name, localSheetId, text)
    let mut current: Option<(String, String, String)> = None;
    walk(workbook_xml, |ev| match ev {
        XmlEvent::Start { name: "definedName", attrs } => {
            current = Some((attr(attrs, "name").to_string(), attr(attrs, "localSheetId").to_string(), String::new()))
        }
        XmlEvent::Text(t) => {
            if let Some(c) = current.as_mut() {
                c.2.push_str(t);
            }
        }
        XmlEvent::End { name: "definedName" } => names.extend(current.take()),
        _ => {}
    });
    for (name, local, text) in names {
        let is_area = name.eq_ignore_ascii_case("_xlnm.Print_Area");
        let is_titles = name.eq_ignore_ascii_case("_xlnm.Print_Titles");
        if !is_area && !is_titles {
            continue;
        }
        for (sheet_name, reference) in split_references(&text) {
            let target = sheet_name
                .or_else(|| local.parse::<usize>().ok().and_then(|i| sheet_order.get(i).cloned()))
                .and_then(|n| book.sheets.iter().position(|ws| ws.name.eq_ignore_ascii_case(&n)));
            let Some(index) = target else { continue };
            let setup = &mut book.sheets[index].page_setup;
            if is_area {
                // Several areas print on separate pages in Excel; we keep the first.
                if setup.print_area.is_none() {
                    setup.print_area = parse_area(&reference);
                }
            } else if reference.trim_start_matches('$').starts_with(|c: char| c.is_ascii_digit()) {
                setup.title_rows = parse_row_span(&reference);
            } else {
                setup.title_cols = parse_col_span(&reference);
            }
        }
    }
}

/// `Sheet1!$A:$B,'My ''Sheet'!$1:$2` -> [(Some("Sheet1"), "$A:$B"), (Some("My 'Sheet"), "$1:$2")].
fn split_references(text: &str) -> Vec<(Option<String>, String)> {
    let mut parts = Vec::new();
    let (mut current, mut in_quotes) = (String::new(), false);
    for ch in text.chars() {
        match ch {
            '\'' => {
                in_quotes = !in_quotes;
                current.push(ch);
            }
            ',' if !in_quotes => parts.push(std::mem::take(&mut current)),
            _ => current.push(ch),
        }
    }
    parts.push(current);
    parts
        .into_iter()
        .filter(|p| !p.trim().is_empty())
        .map(|p| match p.trim().rsplit_once('!') {
            Some((sheet, reference)) => {
                let sheet = match sheet.strip_prefix('\'').and_then(|s| s.strip_suffix('\'')) {
                    Some(quoted) => quoted.replace("''", "'"),
                    None => sheet.to_string(),
                };
                (Some(sheet), reference.to_string())
            }
            None => (None, p.trim().to_string()),
        })
        .collect()
}

// ---------------------------------------------------------------- writing

/// Worksheet children that must come after `<headerFooter>` in CT_Worksheet.
const AFTER_HEADER_FOOTER: &[&str] = &[
    "<rowBreaks",
    "<colBreaks",
    "<customProperties",
    "<cellWatches",
    "<ignoredErrors",
    "<smartTags",
    "<drawing",
    "<legacyDrawing",
    "<legacyDrawingHF",
    "<picture",
    "<oleObjects",
    "<controls",
    "<webPublishItems",
    "<tableParts",
    "<extLst",
    "</worksheet>",
];

/// Add the page setup elements to a finished worksheet part.
pub(super) fn insert_into_worksheet(xml: String, p: &PageSetup) -> String {
    let mut xml = xml;
    if matches!(p.scaling, Scaling::FitTo { .. }) {
        xml = insert_fit_to_page(xml);
    }
    let tail = tail_xml(p);
    let at = AFTER_HEADER_FOOTER.iter().filter_map(|tag| xml.find(tag)).min().unwrap_or(xml.len());
    xml.insert_str(at, &tail);
    xml
}

/// `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`, merged into an existing `<sheetPr>`.
fn insert_fit_to_page(mut xml: String) -> String {
    const PR: &str = r#"<pageSetUpPr fitToPage="1"/>"#;
    if let Some(start) = xml.find("<sheetPr") {
        let end = start + xml[start..].find('>').unwrap_or(0);
        if xml[..end].ends_with('/') {
            // <sheetPr .../> -> <sheetPr ...><pageSetUpPr/></sheetPr>
            xml.replace_range(end - 1..=end, &format!(">{PR}</sheetPr>"));
        } else if let Some(close) = xml.find("</sheetPr>") {
            xml.insert_str(close, PR);
        }
        return xml;
    }
    // First child of <worksheet ...>.
    if let Some(open) = xml.find("<worksheet") {
        if let Some(gt) = xml[open..].find('>') {
            xml.insert_str(open + gt + 1, &format!("<sheetPr>{PR}</sheetPr>"));
        }
    }
    xml
}

fn num(v: f64) -> String {
    // Shortest form: 0.7, 1, 0.25.
    let s = format!("{:.4}", v);
    s.trim_end_matches('0').trim_end_matches('.').to_string()
}

/// `<printOptions/> <pageMargins/> <pageSetup/> <headerFooter/>`, in schema order.
fn tail_xml(p: &PageSetup) -> String {
    let mut out = String::new();
    let mut opts = String::new();
    for (key, on) in [
        ("horizontalCentered", p.center_horizontally),
        ("verticalCentered", p.center_vertically),
        ("headings", p.headings),
        ("gridLines", p.gridlines),
    ] {
        if on {
            opts.push_str(&format!(r#" {key}="1""#));
        }
    }
    if !opts.is_empty() {
        out.push_str(&format!("<printOptions{opts}/>"));
    }
    let m = &p.margins;
    out.push_str(&format!(
        r#"<pageMargins left="{}" right="{}" top="{}" bottom="{}" header="{}" footer="{}"/>"#,
        num(m.left),
        num(m.right),
        num(m.top),
        num(m.bottom),
        num(m.header),
        num(m.footer)
    ));
    let mut setup = format!(r#"<pageSetup paperSize="{}""#, p.paper_size);
    match p.scaling {
        Scaling::Percent(100) => {}
        Scaling::Percent(s) => setup.push_str(&format!(r#" scale="{s}""#)),
        Scaling::FitTo { width, height } => setup.push_str(&format!(r#" fitToWidth="{width}" fitToHeight="{height}""#)),
    }
    let orientation = match p.orientation {
        Orientation::Portrait => "portrait",
        Orientation::Landscape => "landscape",
    };
    setup.push_str(&format!(r#" orientation="{orientation}"/>"#));
    out.push_str(&setup);
    if !p.header.is_empty() || !p.footer.is_empty() {
        out.push_str("<headerFooter>");
        if !p.header.is_empty() {
            out.push_str(&format!("<oddHeader>{}</oddHeader>", escape(&p.header)));
        }
        if !p.footer.is_empty() {
            out.push_str(&format!("<oddFooter>{}</oddFooter>", escape(&p.footer)));
        }
        out.push_str("</headerFooter>");
    }
    out
}

/// Sheet name as it appears in a reference: quoted when it isn't a plain word.
fn quoted_sheet(name: &str) -> String {
    let plain = !name.is_empty()
        && !name.starts_with(|c: char| c.is_ascii_digit())
        && name.chars().all(|c| c.is_alphanumeric() || c == '_' || c == '.')
        // Names like "A1" or "R1C1" would read as cell references.
        && sheets_engine::CellRef::parse(name).is_none();
    if plain {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

/// `<definedName>` entries for print areas and titles.
fn defined_names(book: &Workbook) -> String {
    let mut out = String::new();
    for (i, ws) in book.sheets.iter().enumerate() {
        let sheet = quoted_sheet(&ws.name);
        let p = &ws.page_setup;
        if let Some(area) = p.print_area {
            let area = clamp_area(area);
            out.push_str(&format!(
                r#"<definedName name="_xlnm.Print_Area" localSheetId="{i}">{}</definedName>"#,
                escape(&format!("{sheet}!{}", format_area(area, true)))
            ));
        }
        let mut titles = Vec::new();
        if let Some(cols) = p.title_cols {
            titles.push(format!("{sheet}!{}", format_col_span(cols, true)));
        }
        if let Some(rows) = p.title_rows {
            titles.push(format!("{sheet}!{}", format_row_span(rows, true)));
        }
        if !titles.is_empty() {
            out.push_str(&format!(
                r#"<definedName name="_xlnm.Print_Titles" localSheetId="{i}">{}</definedName>"#,
                escape(&titles.join(","))
            ));
        }
    }
    out
}

fn clamp_area(r: sheets_engine::Range) -> sheets_engine::Range {
    let mut r = r;
    r.end.row = r.end.row.min(MAX_ROWS - 1);
    r.end.col = r.end.col.min(MAX_COLS - 1);
    r
}

/// Add print names to a finished workbook.xml, joining an existing
/// `<definedNames>` if another feature wrote one.
pub(super) fn insert_defined_names(xml: String, book: &Workbook) -> String {
    let names = defined_names(book);
    if names.is_empty() {
        return xml;
    }
    let mut xml = xml;
    if let Some(close) = xml.find("</definedNames>") {
        xml.insert_str(close, &names);
    } else {
        // definedNames follows <sheets> (and functionGroups/externalReferences) and precedes calcPr.
        let at = ["<calcPr", "<oleSize", "<customWorkbookViews", "<pivotCaches", "<extLst", "</workbook>"]
            .iter()
            .filter_map(|tag| xml.find(tag))
            .min()
            .unwrap_or(xml.len());
        xml.insert_str(at, &format!("<definedNames>{names}</definedNames>"));
    }
    xml
}

#[cfg(test)]
mod tests {
    use super::*;
    use sheets_engine::page_setup::Margins;
    use sheets_engine::{CellRef, Range};

    #[test]
    fn whole_file_round_trip() {
        let mut book = Workbook::new();
        book.sheets[0].sheet.set_input(CellRef::new(0, 0), "x");
        let p = &mut book.sheets[0].page_setup;
        p.set("orientation", "landscape").unwrap();
        p.set("paper", "1").unwrap();
        p.set("fit", "1,0").unwrap();
        p.set("margins", "wide").unwrap();
        p.set("headings", "1").unwrap();
        p.set("printArea", "A1:C10").unwrap();
        p.set("titleRows", "1").unwrap();
        p.set("footer", "&CPage &P of &N").unwrap();
        book.add_sheet();
        let back = crate::xlsx::read(&crate::xlsx::write(&book).unwrap()).unwrap();
        assert_eq!(back.sheets[0].page_setup, book.sheets[0].page_setup);
        assert!(back.sheets[1].page_setup.is_default());
    }

    #[test]
    fn splits_references() {
        assert_eq!(
            split_references("Sheet1!$A:$B,'My ''Q'' Sheet'!$1:$2"),
            vec![(Some("Sheet1".into()), "$A:$B".into()), (Some("My 'Q' Sheet".into()), "$1:$2".into())]
        );
        assert_eq!(split_references("$A$1:$B$2"), vec![(None, "$A$1:$B$2".into())]);
    }

    #[test]
    fn quotes_sheet_names_when_needed() {
        assert_eq!(quoted_sheet("Sheet1"), "Sheet1");
        assert_eq!(quoted_sheet("My Sheet"), "'My Sheet'");
        assert_eq!(quoted_sheet("Bob's"), "'Bob''s'");
        assert_eq!(quoted_sheet("2024"), "'2024'");
        assert_eq!(quoted_sheet("AB12"), "'AB12'");
    }

    #[test]
    fn inserts_in_schema_order() {
        let p = PageSetup {
            scaling: Scaling::FitTo { width: 1, height: 0 },
            gridlines: true,
            footer: "&CPage &P".into(),
            ..PageSetup::default()
        };
        let xml = insert_into_worksheet(
            r#"<worksheet xmlns="x"><sheetData/><mergeCells/><drawing r:id="rId1"/></worksheet>"#.into(),
            &p,
        );
        assert_eq!(
            xml,
            concat!(
                r#"<worksheet xmlns="x"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/><mergeCells/>"#,
                r#"<printOptions gridLines="1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#,
                r#"<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0" orientation="portrait"/>"#,
                r#"<headerFooter><oddFooter>&amp;CPage &amp;P</oddFooter></headerFooter><drawing r:id="rId1"/></worksheet>"#
            )
        );
        // Joins a <sheetPr> written by another feature.
        let xml = insert_fit_to_page(r#"<worksheet><sheetPr filterMode="1"/><sheetData/></worksheet>"#.into());
        assert_eq!(
            xml,
            r#"<worksheet><sheetPr filterMode="1"><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>"#
        );
        let xml = insert_fit_to_page(r#"<worksheet><sheetPr><tabColor rgb="FF0000FF"/></sheetPr></worksheet>"#.into());
        assert_eq!(
            xml,
            r#"<worksheet><sheetPr><tabColor rgb="FF0000FF"/><pageSetUpPr fitToPage="1"/></sheetPr></worksheet>"#
        );
    }

    #[test]
    fn reads_what_excel_writes() {
        let xml = concat!(
            r#"<worksheet><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/>"#,
            r#"<printOptions headings="1" gridLines="1" horizontalCentered="1"/>"#,
            r#"<pageMargins left="0.25" right="0.25" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#,
            r#"<pageSetup paperSize="5" orientation="landscape" fitToHeight="0" r:id="rId1"/>"#,
            r#"<headerFooter><oddHeader>&amp;C&amp;A</oddHeader><oddFooter>Page &amp;P of &amp;N</oddFooter></headerFooter></worksheet>"#
        );
        let p = read_sheet(xml);
        assert_eq!(p.margins, Margins::NARROW);
        assert_eq!(p.paper_size, 5);
        assert_eq!(p.orientation, Orientation::Landscape);
        assert_eq!(p.scaling, Scaling::FitTo { width: 1, height: 0 });
        assert!(p.gridlines && p.headings && p.center_horizontally && !p.center_vertically);
        assert_eq!(p.header, "&C&A");
        assert_eq!(p.footer, "Page &P of &N");
        // No page setup at all: our defaults.
        assert!(read_sheet("<worksheet><sheetData/></worksheet>").is_default());
        // Scale without fit-to-page.
        assert_eq!(read_sheet(r#"<pageSetup scale="75" fitToWidth="2"/>"#).scaling, Scaling::Percent(75));
    }

    #[test]
    fn defined_names_round_trip() {
        let mut book = Workbook::new();
        book.sheets[0].name = "My Sheet".into();
        book.add_sheet();
        book.sheets[0].page_setup.print_area = Some(Range::new(CellRef::new(0, 0), CellRef::new(19, 3)));
        book.sheets[0].page_setup.title_rows = Some((0, 1));
        book.sheets[0].page_setup.title_cols = Some((0, 0));
        book.sheets[1].page_setup.print_area = parse_area("B:C");
        let xml = insert_defined_names(r#"<workbook><sheets/><calcPr calcId="0"/></workbook>"#.into(), &book);
        assert_eq!(
            xml,
            concat!(
                r#"<workbook><sheets/><definedNames>"#,
                r#"<definedName name="_xlnm.Print_Area" localSheetId="0">'My Sheet'!$A$1:$D$20</definedName>"#,
                r#"<definedName name="_xlnm.Print_Titles" localSheetId="0">'My Sheet'!$A:$A,'My Sheet'!$1:$2</definedName>"#,
                r#"<definedName name="_xlnm.Print_Area" localSheetId="1">Sheet1!$B:$C</definedName>"#,
                r#"</definedNames><calcPr calcId="0"/></workbook>"#
            )
        );
        let mut back = Workbook::new();
        back.sheets[0].name = "My Sheet".into();
        back.add_sheet();
        apply_defined_names(&xml, &["My Sheet".into(), "Sheet1".into()], &mut back);
        assert_eq!(back.sheets[0].page_setup, book.sheets[0].page_setup);
        assert_eq!(back.sheets[1].page_setup.print_area, book.sheets[1].page_setup.print_area);

        // Joins names another feature already wrote.
        let xml = insert_defined_names(
            r#"<workbook><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">x</definedName></definedNames></workbook>"#.into(),
            &book,
        );
        assert_eq!(xml.matches("<definedNames>").count(), 1);
        assert!(xml.contains(r#"hidden="1">x</definedName><definedName name="_xlnm.Print_Area""#));
    }
}
