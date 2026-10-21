//! Excel `.xlsx` import/export.
//!
//! An `.xlsx` file is a ZIP of XML "parts":
//!
//! ```text
//! [Content_Types].xml          what each part is
//! _rels/.rels                  points to xl/workbook.xml
//! xl/workbook.xml              sheet names + relationship ids
//! xl/_rels/workbook.xml.rels   relationship id -> worksheet part
//! xl/sharedStrings.xml         text table referenced by index (optional)
//! xl/styles.xml                fonts, fills, number formats; cells point at a <cellXfs> entry via s="N"
//! xl/worksheets/sheetN.xml     <row><c r="A1" t="..."><f>formula</f><v>value</v></c></row>
//! ```
//!
//! Supported today: values (numbers, text, booleans, errors), formulas, cell
//! formatting (bold/italic/underline/strike, text and fill colours, horizontal
//! alignment, number formats) and column widths.
//! Not yet: merged cells, borders, fonts/sizes, row heights, charts.
//! Excel "shared formulas" are imported as their cached values for every cell
//! except the first, because we don't yet rewrite references when copying formulas.

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use office_core::{OfficeError, Result};
use quick_xml::escape::resolve_predefined_entity;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use sheets_engine::{Align, CellRef, ErrorKind, NumberFormat, Sheet, Style, Value, Workbook};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[path = "xlsx_data_tools.rs"]
mod data_tools;

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Functions added after Excel 2007 must be written as `_xlfn.NAME` in files.
const NEWER_FUNCTIONS: &[&str] = &[
    "AVERAGEIFS",
    "CONCAT",
    "COUNTIFS",
    "DAYS",
    "IFNA",
    "IFS",
    "MAXIFS",
    "MINIFS",
    "MODE.SNGL",
    "RANK.EQ",
    "STDEV.P",
    "STDEV.S",
    "SUMIFS",
    "SWITCH",
    "TEXTJOIN",
    "VAR.P",
    "VAR.S",
    "XLOOKUP",
    "XOR",
];

// ---------------------------------------------------------------- reading

pub fn read(bytes: &[u8]) -> Result<Workbook> {
    let mut zip =
        ZipArchive::new(Cursor::new(bytes)).map_err(|e| OfficeError::Parse(format!("not a zip file: {e}")))?;

    let workbook_path = part(&mut zip, "_rels/.rels")?
        .and_then(|xml| relationships(&xml).into_iter().find(|r| r.kind.ends_with("/officeDocument")))
        .map(|r| r.target.trim_start_matches('/').to_string())
        .unwrap_or_else(|| "xl/workbook.xml".into());
    let base = workbook_path.rsplit_once('/').map_or("", |(dir, _)| dir).to_string();
    let resolve = |target: &str| match target.strip_prefix('/') {
        Some(abs) => abs.to_string(),
        None if base.is_empty() => target.to_string(),
        None => format!("{base}/{target}"),
    };

    let workbook_xml =
        part(&mut zip, &workbook_path)?.ok_or_else(|| OfficeError::Parse("missing workbook.xml".into()))?;
    let rels_path = match workbook_path.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{workbook_path}.rels"),
    };
    let rels = part(&mut zip, &rels_path)?.map(|xml| relationships(&xml)).unwrap_or_default();

    let shared = match rels.iter().find(|r| r.kind.ends_with("/sharedStrings")) {
        Some(r) => part(&mut zip, &resolve(&r.target))?.map(|xml| shared_strings(&xml)).unwrap_or_default(),
        None => Vec::new(),
    };

    let styles = match rels.iter().find(|r| r.kind.ends_with("/styles")) {
        Some(r) => part(&mut zip, &resolve(&r.target))?.map(|xml| read_styles(&xml)).unwrap_or_default(),
        None => Vec::new(),
    };

    // Conditional-format styles (`<dxfs>`), for the Data tools.
    let dxfs = match rels.iter().find(|r| r.kind.ends_with("/styles")) {
        Some(r) => part(&mut zip, &resolve(&r.target))?.map(|xml| data_tools::read_dxfs(&xml)).unwrap_or_default(),
        None => Vec::new(),
    };

    let mut book = Workbook::empty();
    for (name, rid) in sheet_list(&workbook_xml) {
        let Some(rel) = rels.iter().find(|r| r.id == rid) else { continue };
        let xml = part(&mut zip, &resolve(&rel.target))?.unwrap_or_default();
        let data = sheet_data(&xml, &shared);
        let mut sheet = Sheet::new();
        // Styles first, so typed-format detection doesn't override them.
        for (at, index) in data.styles {
            if let Some(style) = styles.get(index) {
                sheet.set_style(at, style.clone());
            }
        }
        for (col, width) in data.widths {
            sheet.set_col_width(col, Some(width));
        }
        sheet.set_inputs(data.cells);
        data_tools::apply(&mut sheet, data_tools::read_sheet(&xml, &dxfs));
        book.push_sheet(name, sheet);
    }
    if book.sheets.is_empty() {
        return Err(OfficeError::Parse("workbook has no worksheets".into()));
    }
    Ok(book)
}

fn part(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<Option<String>> {
    let mut file = match zip.by_name(name) {
        Ok(f) => f,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(e) => return Err(OfficeError::Parse(e.to_string())),
    };
    let mut s = String::new();
    file.read_to_string(&mut s).map_err(|e| OfficeError::Io(e.to_string()))?;
    Ok(Some(s))
}

struct Relationship {
    id: String,
    kind: String,
    target: String,
}

fn relationships(xml: &str) -> Vec<Relationship> {
    let mut out = Vec::new();
    walk(xml, |ev| {
        if let XmlEvent::Start { name: "Relationship", attrs } = ev {
            out.push(Relationship {
                id: attr(attrs, "Id").to_string(),
                kind: attr(attrs, "Type").to_string(),
                target: attr(attrs, "Target").to_string(),
            });
        }
    });
    out
}

/// (sheet name, relationship id) in workbook order.
fn sheet_list(xml: &str) -> Vec<(String, String)> {
    let mut out = Vec::new();
    walk(xml, |ev| {
        if let XmlEvent::Start { name: "sheet", attrs } = ev {
            out.push((attr(attrs, "name").to_string(), attr(attrs, "id").to_string()));
        }
    });
    out
}

fn shared_strings(xml: &str) -> Vec<String> {
    let mut out = Vec::new();
    let (mut current, mut in_t, mut in_phonetic) = (String::new(), false, false);
    walk(xml, |ev| match ev {
        XmlEvent::Start { name: "si", .. } => current.clear(),
        XmlEvent::End { name: "si" } => out.push(std::mem::take(&mut current)),
        XmlEvent::Start { name: "rPh", .. } => in_phonetic = true,
        XmlEvent::End { name: "rPh" } => in_phonetic = false,
        XmlEvent::Start { name: "t", .. } => in_t = true,
        XmlEvent::End { name: "t" } => in_t = false,
        XmlEvent::Text(t) if in_t && !in_phonetic => current.push_str(t),
        _ => {}
    });
    out
}

#[derive(Default)]
struct SheetData {
    /// (position, input string) ready for `Sheet::set_inputs`.
    cells: Vec<(CellRef, String)>,
    /// (position, index into cellXfs)
    styles: Vec<(CellRef, usize)>,
    /// (column, width in pixels)
    widths: Vec<(u32, f64)>,
}

fn sheet_data(xml: &str, shared: &[String]) -> SheetData {
    #[derive(Default)]
    struct C {
        at: Option<CellRef>,
        kind: String,
        formula: String,
        value: String,
        inline: String,
    }
    #[derive(Clone, Copy, PartialEq)]
    enum In {
        None,
        F,
        V,
        InlineT,
    }
    let mut out = SheetData::default();
    let mut c = C::default();
    let mut inside = In::None;
    let (mut row, mut next_col) = (0u32, 0u32);
    walk(xml, |ev| match ev {
        XmlEvent::Start { name: "col", attrs } => {
            let (Ok(min), Ok(max), Ok(width)) = (
                attr(attrs, "min").parse::<u32>(),
                attr(attrs, "max").parse::<u32>(),
                attr(attrs, "width").parse::<f64>(),
            ) else {
                return;
            };
            // Whole-sheet column defaults (max = 16384) aren't worth expanding.
            if min >= 1 && max >= min && max - min < 1000 {
                for col in min..=max {
                    out.widths.push((col - 1, chars_to_px(width)));
                }
            }
        }
        XmlEvent::Start { name: "row", attrs } => {
            if let Ok(r) = attr(attrs, "r").parse::<u32>() {
                row = r.saturating_sub(1);
            }
            next_col = 0;
        }
        XmlEvent::End { name: "row" } => row += 1,
        XmlEvent::Start { name: "c", attrs } => {
            let at = CellRef::parse(attr(attrs, "r")).unwrap_or(CellRef::new(row, next_col));
            next_col = at.col + 1;
            if let Ok(index) = attr(attrs, "s").parse::<usize>() {
                if index > 0 {
                    out.styles.push((at, index));
                }
            }
            c = C { at: Some(at), kind: attr(attrs, "t").to_string(), ..C::default() };
        }
        XmlEvent::Start { name: "f", .. } => inside = In::F,
        XmlEvent::Start { name: "v", .. } => inside = In::V,
        XmlEvent::Start { name: "t", .. } if c.kind == "inlineStr" => inside = In::InlineT,
        XmlEvent::End { name: "f" | "v" | "t" } => inside = In::None,
        XmlEvent::Text(t) => match inside {
            In::F => c.formula.push_str(t),
            In::V => c.value.push_str(t),
            In::InlineT => c.inline.push_str(t),
            In::None => {}
        },
        XmlEvent::End { name: "c" } => {
            let input = if !c.formula.is_empty() {
                format!("={}", strip_function_prefixes(&c.formula))
            } else {
                match c.kind.as_str() {
                    "s" => c
                        .value
                        .trim()
                        .parse::<usize>()
                        .ok()
                        .and_then(|i| shared.get(i))
                        .map_or(String::new(), |s| text_input(s)),
                    "inlineStr" => text_input(&c.inline),
                    "str" => text_input(&c.value),
                    "b" => if c.value.trim() == "1" { "TRUE" } else { "FALSE" }.to_string(),
                    _ => c.value.trim().to_string(),
                }
            };
            if let (Some(at), false) = (c.at, input.is_empty()) {
                out.cells.push((at, input));
            }
        }
        _ => {}
    });
    out
}

/// Excel column width (in characters of the default font) to pixels, and back.
fn chars_to_px(width: f64) -> f64 {
    (width * 7.0 + 5.0).round()
}

fn px_to_chars(px: f64) -> f64 {
    ((px - 5.0) / 7.0 * 100.0).round() / 100.0
}

/// Built-in number format codes that files reference by id without defining them.
fn builtin_format(id: u32) -> Option<&'static str> {
    Some(match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        14 => "dd-mm-yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        49 => "@",
        _ => return None,
    })
}

/// One `Style` per `<cellXfs>` entry, in order.
fn read_styles(xml: &str) -> Vec<Style> {
    #[derive(Default, Clone)]
    struct Font {
        bold: bool,
        italic: bool,
        underline: bool,
        strike: bool,
        color: Option<String>,
    }
    let mut num_fmts: HashMap<u32, String> = HashMap::new();
    let mut fonts: Vec<Font> = Vec::new();
    let mut fills: Vec<Option<String>> = Vec::new();
    let mut out = Vec::new();
    let (mut in_font, mut in_fill, mut in_cell_xfs) = (false, false, false);
    let mut pattern_solid = false;
    walk(xml, |ev| match ev {
        XmlEvent::Start { name: "numFmt", attrs } => {
            if let Ok(id) = attr(attrs, "numFmtId").parse() {
                num_fmts.insert(id, attr(attrs, "formatCode").to_string());
            }
        }
        XmlEvent::Start { name: "font", .. } => {
            in_font = true;
            fonts.push(Font::default());
        }
        XmlEvent::End { name: "font" } => in_font = false,
        XmlEvent::Start { name, attrs } if in_font => {
            let font = fonts.last_mut().unwrap();
            let on = attr(attrs, "val") != "0" && attr(attrs, "val") != "false";
            match *name {
                "b" => font.bold = on,
                "i" => font.italic = on,
                "strike" => font.strike = on,
                "u" => font.underline = attr(attrs, "val") != "none",
                "color" => font.color = argb(attr(attrs, "rgb")),
                _ => {}
            }
        }
        XmlEvent::Start { name: "fill", .. } => {
            in_fill = true;
            pattern_solid = false;
            fills.push(None);
        }
        XmlEvent::End { name: "fill" } => in_fill = false,
        XmlEvent::Start { name: "patternFill", attrs } if in_fill => {
            pattern_solid = attr(attrs, "patternType") == "solid"
        }
        XmlEvent::Start { name: "fgColor", attrs } if in_fill && pattern_solid => {
            *fills.last_mut().unwrap() = argb(attr(attrs, "rgb"));
        }
        XmlEvent::Start { name: "cellXfs", .. } => in_cell_xfs = true,
        XmlEvent::End { name: "cellXfs" } => in_cell_xfs = false,
        XmlEvent::Start { name: "xf", attrs } if in_cell_xfs => {
            let index = |key: &str| attr(attrs, key).parse::<usize>().unwrap_or(0);
            let font = fonts.get(index("fontId")).cloned().unwrap_or_default();
            let fmt_id = index("numFmtId") as u32;
            let code =
                num_fmts.get(&fmt_id).map(String::as_str).or_else(|| builtin_format(fmt_id)).unwrap_or("General");
            out.push(Style {
                bold: font.bold,
                italic: font.italic,
                underline: font.underline,
                strike: font.strike,
                color: font.color.filter(|c| c != "#000000"),
                fill: fills.get(index("fillId")).cloned().flatten(),
                align: None,
                number_format: NumberFormat::from_code(code),
            });
        }
        XmlEvent::Start { name: "alignment", attrs } if in_cell_xfs => {
            if let Some(style) = out.last_mut() {
                style.align = match attr(attrs, "horizontal") {
                    "left" => Some(Align::Left),
                    "center" | "centerContinuous" => Some(Align::Center),
                    "right" => Some(Align::Right),
                    _ => None,
                };
            }
        }
        _ => {}
    });
    out
}

/// "FFRRGGBB" (or "RRGGBB") -> "#RRGGBB".
fn argb(s: &str) -> Option<String> {
    let hex = match s.len() {
        8 => &s[2..],
        6 => s,
        _ => return None,
    };
    hex.chars().all(|c| c.is_ascii_hexdigit()).then(|| format!("#{}", hex.to_ascii_uppercase()))
}

/// Text that would be misread as a number, boolean or formula gets a leading `'`.
fn text_input(s: &str) -> String {
    Value::Text(s.to_string()).to_input()
}

fn strip_function_prefixes(formula: &str) -> String {
    formula.replace("_xlfn.", "").replace("_xlws.", "")
}

// ---------------------------------------------------------------- writing

pub fn write(book: &Workbook) -> Result<Vec<u8>> {
    let mut parts: Vec<(String, String)> = Vec::new();
    let n = book.sheets.len();
    let mut styles = StyleTable::default();

    let mut types = String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        r#"<Default Extension="xml" ContentType="application/xml"/>"#,
        r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
        r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
    ));
    for i in 1..=n {
        types.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ));
    }
    types.push_str("</Types>");
    parts.push(("[Content_Types].xml".into(), types));

    parts.push((
        "_rels/.rels".into(),
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{NS_PKG_REL}"><Relationship Id="rId1" Type="{NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        ),
    ));

    let mut workbook = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}"><sheets>"#
    );
    let mut rels =
        format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{NS_PKG_REL}">"#);
    for (i, ws) in book.sheets.iter().enumerate() {
        let id = i + 1;
        workbook.push_str(&format!(r#"<sheet name="{}" sheetId="{id}" r:id="rId{id}"/>"#, escape(&ws.name)));
        rels.push_str(&format!(
            r#"<Relationship Id="rId{id}" Type="{NS_REL}/worksheet" Target="worksheets/sheet{id}.xml"/>"#
        ));
        parts.push((format!("xl/worksheets/sheet{id}.xml"), worksheet_xml(&ws.sheet, &mut styles)));
    }
    // Ask Excel to recalculate on open, since our cached values may differ slightly.
    workbook.push_str(r#"</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>"#);
    rels.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="{NS_REL}/styles" Target="styles.xml"/></Relationships>"#,
        n + 1
    ));
    parts.push(("xl/workbook.xml".into(), workbook));
    parts.push(("xl/_rels/workbook.xml.rels".into(), rels));
    parts.push(("xl/styles.xml".into(), styles.finish_xml()));

    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, xml) in parts {
        zip.start_file(name, options).map_err(|e| OfficeError::Io(e.to_string()))?;
        zip.write_all(xml.as_bytes()).map_err(|e| OfficeError::Io(e.to_string()))?;
    }
    Ok(zip.finish().map_err(|e| OfficeError::Io(e.to_string()))?.into_inner())
}

/// Collects the distinct styles used while writing sheets, then emits
/// `styles.xml`. Index 0 of every list is Excel's required default entry.
#[derive(Default)]
struct StyleTable {
    xfs: Vec<Style>,
    fonts: Vec<(bool, bool, bool, bool, Option<String>)>,
    fills: Vec<String>,
    custom_formats: Vec<String>,
    /// Conditional-format styles, referenced by `dxfId` (see `xlsx_data_tools.rs`).
    dxfs: Vec<sheets_engine::conditional_format::CfFormat>,
}

impl StyleTable {
    /// The `s` attribute for a cell with this style (0 = default).
    fn index(&mut self, style: &Style) -> usize {
        if style.is_default() {
            return 0;
        }
        match self.xfs.iter().position(|s| s == style) {
            Some(i) => i + 1,
            None => {
                self.xfs.push(style.clone());
                self.xfs.len()
            }
        }
    }

    fn finish_xml(&mut self) -> String {
        let mut cell_xfs = String::from(r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#);
        for style in self.xfs.clone() {
            let font_key = (style.bold, style.italic, style.underline, style.strike, style.color.clone());
            let font_id = 1 + position_or_push(&mut self.fonts, font_key);
            let fill_id = style.fill.clone().map_or(0, |f| 2 + position_or_push(&mut self.fills, f));
            let fmt_id = self.format_id(&style.number_format);
            let align = match style.align {
                Some(Align::Left) => r#"<alignment horizontal="left"/>"#,
                Some(Align::Center) => r#"<alignment horizontal="center"/>"#,
                Some(Align::Right) => r#"<alignment horizontal="right"/>"#,
                None => "",
            };
            let apply_align = if align.is_empty() { "" } else { r#" applyAlignment="1""# };
            let inner_close = if align.is_empty() { "/>".to_string() } else { format!(">{align}</xf>") };
            cell_xfs.push_str(&format!(
                r#"<xf numFmtId="{fmt_id}" fontId="{font_id}" fillId="{fill_id}" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"{apply_align}{inner_close}"#
            ));
        }

        let mut xml =
            format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{NS_MAIN}">"#);
        if !self.custom_formats.is_empty() {
            xml.push_str(&format!(r#"<numFmts count="{}">"#, self.custom_formats.len()));
            for (i, code) in self.custom_formats.iter().enumerate() {
                xml.push_str(&format!(r#"<numFmt numFmtId="{}" formatCode="{}"/>"#, 164 + i, escape(code)));
            }
            xml.push_str("</numFmts>");
        }
        xml.push_str(&format!(
            r#"<fonts count="{}"><font><sz val="11"/><name val="Calibri"/></font>"#,
            self.fonts.len() + 1
        ));
        for (bold, italic, underline, strike, color) in &self.fonts {
            xml.push_str("<font>");
            if *bold {
                xml.push_str("<b/>");
            }
            if *italic {
                xml.push_str("<i/>");
            }
            if *strike {
                xml.push_str("<strike/>");
            }
            if *underline {
                xml.push_str("<u/>");
            }
            xml.push_str(r#"<sz val="11"/>"#);
            if let Some(c) = color {
                xml.push_str(&format!(r#"<color rgb="FF{}"/>"#, c.trim_start_matches('#')));
            }
            xml.push_str(r#"<name val="Calibri"/></font>"#);
        }
        xml.push_str("</fonts>");
        xml.push_str(&format!(
            r#"<fills count="{}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>"#,
            self.fills.len() + 2
        ));
        for fill in &self.fills {
            xml.push_str(&format!(
                r#"<fill><patternFill patternType="solid"><fgColor rgb="FF{}"/><bgColor indexed="64"/></patternFill></fill>"#,
                fill.trim_start_matches('#')
            ));
        }
        xml.push_str("</fills>");
        xml.push_str(r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#);
        xml.push_str(r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#);
        xml.push_str(&format!(r#"<cellXfs count="{}">{cell_xfs}</cellXfs>"#, self.xfs.len() + 1));
        xml.push_str(r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#);
        xml.push_str(&self.dxfs_xml());
        xml.push_str("</styleSheet>");
        xml
    }

    fn format_id(&mut self, nf: &NumberFormat) -> usize {
        let code = nf.to_code();
        if let Some(id) = (0..50).find(|id| builtin_format(*id) == Some(code.as_str())) {
            return id as usize;
        }
        164 + position_or_push(&mut self.custom_formats, code)
    }
}

fn position_or_push<T: PartialEq>(list: &mut Vec<T>, item: T) -> usize {
    match list.iter().position(|x| *x == item) {
        Some(i) => i,
        None => {
            list.push(item);
            list.len() - 1
        }
    }
}

fn worksheet_xml(sheet: &Sheet, styles: &mut StyleTable) -> String {
    let mut xml = format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS_MAIN}">"#);
    let widths: Vec<(u32, f64)> = sheet.col_widths().collect();
    if !widths.is_empty() {
        xml.push_str("<cols>");
        for (col, px) in widths {
            let c = col + 1;
            xml.push_str(&format!(r#"<col min="{c}" max="{c}" width="{}" customWidth="1"/>"#, px_to_chars(px)));
        }
        xml.push_str("</cols>");
    }
    xml.push_str("<sheetData>");
    // Cells with content plus formatted empty cells, row by row.
    let mut positions: Vec<CellRef> = sheet.cells().into_iter().map(|(at, _, _)| at).collect();
    positions.extend(sheet.styled_cells().into_iter().map(|(at, _)| at));
    positions.sort();
    positions.dedup();
    let mut current_row = None;
    for at in positions {
        if current_row != Some(at.row) {
            if current_row.is_some() {
                xml.push_str("</row>");
            }
            xml.push_str(&format!(r#"<row r="{}">"#, at.row + 1));
            current_row = Some(at.row);
        }
        let s = styles.index(sheet.style(at));
        xml.push_str(&cell_xml(at, sheet.input(at), sheet.value(at), s));
    }
    if current_row.is_some() {
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData>");
    xml.push_str(&data_tools::autofilter_xml(sheet));
    // <mergeCells> belongs here, between <autoFilter> and <conditionalFormatting>.
    xml.push_str(&data_tools::cf_and_validation_xml(sheet, styles));
    xml.push_str("</worksheet>");
    xml
}

fn cell_xml(at: CellRef, input: &str, value: &Value, style: usize) -> String {
    // `r="A1"` plus the style index when the cell is formatted.
    let r = if style == 0 { format!(r#"r="{at}""#) } else { format!(r#"r="{at}" s="{style}""#) };
    if let Some(formula) = input.strip_prefix('=') {
        let f = escape(&add_function_prefixes(formula));
        return match value {
            Value::Number(n) => format!(r#"<c {r}><f>{f}</f><v>{n}</v></c>"#),
            Value::Text(s) => format!(r#"<c {r} t="str"><f>{f}</f><v>{}</v></c>"#, escape(s)),
            Value::Bool(b) => format!(r#"<c {r} t="b"><f>{f}</f><v>{}</v></c>"#, *b as u8),
            Value::Error(e) => format!(r#"<c {r} t="e"><f>{f}</f><v>{}</v></c>"#, excel_error(*e)),
            Value::Empty => format!(r#"<c {r}><f>{f}</f></c>"#),
        };
    }
    match value {
        Value::Number(n) => format!(r#"<c {r}><v>{n}</v></c>"#),
        Value::Text(s) => format!(r#"<c {r} t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>"#, escape(s)),
        Value::Bool(b) => format!(r#"<c {r} t="b"><v>{}</v></c>"#, *b as u8),
        Value::Error(e) => format!(r#"<c {r} t="e"><v>{}</v></c>"#, excel_error(*e)),
        Value::Empty if style > 0 => format!(r#"<c {r}/>"#),
        Value::Empty => String::new(),
    }
}

/// Our extra error kinds don't exist in Excel.
fn excel_error(e: ErrorKind) -> &'static str {
    match e {
        ErrorKind::Circ | ErrorKind::Parse => ErrorKind::Value.as_str(),
        other => other.as_str(),
    }
}

/// `CONCAT(` -> `_xlfn.CONCAT(`, leaving string literals untouched.
fn add_function_prefixes(formula: &str) -> String {
    let mut out = String::with_capacity(formula.len());
    let mut word = String::new();
    let mut in_string = false;
    for ch in formula.chars() {
        if in_string {
            out.push(ch);
            in_string = ch != '"';
            continue;
        }
        if ch.is_alphanumeric() || ch == '_' || ch == '.' {
            word.push(ch);
            continue;
        }
        if ch == '(' && NEWER_FUNCTIONS.iter().any(|f| f.eq_ignore_ascii_case(&word)) {
            out.push_str("_xlfn.");
        }
        out.push_str(&word);
        word.clear();
        out.push(ch);
        in_string = ch == '"';
    }
    out.push_str(&word);
    out
}

fn escape(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

// ---------------------------------------------------------------- XML helper

/// The only place that talks to quick-xml. Element and attribute names have
/// their namespace prefix removed (`r:id` -> `id`); text is unescaped.
enum XmlEvent<'a> {
    Start { name: &'a str, attrs: &'a [(String, String)] },
    End { name: &'a str },
    Text(&'a str),
}

fn walk(xml: &str, mut on: impl FnMut(&XmlEvent)) {
    let v = XmlVersion::Implicit1_0;
    let mut reader = Reader::from_str(xml);
    loop {
        // Malformed XML ends the walk; whatever was read so far is kept.
        let Ok(event) = reader.read_event() else { return };
        match event {
            Event::Start(e) => start(&e, false, &mut on),
            Event::Empty(e) => start(&e, true, &mut on),
            Event::End(e) => on(&XmlEvent::End { name: e.local_name().as_ref() }),
            Event::Text(t) => on(&XmlEvent::Text(&t.xml_content(v))),
            Event::CData(t) => on(&XmlEvent::Text(&t.xml_content(v))),
            Event::GeneralRef(r) => {
                let text = match r.resolve_char_ref() {
                    Ok(Some(ch)) => ch.to_string(),
                    _ => resolve_predefined_entity(&r.xml_content(v)).unwrap_or_default().to_string(),
                };
                on(&XmlEvent::Text(&text));
            }
            Event::Eof => return,
            _ => {}
        }
    }
}

fn start(e: &BytesStart, is_empty: bool, on: &mut impl FnMut(&XmlEvent)) {
    let name = e.local_name();
    let attrs: Vec<(String, String)> = e
        .attributes()
        .flatten()
        .map(|a| {
            let value = a.normalized_value(XmlVersion::Implicit1_0).map(|c| c.into_owned()).unwrap_or_default();
            (a.key.local_name().as_ref().to_string(), value)
        })
        .collect();
    on(&XmlEvent::Start { name: name.as_ref(), attrs: &attrs });
    if is_empty {
        on(&XmlEvent::End { name: name.as_ref() });
    }
}

fn attr<'a>(attrs: &'a [(String, String)], key: &str) -> &'a str {
    attrs.iter().find(|(k, _)| k == key).map_or("", |(_, v)| v.as_str())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn prefixes_newer_functions_outside_strings() {
        assert_eq!(
            add_function_prefixes(r#"CONCAT("CONCAT(",A1)&textjoin(",",1,B1)"#),
            r#"_xlfn.CONCAT("CONCAT(",A1)&_xlfn.textjoin(",",1,B1)"#
        );
        assert_eq!(strip_function_prefixes("_xlfn.CONCAT(A1)"), "CONCAT(A1)");
    }

    #[test]
    fn round_trips_a_workbook() {
        let mut book = Workbook::new();
        let s = &mut book.sheets[0].sheet;
        s.set_inputs([
            (CellRef::new(0, 0), "Item & <name>".to_string()),
            (CellRef::new(0, 1), "'00123".to_string()),
            (CellRef::new(1, 0), "2.5".to_string()),
            (CellRef::new(1, 1), "=A2*4".to_string()),
            (CellRef::new(2, 0), "TRUE".to_string()),
            (CellRef::new(2, 1), "=CONCAT(\"a\",\"b\")".to_string()),
            (CellRef::new(3, 0), "=1/0".to_string()),
        ]);
        let other = book.add_sheet();
        book.sheets[other].sheet.set_input(CellRef::new(4, 2), "second");

        let bytes = write(&book).unwrap();
        let back = read(&bytes).unwrap();

        assert_eq!(back.sheets.len(), 2);
        assert_eq!(back.sheets[1].name, "Sheet2");
        let s = &back.sheets[0].sheet;
        assert_eq!(s.input(CellRef::new(0, 0)), "Item & <name>");
        assert_eq!(s.input(CellRef::new(0, 1)), "'00123");
        assert_eq!(s.value(CellRef::new(1, 1)), &Value::Number(10.0));
        assert_eq!(s.input(CellRef::new(1, 1)), "=A2*4");
        assert_eq!(s.value(CellRef::new(2, 0)), &Value::Bool(true));
        assert_eq!(s.value(CellRef::new(2, 1)), &Value::Text("ab".into()));
        assert_eq!(s.value(CellRef::new(3, 0)), &Value::Error(ErrorKind::Div0));
        assert_eq!(back.sheets[1].sheet.input(CellRef::new(4, 2)), "second");
    }

    #[test]
    fn round_trips_formatting_and_widths() {
        let mut book = Workbook::new();
        let s = &mut book.sheets[0].sheet;
        let fancy = Style {
            bold: true,
            italic: true,
            underline: true,
            strike: false,
            color: Some("#C00000".into()),
            fill: Some("#FFF2CC".into()),
            align: Some(Align::Center),
            number_format: NumberFormat::Currency { symbol: "₹".into(), decimals: 2 },
        };
        s.set_style(CellRef::new(0, 0), fancy.clone());
        s.set_input(CellRef::new(0, 0), "1234.5");
        s.set_input(CellRef::new(1, 0), "12.5%");
        s.set_input(CellRef::new(2, 0), "15-08-2025");
        s.set_style(CellRef::new(5, 5), Style { bold: true, ..Style::default() });
        s.set_col_width(0, Some(180.0));

        let back = read(&write(&book).unwrap()).unwrap();
        let s = &back.sheets[0].sheet;
        assert_eq!(s.style(CellRef::new(0, 0)), &fancy);
        assert_eq!(s.display(CellRef::new(0, 0)), "₹1,234.50");
        assert_eq!(s.display(CellRef::new(1, 0)), "12.5%");
        assert_eq!(s.display(CellRef::new(2, 0)), "15-08-2025");
        assert!(s.style(CellRef::new(5, 5)).bold);
        assert_eq!(s.col_width(0), 180.0);
    }
}
