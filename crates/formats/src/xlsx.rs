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
//! formatting (font name and size, bold/italic/underline/strike, text and fill
//! colours, borders, horizontal and vertical alignment, wrap text, number
//! formats), column widths, and the sheet layout in `xlsx_layout.rs` (row
//! heights, hidden rows/columns, merged cells, frozen panes).
//! Charts and pictures: see `xlsx/drawings.rs`.
//! Not yet: theme/indexed colours (read as automatic).
//! Excel "shared formulas" are imported as their cached values for every cell
//! except the first, because we don't yet rewrite references when copying formulas.

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use office_core::{OfficeError, Result};
use sheets_engine::{Align, CellRef, ErrorKind, NumberFormat, Sheet, Style, VAlign, Value, Workbook};
use sheets_engine::{BorderLine, BorderStyle, Borders, Side};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

mod drawings;
use crate::xml::{attr, walk, XmlEvent};

#[path = "xlsx_data_tools.rs"]
mod data_tools;
#[path = "xlsx_layout.rs"]
mod layout;
#[path = "xlsx_page_setup.rs"]
mod page_setup;

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

    let styles_xml = match rels.iter().find(|r| r.kind.ends_with("/styles")) {
        Some(r) => part(&mut zip, &resolve(&r.target))?.unwrap_or_default(),
        None => String::new(),
    };
    let styles = read_styles(&styles_xml);
    // Conditional-format styles (`<dxfs>`), for the Data tools.
    let dxfs = data_tools::read_dxfs(&styles_xml);
    drop(styles_xml);
    let shared = SharedStrings::new(shared);
    let mut book = Workbook::empty();
    let sheet_order: Vec<String> = sheet_list(&workbook_xml).into_iter().map(|(name, _)| name).collect();
    for (name, rid) in sheet_list(&workbook_xml) {
        let Some(rel) = rels.iter().find(|r| r.id == rid) else { continue };
        let mut timer = Timer::start();
        let xml = part(&mut zip, &resolve(&rel.target))?.unwrap_or_default();
        let unzip = timer.lap();
        let data = sheet_data(&xml, &shared, &dxfs);
        // Page setup is in <sheetPr> before the cells and after </sheetData>;
        // read those parts only, not the cells again.
        let page_setup = match (xml.find("<sheetData"), xml.rfind("</sheetData>")) {
            (Some(a), Some(b)) if a < b => page_setup::read_sheet(&format!("{}{}", &xml[..a], &xml[b..])),
            _ => page_setup::read_sheet(&xml),
        };
        drop(xml);
        let parse = timer.lap();
        let mut sheet = Sheet::new();
        // Styles first: a cell formatted as Text keeps its input as text.
        for (at, index) in data.styles {
            if let Some(style) = styles.get(index) {
                sheet.set_style(at, style.clone());
            }
        }
        for (col, width) in data.widths {
            sheet.set_col_width(col, Some(width));
        }
        let cells = data.cells.len();
        sheet.load_cells(data.cells);
        *sheet.layout_mut() = data.layout;
        data_tools::apply(&mut sheet, data.tools);
        drawings::read_sheet_drawings(&mut zip, &resolve(&rel.target), &name, &mut sheet)?;
        let load = timer.lap();
        Timer::report(|| format!("{name}: unzip {unzip}, parse {parse} ({cells} cells), load+calc {load}"));
        let index = book.push_sheet(name, sheet);
        book.sheets[index].page_setup = page_setup;
    }
    page_setup::apply_defined_names(&workbook_xml, &sheet_order, &mut book);
    if book.sheets.is_empty() {
        return Err(OfficeError::Parse("workbook has no worksheets".into()));
    }
    // Sheets were calculated one by one; formulas reading other sheets need the whole book.
    book.recalc_links();
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
    /// (position, input, typed value) ready for `Sheet::load_cells`. The value
    /// is `None` for formulas, which the sheet parses and calculates; the input
    /// is `None` when it is just the value written back.
    cells: Vec<(CellRef, Option<String>, Option<Value>)>,
    /// (position, index into cellXfs)
    styles: Vec<(CellRef, usize)>,
    /// (column, width in pixels)
    widths: Vec<(u32, f64)>,
    layout: sheets_engine::layout::SheetLayout,
    tools: data_tools::SheetTools,
}

/// Shared strings, with the input of the few that need a leading `'` worked out once.
struct SharedStrings {
    text: Vec<String>,
    input: Vec<Option<String>>,
}

impl SharedStrings {
    fn new(text: Vec<String>) -> Self {
        let input = text.iter().map(|s| text_cell_input(s)).collect();
        SharedStrings { text, input }
    }
}

/// The input to store for a text cell: `None` when it is the text itself.
fn text_cell_input(s: &str) -> Option<String> {
    (!Value::text_round_trips(s)).then(|| format!("'{s}"))
}

/// One pass over a worksheet: cells, styles, column widths and the layout
/// (row heights, hidden lines, merges, frozen panes). Hot loop for big
/// files, so it reuses its buffers instead of allocating per cell.
fn sheet_data(xml: &str, shared: &SharedStrings, dxfs: &[sheets_engine::conditional_format::CfFormat]) -> SheetData {
    #[derive(Clone, Copy, PartialEq)]
    enum Kind {
        Number,
        Shared,
        Str,
        Inline,
        Bool,
        Error,
    }
    #[derive(Clone, Copy, PartialEq)]
    enum In {
        None,
        F,
        V,
        InlineT,
    }
    let mut out = SheetData::default();
    let mut layout = layout::LayoutReader::default();
    let mut tools = data_tools::ToolsReader::default();
    let (mut formula, mut value, mut inline) = (String::new(), String::new(), String::new());
    let mut at: Option<CellRef> = None;
    let mut kind = Kind::Number;
    let mut inside = In::None;
    let (mut row, mut next_col) = (0u32, 0u32);
    walk(xml, |ev| {
        match ev {
            XmlEvent::Start { name: "c", attrs } => {
                let cell = CellRef::parse(&attr(attrs, "r")).unwrap_or(CellRef::new(row, next_col));
                next_col = cell.col + 1;
                if let Ok(index) = attr(attrs, "s").parse::<usize>() {
                    if index > 0 {
                        out.styles.push((cell, index));
                    }
                }
                kind = match &*attr(attrs, "t") {
                    "s" => Kind::Shared,
                    "str" => Kind::Str,
                    "inlineStr" => Kind::Inline,
                    "b" => Kind::Bool,
                    "e" => Kind::Error,
                    _ => Kind::Number,
                };
                at = Some(cell);
                formula.clear();
                value.clear();
                inline.clear();
                return;
            }
            XmlEvent::Start { name: "v", .. } => inside = In::V,
            XmlEvent::Start { name: "f", .. } => inside = In::F,
            XmlEvent::Start { name: "t", .. } if kind == Kind::Inline => inside = In::InlineT,
            XmlEvent::End { name: "f" | "v" | "t" } => inside = In::None,
            XmlEvent::Text(t) => {
                match inside {
                    In::F => formula.push_str(t),
                    In::V => value.push_str(t),
                    In::InlineT => inline.push_str(t),
                    // Validation and conditional-format formulas.
                    In::None => tools.event(ev, dxfs),
                }
                return;
            }
            XmlEvent::End { name: "c" } => {
                let Some(cell) = at.take() else { return };
                let text = |s: &str| (text_cell_input(s), Some(Value::Text(s.to_string())));
                let loaded = if !formula.is_empty() {
                    Some((Some(format!("={}", strip_function_prefixes(&formula))), None))
                } else {
                    let v = value.trim();
                    match kind {
                        Kind::Shared => v
                            .parse::<usize>()
                            .ok()
                            .filter(|i| *i < shared.text.len())
                            .map(|i| (shared.input[i].clone(), Some(Value::Text(shared.text[i].clone())))),
                        Kind::Inline if inline.is_empty() => None,
                        Kind::Inline => Some(text(&inline)),
                        Kind::Str if value.is_empty() => None,
                        Kind::Str => Some(text(&value)),
                        Kind::Bool => Some((None, Some(Value::Bool(v == "1" || v.eq_ignore_ascii_case("true"))))),
                        Kind::Error => Some(
                            ErrorKind::parse(v).map_or((Some(v.to_string()), None), |e| (None, Some(Value::Error(e)))),
                        ),
                        Kind::Number if v.is_empty() => None,
                        Kind::Number => Some(
                            v.parse::<f64>().map_or((Some(v.to_string()), None), |n| (None, Some(Value::Number(n)))),
                        ),
                    }
                };
                if let Some((input, typed)) = loaded {
                    out.cells.push((cell, input, typed));
                }
                return;
            }
            XmlEvent::Start { name: "col", attrs } => {
                let (Ok(min), Ok(max), Ok(width)) = (
                    attr(attrs, "min").parse::<u32>(),
                    attr(attrs, "max").parse::<u32>(),
                    attr(attrs, "width").parse::<f64>(),
                ) else {
                    layout.event(ev);
                    tools.event(ev, dxfs);
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
            _ => {}
        }
        layout.event(ev);
        tools.event(ev, dxfs);
    });
    out.layout = layout.finish();
    out.tools = tools.finish();
    out
}

/// Maximum digit width of Calibri 11 (Excel's default font) in pixels.
const MAX_DIGIT_WIDTH: f64 = 7.0;

/// `<col width>` to pixels. The stored width already includes Excel's 5px
/// cell padding (ECMA-376 §18.3.1.13), so Excel's default 8.43-character
/// column is stored as 9.140625 and is 64px wide.
fn chars_to_px(width: f64) -> f64 {
    ((256.0 * width + (128.0 / MAX_DIGIT_WIDTH).trunc()) / 256.0 * MAX_DIGIT_WIDTH).trunc()
}

/// Pixels to `<col width>`, the inverse of `chars_to_px` for whole pixels.
fn px_to_chars(px: f64) -> f64 {
    (px / MAX_DIGIT_WIDTH * 256.0).trunc() / 256.0
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
        name: Option<String>,
        size: Option<u16>,
    }
    let mut num_fmts: HashMap<u32, String> = HashMap::new();
    let mut fonts: Vec<Font> = Vec::new();
    let mut fills: Vec<Option<String>> = Vec::new();
    let mut borders = BorderReader::default();
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
                "color" => font.color = argb(&attr(attrs, "rgb")),
                "name" => font.name = Some(attr(attrs, "val").to_string()),
                "sz" => font.size = attr(attrs, "val").parse::<f64>().ok().map(|pt| (pt * 10.0).round() as u16),
                _ => {}
            }
        }
        XmlEvent::Start { .. } | XmlEvent::End { .. } if borders.read(ev) => {}
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
            *fills.last_mut().unwrap() = argb(&attr(attrs, "rgb"));
        }
        XmlEvent::Start { name: "cellXfs", .. } => in_cell_xfs = true,
        XmlEvent::End { name: "cellXfs" } => in_cell_xfs = false,
        XmlEvent::Start { name: "xf", attrs } if in_cell_xfs => {
            let index = |key: &str| attr(attrs, key).parse::<usize>().unwrap_or(0);
            let font = fonts.get(index("fontId")).cloned().unwrap_or_default();
            let font_size = font.size.filter(|t| *t != sheets_engine::style::DEFAULT_FONT_SIZE);
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
                font: font.name.filter(|n| !n.is_empty() && n != sheets_engine::style::DEFAULT_FONT),
                font_size: font_size.map(|t| t.clamp(10, 4090)),
                wrap: false,
                valign: VAlign::Bottom,
                borders: borders.list.get(index("borderId")).cloned().unwrap_or_default(),
            });
        }
        XmlEvent::Start { name: "alignment", attrs } if in_cell_xfs => {
            if let Some(style) = out.last_mut() {
                style.align = match &*attr(attrs, "horizontal") {
                    "left" => Some(Align::Left),
                    "center" | "centerContinuous" => Some(Align::Center),
                    "right" => Some(Align::Right),
                    _ => None,
                };
                style.wrap = matches!(&*attr(attrs, "wrapText"), "1" | "true");
                style.valign = match &*attr(attrs, "vertical") {
                    "top" | "justify" => VAlign::Top,
                    "center" | "distributed" => VAlign::Center,
                    _ => VAlign::Bottom,
                };
            }
        }
        _ => {}
    });
    out
}

/// Reads the `<borders>` list of `styles.xml`, one `Borders` per `<border>`.
#[derive(Default)]
struct BorderReader {
    list: Vec<Borders>,
    in_border: bool,
    side: Option<Side>,
}

impl BorderReader {
    /// Consumes events inside `<border>`; returns false for everything else.
    fn read(&mut self, ev: &XmlEvent) -> bool {
        match ev {
            XmlEvent::Start { name: "border", .. } => {
                self.in_border = true;
                self.list.push(Borders::default());
            }
            XmlEvent::End { name: "border" } => self.in_border = false,
            _ if !self.in_border => return false,
            XmlEvent::Start { name, attrs } => {
                let side = match *name {
                    "left" | "start" => Some(Side::Left),
                    "right" | "end" => Some(Side::Right),
                    "top" => Some(Side::Top),
                    "bottom" => Some(Side::Bottom),
                    _ => None,
                };
                if let Some(side) = side {
                    self.side = Some(side);
                    let line = BorderStyle::from_name(&attr(attrs, "style")).map(|s| BorderLine::new(s, None));
                    *self.list.last_mut().unwrap().side_mut(side) = line;
                } else if let (Some(side), "color") = (self.side, *name) {
                    if let Some(line) = self.list.last_mut().unwrap().side_mut(side) {
                        line.color = argb(&attr(attrs, "rgb")).filter(|c| c != "#000000");
                    }
                }
            }
            XmlEvent::End { name: "left" | "start" | "right" | "end" | "top" | "bottom" } => self.side = None,
            _ => {}
        }
        true
    }
}

/// `<border>` element for `styles.xml`, sides in the schema's order.
fn border_xml(b: &Borders) -> String {
    let mut xml = String::from("<border>");
    for (tag, side) in [("left", Side::Left), ("right", Side::Right), ("top", Side::Top), ("bottom", Side::Bottom)] {
        match b.side(side) {
            None => xml.push_str(&format!("<{tag}/>")),
            Some(line) => {
                let color = match &line.color {
                    Some(c) => format!(r#"<color rgb="FF{}"/>"#, c.trim_start_matches('#')),
                    None => r#"<color indexed="64"/>"#.to_string(),
                };
                xml.push_str(&format!(r#"<{tag} style="{}">{color}</{tag}>"#, line.style.name()));
            }
        }
    }
    xml.push_str("<diagonal/></border>");
    xml
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

fn strip_function_prefixes(formula: &str) -> String {
    formula.replace("_xlfn.", "").replace("_xlws.", "")
}

/// Phase timings for `OFFICE_PROFILE=1` runs of the native benchmarks. A no-op in
/// WebAssembly, where `std::time::Instant` isn't available (it would panic).
struct Timer(#[cfg(not(target_arch = "wasm32"))] std::time::Instant);

impl Timer {
    fn start() -> Timer {
        Timer(
            #[cfg(not(target_arch = "wasm32"))]
            std::time::Instant::now(),
        )
    }

    /// Milliseconds since `start` or the previous lap, as text; starts the next lap.
    fn lap(&mut self) -> String {
        #[cfg(not(target_arch = "wasm32"))]
        {
            let now = std::time::Instant::now();
            let ms = (now - self.0).as_millis();
            self.0 = now;
            format!("{ms}ms")
        }
        #[cfg(target_arch = "wasm32")]
        String::new()
    }

    fn report(message: impl FnOnce() -> String) {
        #[cfg(not(target_arch = "wasm32"))]
        if std::env::var_os("OFFICE_PROFILE").is_some() {
            eprintln!("  {}", message());
        }
        #[cfg(target_arch = "wasm32")]
        let _ = message;
    }
}

// ---------------------------------------------------------------- writing

pub fn write(book: &Workbook) -> Result<Vec<u8>> {
    let mut parts: Vec<(String, String)> = Vec::new();
    let n = book.sheets.len();
    let mut styles = StyleTable::default();
    let mut drawing_parts = drawings::DrawingWriter::default();
    let mut strings = SharedStringTable::default();

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
        let sheet_xml = drawing_parts.add_sheet(id, book, i, worksheet_xml(&ws.sheet, &mut styles, &mut strings));
        let sheet_xml = page_setup::insert_into_worksheet(sheet_xml, &ws.page_setup);
        parts.push((format!("xl/worksheets/sheet{id}.xml"), sheet_xml));
    }
    // Ask Excel to recalculate on open, since our cached values may differ slightly.
    workbook.push_str(r#"</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>"#);
    rels.push_str(&format!(r#"<Relationship Id="rId{}" Type="{NS_REL}/styles" Target="styles.xml"/>"#, n + 1));
    if !strings.list.is_empty() {
        rels.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{NS_REL}/sharedStrings" Target="sharedStrings.xml"/>"#,
            n + 2
        ));
        parts.push(("xl/sharedStrings.xml".into(), strings.to_xml()));
        // Register the part in [Content_Types].xml (the first part).
        parts[0].1 = parts[0].1.replace(
            "</Types>",
            r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"#,
        );
    }
    rels.push_str("</Relationships>");
    let workbook = page_setup::insert_defined_names(workbook, book);
    parts.push(("xl/workbook.xml".into(), workbook));

    parts.push(("xl/_rels/workbook.xml.rels".into(), rels));
    parts.push(("xl/styles.xml".into(), styles.finish_xml()));
    let media = drawing_parts.finish(&mut parts);

    let mut t_zip = Timer::start();
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    // Level 3: about twice as fast as the default for files nearly as small.
    let options =
        SimpleFileOptions::default().compression_method(CompressionMethod::Deflated).compression_level(Some(3));
    for (name, xml) in parts {
        zip.start_file(name, options).map_err(|e| OfficeError::Io(e.to_string()))?;
        zip.write_all(xml.as_bytes()).map_err(|e| OfficeError::Io(e.to_string()))?;
    }
    for (name, bytes) in media {
        zip.start_file(name, SimpleFileOptions::default().compression_method(CompressionMethod::Stored))
            .map_err(|e| OfficeError::Io(e.to_string()))?;
        zip.write_all(&bytes).map_err(|e| OfficeError::Io(e.to_string()))?;
    }
    let bytes = zip.finish().map_err(|e| OfficeError::Io(e.to_string()))?.into_inner();
    let zip_time = t_zip.lap();
    Timer::report(|| format!("zip {zip_time}"));
    Ok(bytes)
}

/// (bold, italic, underline, strike, colour, name, size in tenths of a point).
type FontKey = (bool, bool, bool, bool, Option<String>, Option<String>, Option<u16>);

/// Collects the distinct styles used while writing sheets, then emits
/// `styles.xml`. Index 0 of every list is Excel's required default entry.
#[derive(Default)]
struct StyleTable {
    /// Style -> its index in `xfs` plus one, so lookups don't scan the list.
    lookup: rustc_hash::FxHashMap<Style, usize>,
    xfs: Vec<Style>,
    fonts: Vec<FontKey>,
    fills: Vec<String>,
    /// Non-empty borders; index 0 in the file is the empty border.
    borders: Vec<Borders>,
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
        if let Some(i) = self.lookup.get(style) {
            return *i;
        }
        self.xfs.push(style.clone());
        self.lookup.insert(style.clone(), self.xfs.len());
        self.xfs.len()
    }

    fn finish_xml(&mut self) -> String {
        let mut cell_xfs = String::from(r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#);
        for style in self.xfs.clone() {
            let font_key = (
                style.bold,
                style.italic,
                style.underline,
                style.strike,
                style.color.clone(),
                style.font.clone(),
                style.font_size,
            );
            let font_id = 1 + position_or_push(&mut self.fonts, font_key);
            let fill_id = style.fill.clone().map_or(0, |f| 2 + position_or_push(&mut self.fills, f));
            let fmt_id = self.format_id(&style.number_format);
            let border_id = if style.borders.is_empty() {
                0
            } else {
                1 + position_or_push(&mut self.borders, style.borders.clone())
            };
            let align = alignment_xml(&style);
            let apply_align = if align.is_empty() { "" } else { r#" applyAlignment="1""# };
            let apply_border = if border_id == 0 { "" } else { r#" applyBorder="1""# };
            let inner_close = if align.is_empty() { "/>".to_string() } else { format!(">{align}</xf>") };
            cell_xfs.push_str(&format!(
                r#"<xf numFmtId="{fmt_id}" fontId="{font_id}" fillId="{fill_id}" borderId="{border_id}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"{apply_border}{apply_align}{inner_close}"#
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
        for (bold, italic, underline, strike, color, name, size) in &self.fonts {
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
            let pt = size.unwrap_or(sheets_engine::style::DEFAULT_FONT_SIZE) as f64 / 10.0;
            xml.push_str(&format!(r#"<sz val="{pt}"/>"#));
            if let Some(c) = color {
                xml.push_str(&format!(r#"<color rgb="FF{}"/>"#, c.trim_start_matches('#')));
            }
            let name = escape(name.as_deref().unwrap_or(sheets_engine::style::DEFAULT_FONT));
            xml.push_str(&format!(r#"<name val="{name}"/><family val="2"/></font>"#));
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
        xml.push_str(&format!(r#"<borders count="{}">"#, self.borders.len() + 1));
        xml.push_str(&border_xml(&Borders::default()));
        for b in &self.borders {
            xml.push_str(&border_xml(b));
        }
        xml.push_str("</borders>");
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

/// `<alignment>` for an `<xf>`, or "" when everything is default.
fn alignment_xml(style: &Style) -> String {
    let mut attrs = String::new();
    match style.align {
        Some(Align::Left) => attrs.push_str(r#" horizontal="left""#),
        Some(Align::Center) => attrs.push_str(r#" horizontal="center""#),
        Some(Align::Right) => attrs.push_str(r#" horizontal="right""#),
        None => {}
    }
    match style.valign {
        VAlign::Top => attrs.push_str(r#" vertical="top""#),
        VAlign::Center => attrs.push_str(r#" vertical="center""#),
        VAlign::Bottom => {}
    }
    if style.wrap {
        attrs.push_str(r#" wrapText="1""#);
    }
    if attrs.is_empty() {
        String::new()
    } else {
        format!("<alignment{attrs}/>")
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

fn worksheet_xml<'a>(sheet: &'a Sheet, styles: &mut StyleTable, strings: &mut SharedStringTable<'a>) -> String {
    let positions = {
        // Cells with content plus formatted empty cells, row by row.
        let mut p: Vec<CellRef> = sheet.cell_positions();
        p.extend(sheet.styled_cells().into_iter().map(|(at, _)| at));
        p.sort_unstable();
        p.dedup();
        p
    };
    // About 40 bytes of XML per cell; reserving avoids repeated copying of a huge string.
    let mut xml = String::with_capacity(positions.len() * 40 + 1024);
    xml.push_str(&format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS_MAIN}">"#));
    xml.push_str(&layout::sheet_views_xml(sheet));
    xml.push_str(&layout::sheet_format_xml(sheet));
    xml.push_str(&layout::cols_xml(sheet));
    xml.push_str("<sheetData>");
    let mut rows = layout::RowWriter::new(sheet);
    let mut current_row = None;
    for at in positions {
        if current_row != Some(at.row) {
            if current_row.is_some() {
                xml.push_str("</row>");
            }
            rows.before(at.row, &mut xml);
            xml.push_str(&rows.open(at.row));
            current_row = Some(at.row);
        }
        let s = styles.index(sheet.style(at));
        write_cell(&mut xml, at, sheet.formula_input(at), sheet.value(at), s, strings);
    }
    if current_row.is_some() {
        xml.push_str("</row>");
    }
    rows.rest(&mut xml);
    xml.push_str("</sheetData>");
    // Excel's order: autoFilter, mergeCells, conditionalFormatting, dataValidations.
    xml.push_str(&data_tools::autofilter_xml(sheet));
    xml.push_str(&layout::merge_cells_xml(sheet));
    xml.push_str(&data_tools::cf_and_validation_xml(sheet, styles));
    xml.push_str("</worksheet>");
    xml
}

/// Append one `<c>` element. Hot path when saving big sheets: writes straight
/// into `xml` without building intermediate strings.
fn write_cell<'a>(
    xml: &mut String,
    at: CellRef,
    formula: Option<&str>,
    value: &'a Value,
    style: usize,
    strings: &mut SharedStringTable<'a>,
) {
    use std::fmt::Write as _;
    if formula.is_none() && value.is_empty() && style == 0 {
        return;
    }
    xml.push_str("<c r=\"");
    push_cell_ref(xml, at);
    xml.push('"');
    if style > 0 {
        let _ = write!(xml, " s=\"{style}\"");
    }
    if let Some(formula) = formula.and_then(|f| f.strip_prefix('=')) {
        xml.push_str(match value {
            Value::Text(_) => " t=\"str\"><f>",
            Value::Bool(_) => " t=\"b\"><f>",
            Value::Error(_) => " t=\"e\"><f>",
            _ => "><f>",
        });
        escape_into(xml, &add_function_prefixes(formula));
        xml.push_str("</f>");
        match value {
            Value::Number(n) => {
                let _ = write!(xml, "<v>{n}</v>");
            }
            Value::Text(t) => {
                xml.push_str("<v>");
                escape_into(xml, t);
                xml.push_str("</v>");
            }
            Value::Bool(b) => {
                let _ = write!(xml, "<v>{}</v>", *b as u8);
            }
            Value::Error(e) => {
                let _ = write!(xml, "<v>{}</v>", excel_error(*e));
            }
            Value::Empty => {}
        }
        xml.push_str("</c>");
        return;
    }
    match value {
        Value::Number(n) => {
            let _ = write!(xml, "><v>{n}</v></c>");
        }
        Value::Text(t) => {
            let _ = write!(xml, " t=\"s\"><v>{}</v></c>", strings.id(t));
        }
        Value::Bool(b) => {
            let _ = write!(xml, " t=\"b\"><v>{}</v></c>", *b as u8);
        }
        Value::Error(e) => {
            let _ = write!(xml, " t=\"e\"><v>{}</v></c>", excel_error(*e));
        }
        Value::Empty => xml.push_str("/>"),
    }
}

/// "AB12" without allocating.
fn push_cell_ref(xml: &mut String, at: CellRef) {
    use std::fmt::Write as _;
    let mut letters = [0u8; 3];
    let mut n = at.col + 1;
    let mut i = 3;
    while n > 0 {
        i -= 1;
        letters[i] = b'A' + ((n - 1) % 26) as u8;
        n = (n - 1) / 26;
    }
    xml.push_str(std::str::from_utf8(&letters[i..]).unwrap_or("A"));
    let _ = write!(xml, "{}", at.row + 1);
}

/// XML-escape `s` onto the end of `xml`.
fn escape_into(xml: &mut String, s: &str) {
    if !s.contains(['&', '<', '>', '"']) {
        xml.push_str(s);
        return;
    }
    for ch in s.chars() {
        match ch {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            c => xml.push(c),
        }
    }
}

/// Text values written once to `xl/sharedStrings.xml` and referenced by index,
/// as Excel does: smaller files, and faster to open.
#[derive(Default)]
struct SharedStringTable<'a> {
    index: rustc_hash::FxHashMap<&'a str, usize>,
    list: Vec<&'a str>,
    uses: usize,
}

impl<'a> SharedStringTable<'a> {
    fn id(&mut self, s: &'a str) -> usize {
        self.uses += 1;
        if let Some(i) = self.index.get(s) {
            return *i;
        }
        self.list.push(s);
        self.index.insert(s, self.list.len() - 1);
        self.list.len() - 1
    }

    fn to_xml(&self) -> String {
        let mut xml = String::with_capacity(self.list.iter().map(|s| s.len() + 40).sum::<usize>() + 256);
        xml.push_str(&format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{NS_MAIN}" count="{}" uniqueCount="{}">"#,
            self.uses,
            self.list.len()
        ));
        for s in &self.list {
            xml.push_str(r#"<si><t xml:space="preserve">"#);
            escape_into(&mut xml, s);
            xml.push_str("</t></si>");
        }
        xml.push_str("</sst>");
        xml
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
            ..Style::default()
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

    #[test]
    fn round_trips_layout() {
        use sheets_engine::layout::Freeze;
        use sheets_engine::Range;
        let mut book = Workbook::new();
        let s = &mut book.sheets[0].sheet;
        s.set_input(CellRef::new(3, 0), "tall");
        s.set_col_width(2, Some(150.0));
        let l = s.layout_mut();
        l.set_row_height(3, 3, Some(40.0));
        l.set_row_height(7, 8, Some(60.0));
        l.hidden_rows.set(5, 5, Some(()));
        l.hidden_cols.set(2, 4, Some(()));
        l.set_merges([Range::new(CellRef::new(0, 0), CellRef::new(1, 2))]);
        l.freeze = Some(Freeze { rows: 1, cols: 2 });
        let bytes = write(&book).unwrap();
        let sheet_xml = {
            let mut zip = ZipArchive::new(Cursor::new(&bytes[..])).unwrap();
            part(&mut zip, "xl/worksheets/sheet1.xml").unwrap().unwrap()
        };
        assert!(sheet_xml
            .contains(r#"<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>"#));
        assert!(sheet_xml.contains(r#"<row r="4" ht="30" customHeight="1">"#));
        assert!(sheet_xml.contains(r#"<row r="6" hidden="1"/>"#));
        assert!(sheet_xml.contains(r#"<mergeCells count="1"><mergeCell ref="A1:C2"/></mergeCells>"#));
        let custom = px_to_chars(150.0);
        let default = px_to_chars(sheets_engine::sheet::DEFAULT_COL_WIDTH);
        assert!(
            sheet_xml.contains(&format!(r#"<col min="3" max="3" width="{custom}" hidden="1" customWidth="1"/>"#)),
            "{sheet_xml}"
        );
        assert!(sheet_xml.contains(&format!(r#"<col min="4" max="5" width="{default}" hidden="1"/>"#)), "{sheet_xml}");

        let back = read(&bytes).unwrap();
        let s = &back.sheets[0].sheet;
        assert_eq!(s.layout(), book.sheets[0].sheet.layout());
        assert_eq!(s.col_width(2), 150.0);
        assert_eq!(s.input(CellRef::new(3, 0)), "tall");
    }

    #[test]
    fn hidden_tail_uses_zero_height() {
        let mut book = Workbook::new();
        let l = book.sheets[0].sheet.layout_mut();
        l.hidden_rows.set(2, 2, Some(()));
        l.hidden_rows.set(10, sheets_engine::MAX_ROWS - 1, Some(()));
        let bytes = write(&book).unwrap();
        assert!(bytes.len() < 10_000, "the hidden tail isn't written row by row");
        let back = read(&bytes).unwrap();
        assert_eq!(back.sheets[0].sheet.layout(), book.sheets[0].sheet.layout());
    }

    #[test]
    fn round_trips_fonts_borders_and_alignment() {
        let mut book = Workbook::new();
        let s = &mut book.sheets[0].sheet;
        let red_thick = BorderLine::new(BorderStyle::Thick, Some("#FF0000".into()));
        let styled = Style {
            font: Some("Times New Roman".into()),
            font_size: Some(105),
            wrap: true,
            valign: VAlign::Top,
            align: Some(Align::Center),
            borders: Borders {
                top: Some(BorderLine::new(BorderStyle::Thin, None)),
                right: Some(BorderLine::new(BorderStyle::Dashed, None)),
                bottom: Some(BorderLine::new(BorderStyle::Double, Some("#0070C0".into()))),
                left: Some(red_thick.clone()),
            },
            ..Style::default()
        };
        s.set_style(CellRef::new(0, 0), styled.clone());
        s.set_input(CellRef::new(0, 0), "line one\nline two");
        let mangal =
            Style { font: Some("Mangal".into()), font_size: Some(240), valign: VAlign::Center, ..Style::default() };
        s.set_style(CellRef::new(1, 0), mangal.clone());
        s.set_input(CellRef::new(1, 0), "नमस्ते");
        let dotted = Style {
            borders: Borders { bottom: Some(BorderLine::new(BorderStyle::Dotted, None)), ..Borders::default() },
            ..Style::default()
        };
        s.set_style(CellRef::new(2, 2), dotted.clone());

        let bytes = write(&book).unwrap();
        let back = read(&bytes).unwrap();
        let s = &back.sheets[0].sheet;
        assert_eq!(s.style(CellRef::new(0, 0)), &styled);
        assert_eq!(s.input(CellRef::new(0, 0)), "line one\nline two");
        assert_eq!(s.style(CellRef::new(1, 0)), &mangal);
        assert_eq!(s.style(CellRef::new(2, 2)), &dotted);
        assert!(s.style(CellRef::new(3, 3)).is_default());

        // The XML has Excel's structure.
        let mut zip = ZipArchive::new(Cursor::new(bytes.as_slice())).unwrap();
        let styles = part(&mut zip, "xl/styles.xml").unwrap().unwrap();
        assert!(styles.contains(r#"<sz val="10.5"/>"#), "{styles}");
        assert!(styles.contains(r#"<name val="Times New Roman"/>"#));
        assert!(styles.contains(r#"<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>"#));
        assert!(styles.contains(r#"<left style="thick"><color rgb="FFFF0000"/></left>"#));
        assert!(styles.contains(r#"<top style="thin"><color indexed="64"/></top>"#));
        assert!(styles.contains(r#"<alignment horizontal="center" vertical="top" wrapText="1"/>"#));
        assert!(styles.contains(r#"applyBorder="1""#));
    }

    #[test]
    fn reads_excel_styles_xml() {
        // Trimmed from a workbook saved by Excel for Microsoft 365.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="14"/><color rgb="FFC00000"/><name val="Arial"/><family val="2"/></font>
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color indexed="64"/></left><right style="medium"><color rgb="FF00B050"/></right><top/><bottom style="double"><color auto="1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf></cellXfs>
<dxfs count="1"><dxf><font><b/></font><border><left style="thick"/></border></dxf></dxfs>
</styleSheet>"#;
        let styles = read_styles(xml);
        assert_eq!(styles.len(), 4);
        assert!(styles[0].is_default());
        let s = &styles[1];
        assert!(s.bold);
        assert_eq!(s.font.as_deref(), Some("Arial"));
        assert_eq!(s.font_size, Some(140));
        assert_eq!(s.color.as_deref(), Some("#C00000"));
        assert_eq!(s.borders.left, Some(BorderLine::new(BorderStyle::Thin, None)));
        assert_eq!(s.borders.right, Some(BorderLine::new(BorderStyle::Medium, Some("#00B050".into()))));
        assert_eq!(s.borders.top, None);
        assert_eq!(s.borders.bottom, Some(BorderLine::new(BorderStyle::Double, None)));
        assert!(styles[2].wrap);
        assert_eq!(styles[2].valign, VAlign::Center);
        assert_eq!(styles[2].font, None);
        assert_eq!(styles[3].valign, VAlign::Top);
        assert_eq!(styles[3].align, Some(Align::Left));
        assert!(!styles[3].wrap);
    }

    #[test]
    fn column_widths_match_excel() {
        // What Excel writes for its default column (8.43 characters).
        assert_eq!(chars_to_px(9.140625), 64.0);
        assert_eq!(px_to_chars(64.0), 9.140625);
        // A 20-character column as Excel stores it.
        assert_eq!(chars_to_px(20.7109375), 145.0);
        for px in 8..=400 {
            assert_eq!(chars_to_px(px_to_chars(px as f64)), px as f64, "{px}px");
        }
    }
}
