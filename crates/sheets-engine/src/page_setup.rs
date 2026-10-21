//! Page setup of a worksheet: how it prints.
//!
//! Mirrors Excel's Page Layout settings — paper size, orientation, margins,
//! scaling, print area, rows/columns to repeat on every page (print titles),
//! gridlines and headings, and header/footer text. Stored per sheet on
//! [`Worksheet`](crate::Worksheet) and saved in `.xlsx` as `<pageSetup>`,
//! `<pageMargins>`, `<printOptions>`, `<headerFooter>` and the
//! `_xlnm.Print_Area` / `_xlnm.Print_Titles` defined names.
//!
//! The UI changes settings with [`PageSetup::set`] (string keys and values, so
//! the WebAssembly binding stays thin) and applies them with the undoable
//! [`SetPageSetup`] op.

use office_core::Op;

use crate::cell::{col_name, parse_col, CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::workbook::Workbook;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Orientation {
    #[default]
    Portrait,
    Landscape,
}

/// Paper size codes used by `.xlsx` (`<pageSetup paperSize="9">`).
pub mod paper {
    pub const LETTER: u32 = 1;
    pub const TABLOID: u32 = 3;
    pub const LEGAL: u32 = 5;
    pub const EXECUTIVE: u32 = 7;
    pub const A3: u32 = 8;
    pub const A4: u32 = 9;
    pub const A5: u32 = 11;

    /// Width and height in inches (portrait), for the sizes the UI offers.
    pub fn size_inches(code: u32) -> Option<(f64, f64)> {
        Some(match code {
            LETTER => (8.5, 11.0),
            TABLOID => (11.0, 17.0),
            LEGAL => (8.5, 14.0),
            EXECUTIVE => (7.25, 10.5),
            A3 => (297.0 / 25.4, 420.0 / 25.4),
            A4 => (210.0 / 25.4, 297.0 / 25.4),
            A5 => (148.0 / 25.4, 210.0 / 25.4),
            _ => return None,
        })
    }
}

/// Page margins in inches, as Excel stores them.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Margins {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

impl Margins {
    /// Excel's "Normal" preset (the default for new sheets).
    pub const NORMAL: Margins = Margins { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
    pub const WIDE: Margins = Margins { left: 1.0, right: 1.0, top: 1.0, bottom: 1.0, header: 0.5, footer: 0.5 };
    pub const NARROW: Margins = Margins { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

    /// "normal", "wide", "narrow" or "custom".
    pub fn preset_name(&self) -> &'static str {
        if *self == Self::NORMAL {
            "normal"
        } else if *self == Self::WIDE {
            "wide"
        } else if *self == Self::NARROW {
            "narrow"
        } else {
            "custom"
        }
    }
}

impl Default for Margins {
    fn default() -> Self {
        Self::NORMAL
    }
}

/// How the printout is scaled.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Scaling {
    /// Adjust to a percentage of normal size (Excel allows 10–400).
    Percent(u16),
    /// Fit to this many pages wide by tall; 0 means "automatic" (as many as needed).
    FitTo { width: u16, height: u16 },
}

impl Default for Scaling {
    fn default() -> Self {
        Scaling::Percent(100)
    }
}

/// Inclusive span of rows or columns, zero-based.
pub type Span = (u32, u32);

#[derive(Debug, Clone, PartialEq)]
pub struct PageSetup {
    pub orientation: Orientation,
    /// `.xlsx` paper size code, see [`paper`].
    pub paper_size: u32,
    pub scaling: Scaling,
    pub margins: Margins,
    pub center_horizontally: bool,
    pub center_vertically: bool,
    /// Print cell gridlines.
    pub gridlines: bool,
    /// Print row numbers and column letters.
    pub headings: bool,
    /// Only this range prints (Page Layout > Print Area).
    pub print_area: Option<Range>,
    /// Rows repeated at the top of every page.
    pub title_rows: Option<Span>,
    /// Columns repeated at the left of every page.
    pub title_cols: Option<Span>,
    /// Header and footer in Excel's codes (`&L`, `&C`, `&R` sections, `&P` page, `&N` pages...).
    pub header: String,
    pub footer: String,
}

impl Default for PageSetup {
    fn default() -> Self {
        PageSetup {
            orientation: Orientation::Portrait,
            paper_size: paper::A4,
            scaling: Scaling::default(),
            margins: Margins::NORMAL,
            center_horizontally: false,
            center_vertically: false,
            gridlines: false,
            headings: false,
            print_area: None,
            title_rows: None,
            title_cols: None,
            header: String::new(),
            footer: String::new(),
        }
    }
}

/// A rejected [`PageSetup::set`] call.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SetupError(pub String);

impl std::fmt::Display for SetupError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl PageSetup {
    pub fn is_default(&self) -> bool {
        *self == PageSetup::default()
    }

    /// Change one setting. Keys and values (all strings):
    ///
    /// | key | value |
    /// |---|---|
    /// | `orientation` | `portrait` / `landscape` |
    /// | `paper` | `.xlsx` paper code, e.g. `9` (A4) |
    /// | `scale` | percent, 10–400 (turns off fit-to) |
    /// | `fit` | `W,H` pages (0 = automatic), e.g. `1,1` or `1,0` |
    /// | `margins` | `normal` / `wide` / `narrow`, or six inches `left,right,top,bottom,header,footer` |
    /// | `centerH`, `centerV`, `gridlines`, `headings` | `1` / `0` |
    /// | `printArea` | `A1:D20`, `A:D`, `1:20` or empty to clear |
    /// | `titleRows` | `1:2` (one-based, `$` allowed) or empty |
    /// | `titleCols` | `A:B` or empty |
    /// | `header`, `footer` | Excel header/footer codes |
    pub fn set(&mut self, key: &str, value: &str) -> Result<(), SetupError> {
        let bad = || SetupError(format!("invalid value for {key}: {value:?}"));
        let flag = || match value {
            "1" | "true" => Ok(true),
            "0" | "false" | "" => Ok(false),
            _ => Err(bad()),
        };
        match key {
            "orientation" => {
                self.orientation = match value {
                    "portrait" => Orientation::Portrait,
                    "landscape" => Orientation::Landscape,
                    _ => return Err(bad()),
                }
            }
            "paper" => self.paper_size = value.parse().ok().filter(|n| *n > 0).ok_or_else(bad)?,
            "scale" => {
                let pct: u16 = value.trim().parse().map_err(|_| bad())?;
                if !(10..=400).contains(&pct) {
                    return Err(SetupError("Scale must be between 10% and 400%".into()));
                }
                self.scaling = Scaling::Percent(pct);
            }
            "fit" => {
                let (w, h) = value.split_once(',').ok_or_else(bad)?;
                let width: u16 = w.trim().parse().map_err(|_| bad())?;
                let height: u16 = h.trim().parse().map_err(|_| bad())?;
                if width > 32767 || height > 32767 {
                    return Err(bad());
                }
                self.scaling = Scaling::FitTo { width, height };
            }
            "margins" => {
                self.margins = match value {
                    "normal" => Margins::NORMAL,
                    "wide" => Margins::WIDE,
                    "narrow" => Margins::NARROW,
                    custom => {
                        let n: Vec<f64> = custom
                            .split(',')
                            .map(|p| p.trim().parse::<f64>())
                            .collect::<Result<_, _>>()
                            .map_err(|_| bad())?;
                        if n.len() != 6 || n.iter().any(|m| !m.is_finite() || *m < 0.0 || *m > 20.0) {
                            return Err(bad());
                        }
                        Margins { left: n[0], right: n[1], top: n[2], bottom: n[3], header: n[4], footer: n[5] }
                    }
                }
            }
            "centerH" => self.center_horizontally = flag()?,
            "centerV" => self.center_vertically = flag()?,
            "gridlines" => self.gridlines = flag()?,
            "headings" => self.headings = flag()?,
            "printArea" => {
                self.print_area = if value.trim().is_empty() { None } else { Some(parse_area(value).ok_or_else(bad)?) }
            }
            "titleRows" => {
                self.title_rows =
                    if value.trim().is_empty() { None } else { Some(parse_row_span(value).ok_or_else(bad)?) }
            }
            "titleCols" => {
                self.title_cols =
                    if value.trim().is_empty() { None } else { Some(parse_col_span(value).ok_or_else(bad)?) }
            }
            "header" => self.header = value.to_string(),
            "footer" => self.footer = value.to_string(),
            _ => return Err(SetupError(format!("unknown page setup key {key:?}"))),
        }
        Ok(())
    }

    /// Compact JSON for the UI (see `PageSetup` in `@office/engine`).
    pub fn to_json(&self) -> String {
        let m = &self.margins;
        let (scale, fit) = match self.scaling {
            Scaling::Percent(p) => (p, "null".to_string()),
            Scaling::FitTo { width, height } => (100, format!("[{width},{height}]")),
        };
        let opt = |s: Option<String>| s.map_or("null".to_string(), |s| json_string(&s));
        format!(
            concat!(
                r#"{{"orientation":"{}","paper":{},"scale":{},"fit":{},"#,
                r#""margins":{{"preset":"{}","left":{},"right":{},"top":{},"bottom":{},"header":{},"footer":{}}},"#,
                r#""centerH":{},"centerV":{},"gridlines":{},"headings":{},"#,
                r#""printArea":{},"titleRows":{},"titleCols":{},"header":{},"footer":{}}}"#
            ),
            match self.orientation {
                Orientation::Portrait => "portrait",
                Orientation::Landscape => "landscape",
            },
            self.paper_size,
            scale,
            fit,
            m.preset_name(),
            m.left,
            m.right,
            m.top,
            m.bottom,
            m.header,
            m.footer,
            self.center_horizontally,
            self.center_vertically,
            self.gridlines,
            self.headings,
            opt(self.print_area.map(|r| format_area(r, false))),
            opt(self.title_rows.map(|s| format_row_span(s, false))),
            opt(self.title_cols.map(|s| format_col_span(s, false))),
            json_string(&self.header),
            json_string(&self.footer),
        )
    }
}

fn json_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

// ---------------------------------------------------------------- references

fn strip_dollars(s: &str) -> String {
    s.trim().chars().filter(|&c| c != '$').collect()
}

/// `1:3` -> rows 0..=2 (one-based in text, `$` allowed). A single `5` is row 5.
pub fn parse_row_span(s: &str) -> Option<Span> {
    let s = strip_dollars(s);
    let (a, b) = s.split_once(':').unwrap_or((&s, &s));
    let (a, b): (u32, u32) = (a.parse().ok()?, b.parse().ok()?);
    if a == 0 || b == 0 || a > MAX_ROWS || b > MAX_ROWS {
        return None;
    }
    Some((a.min(b) - 1, a.max(b) - 1))
}

/// `A:B` -> columns 0..=1. A single `C` is column C.
pub fn parse_col_span(s: &str) -> Option<Span> {
    let s = strip_dollars(s);
    let (a, b) = s.split_once(':').unwrap_or((&s, &s));
    let (a, b) = (parse_col(a)?, parse_col(b)?);
    Some((a.min(b), a.max(b)))
}

/// A print area: `A1:D20`, `B3`, whole columns `A:D` or whole rows `1:20`.
pub fn parse_area(s: &str) -> Option<Range> {
    let t = strip_dollars(s);
    if let Some((a, b)) = t.split_once(':') {
        if let (Some(a), Some(b)) = (CellRef::parse(a), CellRef::parse(b)) {
            return Some(Range::new(a, b));
        }
    } else if let Some(c) = CellRef::parse(&t) {
        return Some(Range::single(c));
    }
    if let Some((c0, c1)) = parse_col_span(&t).filter(|_| t.contains(':')) {
        return Some(Range::new(CellRef::new(0, c0), CellRef::new(MAX_ROWS - 1, c1)));
    }
    let (r0, r1) = parse_row_span(&t).filter(|_| t.contains(':'))?;
    Some(Range::new(CellRef::new(r0, 0), CellRef::new(r1, MAX_COLS - 1)))
}

/// `1:2`, or `$1:$2` with `absolute`.
pub fn format_row_span((a, b): Span, absolute: bool) -> String {
    let d = if absolute { "$" } else { "" };
    format!("{d}{}:{d}{}", a + 1, b + 1)
}

/// `A:B`, or `$A:$B` with `absolute`.
pub fn format_col_span((a, b): Span, absolute: bool) -> String {
    let d = if absolute { "$" } else { "" };
    format!("{d}{}:{d}{}", col_name(a), col_name(b))
}

/// `A1:D20` (or `$A$1:$D$20`); whole columns/rows use the `A:D` / `1:20` forms.
pub fn format_area(r: Range, absolute: bool) -> String {
    let all_rows = r.start.row == 0 && r.end.row == MAX_ROWS - 1;
    let all_cols = r.start.col == 0 && r.end.col == MAX_COLS - 1;
    if all_rows && !all_cols {
        return format_col_span((r.start.col, r.end.col), absolute);
    }
    if all_cols && !all_rows {
        return format_row_span((r.start.row, r.end.row), absolute);
    }
    let d = if absolute { "$" } else { "" };
    let cell = |c: CellRef| format!("{d}{}{d}{}", col_name(c.col), c.row + 1);
    if r.start == r.end {
        return cell(r.start);
    }
    format!("{}:{}", cell(r.start), cell(r.end))
}

// ---------------------------------------------------------------- op

/// Replace a sheet's page setup (one undo step).
#[derive(Debug, Clone)]
pub struct SetPageSetup {
    pub sheet: usize,
    pub setup: PageSetup,
}

impl Op<Workbook> for SetPageSetup {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            ws.page_setup = self.setup.clone();
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let setup = doc.sheets.get(self.sheet).map(|ws| ws.page_setup.clone()).unwrap_or_default();
        Box::new(SetPageSetup { sheet: self.sheet, setup })
    }
}

#[cfg(test)]
mod tests {
    use office_core::history::History;

    use super::*;

    #[test]
    fn sets_values_and_rejects_bad_ones() {
        let mut p = PageSetup::default();
        p.set("orientation", "landscape").unwrap();
        p.set("paper", "5").unwrap();
        p.set("fit", "1,0").unwrap();
        p.set("margins", "narrow").unwrap();
        p.set("gridlines", "1").unwrap();
        p.set("printArea", "$B$2:$D$9").unwrap();
        p.set("titleRows", "$1:$2").unwrap();
        p.set("titleCols", "a").unwrap();
        assert_eq!(p.orientation, Orientation::Landscape);
        assert_eq!(p.paper_size, paper::LEGAL);
        assert_eq!(p.scaling, Scaling::FitTo { width: 1, height: 0 });
        assert_eq!(p.margins, Margins::NARROW);
        assert!(p.gridlines);
        assert_eq!(p.print_area, Some(Range::new(CellRef::new(1, 1), CellRef::new(8, 3))));
        assert_eq!(p.title_rows, Some((0, 1)));
        assert_eq!(p.title_cols, Some((0, 0)));

        assert!(p.set("scale", "5").is_err());
        assert!(p.set("orientation", "sideways").is_err());
        assert!(p.set("printArea", "nonsense").is_err());
        assert!(p.set("margins", "1,2,3").is_err());
        assert!(p.set("nope", "1").is_err());
        p.set("scale", "80").unwrap();
        assert_eq!(p.scaling, Scaling::Percent(80));
        p.set("printArea", "").unwrap();
        assert_eq!(p.print_area, None);
        p.set("margins", "1,1,0.5,0.5,0.2,0.2").unwrap();
        assert_eq!(p.margins.preset_name(), "custom");
    }

    #[test]
    fn parses_and_formats_areas() {
        let cols = parse_area("$A:$D").unwrap();
        assert_eq!(cols, Range::new(CellRef::new(0, 0), CellRef::new(MAX_ROWS - 1, 3)));
        assert_eq!(format_area(cols, true), "$A:$D");
        let rows = parse_area("3:5").unwrap();
        assert_eq!(format_area(rows, false), "3:5");
        let block = parse_area("d9:b2").unwrap();
        assert_eq!(format_area(block, true), "$B$2:$D$9");
        assert_eq!(format_area(parse_area("C3").unwrap(), false), "C3");
        assert_eq!(parse_row_span("$4:$2"), Some((1, 3)));
        assert_eq!(parse_row_span("0:1"), None);
        assert_eq!(format_col_span((0, 27), true), "$A:$AB");
    }

    #[test]
    fn json_has_every_field() {
        let mut p = PageSetup::default();
        p.set("footer", "&CPage &P of &N \"x\"").unwrap();
        p.set("titleRows", "1").unwrap();
        let json = p.to_json();
        assert!(json.starts_with(
            r#"{"orientation":"portrait","paper":9,"scale":100,"fit":null,"margins":{"preset":"normal","left":0.7"#
        ));
        assert!(json.contains(r#""titleRows":"1:1""#));
        assert!(json.contains(r#""footer":"&CPage &P of &N \"x\"""#));
        assert!(json.contains(r#""printArea":null"#));
    }

    #[test]
    fn set_page_setup_is_undoable() {
        let mut wb = Workbook::new();
        let mut history = History::default();
        let mut setup = PageSetup::default();
        setup.set("orientation", "landscape").unwrap();
        history.apply(&mut wb, Box::new(SetPageSetup { sheet: 0, setup: setup.clone() }));
        assert_eq!(wb.sheets[0].page_setup, setup);
        assert!(history.undo(&mut wb));
        assert!(wb.sheets[0].page_setup.is_default());
        assert!(history.redo(&mut wb));
        assert_eq!(wb.sheets[0].page_setup.orientation, Orientation::Landscape);
    }

    #[test]
    fn paper_sizes() {
        let (w, h) = paper::size_inches(paper::A4).unwrap();
        assert!((w - 8.27).abs() < 0.01 && (h - 11.69).abs() < 0.01);
        assert_eq!(paper::size_inches(999), None);
    }
}
