//! Charts and pictures that float over a sheet ("drawings").
//!
//! A drawing is anchored to cells the way Excel's default "move and size with
//! cells" objects are: its top-left and bottom-right corners are each a cell
//! plus a pixel offset inside that cell ([`Anchor`]). The UI turns anchors into
//! pixels with its column/row layout, so a drawing follows column resizes and
//! row/column inserts.
//!
//! A [`Chart`] stores *references* to its data (`Sheet1!$B$2:$B$5`), never the
//! numbers themselves. [`chart_view`] reads the current cell values every time
//! the chart is drawn, so charts update live as data changes.
//!
//! Every change to a sheet's drawings goes through the [`SetDrawings`] op, so
//! inserting, moving, resizing, formatting and deleting are all undoable.

use std::fmt;
use std::sync::Arc;

use office_core::Op;

use crate::cell::{col_name, CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::formula::refs::Axis;
use crate::sheet::Sheet;
use crate::style::NumberFormat;
use crate::value::{format_number, Value};
use crate::workbook::Workbook;

/// Excel's Office theme accent colours (Office 2013-2022 palette), used for
/// series in order. After six, Excel repeats them darker, then lighter.
pub const PALETTE: [&str; 6] = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];
const PALETTE_DARK: [&str; 6] = ["#264478", "#9E480E", "#636363", "#997300", "#255E91", "#43682B"];
const PALETTE_LIGHT: [&str; 6] = ["#698ED0", "#F1975A", "#B7B7B7", "#FFCD33", "#7CAFDD", "#8CC168"];

/// Automatic colour for the n-th series (or pie slice).
pub fn auto_color(n: usize) -> &'static str {
    let list = match (n / 6) % 3 {
        0 => &PALETTE,
        1 => &PALETTE_DARK,
        _ => &PALETTE_LIGHT,
    };
    list[n % 6]
}

/// English Metric Units per CSS pixel (at 96 dpi), as used in DrawingML.
pub const EMU_PER_PX: f64 = 9525.0;

/// Default size of a new chart in pixels (Excel: 5" x 3" = 480 x 288).
pub const DEFAULT_CHART_SIZE: (f64, f64) = (480.0, 288.0);

// ---------------------------------------------------------------- anchors

/// A corner of a drawing: a cell and an offset in pixels from its top-left.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AnchorPoint {
    pub row: u32,
    pub col: u32,
    pub row_off: f64,
    pub col_off: f64,
}

impl AnchorPoint {
    pub fn new(row: u32, col: u32, row_off: f64, col_off: f64) -> Self {
        AnchorPoint { row, col, row_off: row_off.max(0.0), col_off: col_off.max(0.0) }
    }
}

/// Top-left (`from`) and bottom-right (`to`) corners.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Anchor {
    pub from: AnchorPoint,
    pub to: AnchorPoint,
}

impl Anchor {
    /// An anchor in a uniform grid (every column `col_w`, every row `row_h`
    /// pixels wide/high). Used for defaults and tests; the UI computes real
    /// anchors from its own layout.
    pub fn from_pixels_uniform(x: f64, y: f64, w: f64, h: f64, col_w: f64, row_h: f64) -> Anchor {
        let point = |x: f64, y: f64| {
            let col = ((x / col_w).floor() as u32).min(MAX_COLS - 1);
            let row = ((y / row_h).floor() as u32).min(MAX_ROWS - 1);
            AnchorPoint::new(row, col, y - row as f64 * row_h, x - col as f64 * col_w)
        };
        Anchor { from: point(x, y), to: point(x + w, y + h) }
    }

    fn remapped(&self, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Anchor {
        let limit = if axis == Axis::Row { MAX_ROWS } else { MAX_COLS };
        // A corner in a deleted line moves to the start of the next surviving one.
        let remap = |i: u32, off: f64| -> (u32, f64) {
            if let Some(v) = map(i).filter(|v| *v < limit) {
                return (v, off);
            }
            let next = (i + 1..limit).find_map(|k| map(k).filter(|v| *v < limit));
            match next {
                Some(v) => (v, 0.0),
                None => (limit - 1, 0.0),
            }
        };
        let point = |p: AnchorPoint| match axis {
            Axis::Row => {
                let (row, row_off) = remap(p.row, p.row_off);
                AnchorPoint { row, row_off, ..p }
            }
            Axis::Col => {
                let (col, col_off) = remap(p.col, p.col_off);
                AnchorPoint { col, col_off, ..p }
            }
        };
        Anchor { from: point(self.from), to: point(self.to) }
    }
}

// ---------------------------------------------------------------- references

/// A range a chart reads from. `sheet: None` means the sheet the chart is on,
/// so renaming that sheet needs no rewriting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataRef {
    pub sheet: Option<String>,
    pub range: Range,
}

impl DataRef {
    pub fn local(range: Range) -> Self {
        DataRef { sheet: None, range }
    }

    /// Parse `A1:B5`, `$A$1:$B$5`, `Sheet1!A1:B5` or `'My Sheet'!$A$1`.
    /// A leading `=` is allowed (as typed in Excel's Select Data box).
    pub fn parse(text: &str) -> Option<DataRef> {
        let text = text.trim().trim_start_matches('=').trim();
        let (sheet, area) = match text.rfind('!') {
            Some(i) => {
                let name = &text[..i];
                let name = match name.strip_prefix('\'').and_then(|n| n.strip_suffix('\'')) {
                    Some(quoted) => quoted.replace("''", "'"),
                    None => name.to_string(),
                };
                if name.is_empty() {
                    return None;
                }
                (Some(name), &text[i + 1..])
            }
            None => (None, text),
        };
        let range = match area.split_once(':') {
            Some((a, b)) => Range::new(CellRef::parse(a)?, CellRef::parse(b)?),
            None => Range::single(CellRef::parse(area)?),
        };
        Some(DataRef { sheet, range })
    }

    /// Absolute reference text qualified with a sheet name, as Excel writes it
    /// in chart XML: `Sheet1!$A$1:$B$5`, `'My Sheet'!$A$1`.
    pub fn to_formula(&self, own_sheet: &str) -> String {
        let name = self.sheet.as_deref().unwrap_or(own_sheet);
        format!("{}!{}", quote_sheet_name(name), absolute(&self.range))
    }

    fn remapped(&self, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Option<DataRef> {
        if self.sheet.is_some() {
            return Some(self.clone());
        }
        let (lo, hi) = match axis {
            Axis::Row => (self.range.start.row, self.range.end.row),
            Axis::Col => (self.range.start.col, self.range.end.col),
        };
        let limit = if axis == Axis::Row { MAX_ROWS } else { MAX_COLS };
        let new_lo = (lo..=hi).find_map(|i| map(i).filter(|v| *v < limit))?;
        let new_hi = (lo..=hi).rev().find_map(|i| map(i).filter(|v| *v < limit))?;
        let mut range = self.range;
        match axis {
            Axis::Row => {
                range.start.row = new_lo;
                range.end.row = new_hi;
            }
            Axis::Col => {
                range.start.col = new_lo;
                range.end.col = new_hi;
            }
        }
        Some(DataRef { sheet: None, range })
    }
}

/// `$A$1:$B$5` (or `$A$1` for one cell).
fn absolute(r: &Range) -> String {
    let cell = |c: CellRef| format!("${}${}", col_name(c.col), c.row + 1);
    if r.start == r.end {
        cell(r.start)
    } else {
        format!("{}:{}", cell(r.start), cell(r.end))
    }
}

/// Sheet names with spaces or punctuation are quoted in references.
pub fn quote_sheet_name(name: &str) -> String {
    let plain = name.chars().all(|c| c.is_alphanumeric() || c == '_' || c == '.')
        && !name.chars().next().is_some_and(|c| c.is_ascii_digit());
    if plain {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

// ---------------------------------------------------------------- chart model

/// The chart types on Excel's Insert > Charts menus that we support.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartType {
    ColumnClustered,
    ColumnStacked,
    ColumnPercent,
    BarClustered,
    BarStacked,
    BarPercent,
    Line,
    LineMarkers,
    Pie,
    Doughnut,
    Area,
    AreaStacked,
    Scatter,
    ScatterLines,
}

/// How series are combined at one category.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Grouping {
    Standard,
    Stacked,
    Percent,
}

impl ChartType {
    pub const ALL: [ChartType; 14] = [
        ChartType::ColumnClustered,
        ChartType::ColumnStacked,
        ChartType::ColumnPercent,
        ChartType::BarClustered,
        ChartType::BarStacked,
        ChartType::BarPercent,
        ChartType::Line,
        ChartType::LineMarkers,
        ChartType::Pie,
        ChartType::Doughnut,
        ChartType::Area,
        ChartType::AreaStacked,
        ChartType::Scatter,
        ChartType::ScatterLines,
    ];

    /// Stable id used by the UI and in tests.
    pub fn id(self) -> &'static str {
        match self {
            ChartType::ColumnClustered => "column",
            ChartType::ColumnStacked => "column-stacked",
            ChartType::ColumnPercent => "column-percent",
            ChartType::BarClustered => "bar",
            ChartType::BarStacked => "bar-stacked",
            ChartType::BarPercent => "bar-percent",
            ChartType::Line => "line",
            ChartType::LineMarkers => "line-markers",
            ChartType::Pie => "pie",
            ChartType::Doughnut => "doughnut",
            ChartType::Area => "area",
            ChartType::AreaStacked => "area-stacked",
            ChartType::Scatter => "scatter",
            ChartType::ScatterLines => "scatter-lines",
        }
    }

    pub fn from_id(id: &str) -> Option<ChartType> {
        ChartType::ALL.into_iter().find(|t| t.id() == id)
    }

    pub fn grouping(self) -> Grouping {
        match self {
            ChartType::ColumnStacked | ChartType::BarStacked | ChartType::AreaStacked => Grouping::Stacked,
            ChartType::ColumnPercent | ChartType::BarPercent => Grouping::Percent,
            _ => Grouping::Standard,
        }
    }

    /// Pie and doughnut: one colour per category, no axes.
    pub fn is_round(self) -> bool {
        matches!(self, ChartType::Pie | ChartType::Doughnut)
    }

    pub fn is_scatter(self) -> bool {
        matches!(self, ChartType::Scatter | ChartType::ScatterLines)
    }
}

/// Where the legend goes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegendPos {
    None,
    Right,
    Bottom,
    Top,
    Left,
}

impl LegendPos {
    pub fn id(self) -> &'static str {
        match self {
            LegendPos::None => "none",
            LegendPos::Right => "right",
            LegendPos::Bottom => "bottom",
            LegendPos::Top => "top",
            LegendPos::Left => "left",
        }
    }

    pub fn from_id(id: &str) -> Option<LegendPos> {
        Some(match id {
            "none" => LegendPos::None,
            "right" => LegendPos::Right,
            "bottom" => LegendPos::Bottom,
            "top" => LegendPos::Top,
            "left" => LegendPos::Left,
            _ => return None,
        })
    }
}

/// A series name: read from a cell, typed text, or none ("Series1").
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SeriesName {
    None,
    Ref(DataRef),
    Text(String),
}

#[derive(Debug, Clone, PartialEq)]
pub struct Series {
    pub name: SeriesName,
    /// Category labels (x values for scatter charts). `None`: 1, 2, 3...
    pub categories: Option<DataRef>,
    pub values: DataRef,
    /// `#RRGGBB`; `None` uses the palette.
    pub color: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    pub chart_type: ChartType,
    pub series: Vec<Series>,
    /// The range the series were built from, for "Select Data" and
    /// "Switch Row/Column". `None` for charts whose series don't form one block.
    pub source: Option<DataRef>,
    /// Series come from rows (true) or columns (false) of `source`.
    pub by_rows: bool,
    pub show_title: bool,
    /// `None`: automatic (the series name for one series, else "Chart Title").
    pub title: Option<String>,
    /// Horizontal (category) axis title; `None` hides it.
    pub x_title: Option<String>,
    /// Vertical (value) axis title; `None` hides it.
    pub y_title: Option<String>,
    pub legend: LegendPos,
    pub data_labels: bool,
}

impl Chart {
    /// A chart of `chart_type` with Excel's defaults for the given series:
    /// a title, and a legend at the bottom when there is more than one series
    /// (always for pie charts).
    pub fn new(chart_type: ChartType, series: Vec<Series>) -> Chart {
        let legend = if chart_type.is_round() || series.len() > 1 { LegendPos::Bottom } else { LegendPos::None };
        Chart {
            chart_type,
            series,
            source: None,
            by_rows: false,
            show_title: true,
            title: None,
            x_title: None,
            y_title: None,
            legend,
            data_labels: false,
        }
    }

    fn remapped(&self, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Chart {
        let series = self
            .series
            .iter()
            .filter_map(|s| {
                Some(Series {
                    name: match &s.name {
                        SeriesName::Ref(r) => r.remapped(axis, map).map_or(SeriesName::None, SeriesName::Ref),
                        other => other.clone(),
                    },
                    categories: s.categories.as_ref().and_then(|r| r.remapped(axis, map)),
                    values: s.values.remapped(axis, map)?,
                    color: s.color.clone(),
                })
            })
            .collect();
        Chart { series, source: self.source.as_ref().and_then(|r| r.remapped(axis, map)), ..self.clone() }
    }
}

/// Image bytes. Shared, because undo history keeps many copies of a sheet's drawings.
#[derive(Clone, PartialEq, Eq)]
pub struct Image {
    /// `image/png`, `image/jpeg`, `image/gif` or `image/bmp`.
    pub mime: String,
    pub data: Arc<[u8]>,
}

impl fmt::Debug for Image {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Image({}, {} bytes)", self.mime, self.data.len())
    }
}

impl Image {
    /// File extension for the image type.
    pub fn extension(&self) -> &'static str {
        match self.mime.as_str() {
            "image/jpeg" => "jpeg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            _ => "png",
        }
    }

    /// The MIME type for a file extension, if it is an image type we keep.
    pub fn mime_for_extension(ext: &str) -> Option<&'static str> {
        Some(match ext.to_ascii_lowercase().as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" | "jpe" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            _ => return None,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum DrawingContent {
    Chart(Box<Chart>),
    Picture(Image),
}

/// One floating object.
#[derive(Debug, Clone, PartialEq)]
pub struct Drawing {
    /// Unique within its sheet; stable across edits and undo.
    pub id: u32,
    /// "Chart 1", "Picture 2" (Excel's selection pane names).
    pub name: String,
    pub anchor: Anchor,
    pub content: DrawingContent,
}

impl Drawing {
    pub fn chart(&self) -> Option<&Chart> {
        match &self.content {
            DrawingContent::Chart(c) => Some(c),
            DrawingContent::Picture(_) => None,
        }
    }
}

/// A sheet's drawings, back to front.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Drawings {
    pub items: Vec<Drawing>,
}

impl Drawings {
    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    pub fn get(&self, id: u32) -> Option<&Drawing> {
        self.items.iter().find(|d| d.id == id)
    }

    pub fn get_mut(&mut self, id: u32) -> Option<&mut Drawing> {
        self.items.iter_mut().find(|d| d.id == id)
    }

    /// The id the next drawing should get.
    pub fn next_id(&self) -> u32 {
        self.items.iter().map(|d| d.id).max().unwrap_or(0) + 1
    }

    /// Excel-style default name: "Chart 3" / "Picture 3" using the next id.
    pub fn default_name(&self, kind: &str) -> String {
        format!("{kind} {}", self.next_id())
    }

    /// Append a drawing, giving it a fresh id. Returns the id.
    pub fn push(&mut self, name: String, anchor: Anchor, content: DrawingContent) -> u32 {
        let id = self.next_id();
        self.items.push(Drawing { id, name, anchor, content });
        id
    }

    pub fn remove(&mut self, id: u32) -> bool {
        let before = self.items.len();
        self.items.retain(|d| d.id != id);
        self.items.len() != before
    }

    /// Move a drawing to the front (`true`) or back of the stacking order.
    pub fn reorder(&mut self, id: u32, to_front: bool) {
        let Some(i) = self.items.iter().position(|d| d.id == id) else { return };
        let d = self.items.remove(i);
        if to_front {
            self.items.push(d);
        } else {
            self.items.insert(0, d);
        }
    }

    /// Follow an insert/delete of rows or columns: anchors move, and chart
    /// references to this sheet are rewritten like formulas.
    pub fn remapped(&self, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Drawings {
        let items = self
            .items
            .iter()
            .map(|d| Drawing {
                anchor: d.anchor.remapped(axis, map),
                content: match &d.content {
                    DrawingContent::Chart(c) => DrawingContent::Chart(Box::new(c.remapped(axis, map))),
                    other => other.clone(),
                },
                ..d.clone()
            })
            .collect();
        Drawings { items }
    }
}

// ---------------------------------------------------------------- op

/// Replace all drawings of one sheet. Inverse: the previous drawings. Images
/// are shared (`Arc`), so keeping old versions for undo is cheap.
#[derive(Debug, Clone)]
pub struct SetDrawings {
    pub sheet: usize,
    pub drawings: Drawings,
}

impl Op<Workbook> for SetDrawings {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            ws.sheet.set_drawings(self.drawings.clone());
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(SetDrawings { sheet: self.sheet, drawings: doc.sheets[self.sheet].sheet.drawings().clone() })
    }
}

// ---------------------------------------------------------------- building series

/// Excel's "current region": the block of filled cells around `at`, bounded
/// by empty rows and columns. Used when a chart is inserted with one cell selected.
pub fn current_region(sheet: &Sheet, at: CellRef) -> Range {
    let filled =
        |r: u32, c: u32| !sheet.value(CellRef::new(r, c)).is_empty() || !sheet.input(CellRef::new(r, c)).is_empty();
    let (rows, cols) = sheet.extent();
    let (mut top, mut left, mut bottom, mut right) = (at.row, at.col, at.row, at.col);
    loop {
        let mut grew = false;
        // Expand while any cell just outside the rectangle (including diagonals) is filled.
        if top > 0 && (left.saturating_sub(1)..=(right + 1).min(MAX_COLS - 1)).any(|c| filled(top - 1, c)) {
            top -= 1;
            grew = true;
        }
        if bottom + 1 < rows && (left.saturating_sub(1)..=(right + 1).min(MAX_COLS - 1)).any(|c| filled(bottom + 1, c))
        {
            bottom += 1;
            grew = true;
        }
        if left > 0 && (top.saturating_sub(1)..=(bottom + 1).min(MAX_ROWS - 1)).any(|r| filled(r, left - 1)) {
            left -= 1;
            grew = true;
        }
        if right + 1 < cols && (top.saturating_sub(1)..=(bottom + 1).min(MAX_ROWS - 1)).any(|r| filled(r, right + 1)) {
            right += 1;
            grew = true;
        }
        if !grew {
            break;
        }
    }
    Range::new(CellRef::new(top, left), CellRef::new(bottom, right))
}

fn is_number(sheet: &Sheet, at: CellRef) -> bool {
    matches!(sheet.value(at), Value::Number(_))
}

fn is_date_format(sheet: &Sheet, at: CellRef) -> bool {
    matches!(sheet.style(at).number_format, NumberFormat::Date | NumberFormat::DateTime | NumberFormat::Time)
}

/// Header detection for a data block, the way Excel guesses when you insert
/// a chart: (first row holds series names, first column holds categories).
pub fn detect_headers(sheet: &Sheet, range: Range) -> (bool, bool) {
    let (top, left) = (range.start.row, range.start.col);
    if range.rows() < 2 && range.cols() < 2 {
        return (false, false);
    }
    // An empty top-left cell means "labels across the top and down the side".
    let corner_empty = sheet.value(range.start).is_empty();
    if corner_empty && range.rows() > 1 && range.cols() > 1 {
        return (true, true);
    }
    let row_cells = |r: u32, from: u32| (from..=range.end.col).map(move |c| CellRef::new(r, c));
    let col_cells = |c: u32, from: u32| (from..=range.end.row).map(move |r| CellRef::new(r, c));
    // First column: labels if it has text (or dates) below the first row and no plain numbers.
    let first_col_labels = range.cols() > 1 && {
        let cells: Vec<CellRef> = col_cells(left, top + 1).collect();
        let cells = if cells.is_empty() { vec![range.start] } else { cells };
        cells.iter().any(|at| !sheet.value(*at).is_empty())
            && cells.iter().all(|at| !is_number(sheet, *at) || is_date_format(sheet, *at))
            && !cells.iter().all(|at| sheet.value(*at).is_empty())
    };
    // First row: names if its value cells are text (none are numbers).
    let first_row_names = range.rows() > 1 && {
        let from = if first_col_labels { left + 1 } else { left };
        let cells: Vec<CellRef> = row_cells(top, from).collect();
        cells.iter().any(|at| matches!(sheet.value(*at), Value::Text(_)))
            && cells.iter().all(|at| !is_number(sheet, *at))
    };
    (first_row_names, first_col_labels)
}

/// Build series from a block of cells like Excel's Insert Chart does.
/// `by_rows: None` lets Excel's rule pick: series run along the longer side.
/// Returns the series and whether they come from rows.
pub fn series_from_range(
    sheet: &Sheet,
    range: Range,
    chart_type: ChartType,
    by_rows: Option<bool>,
) -> (Vec<Series>, bool) {
    let (names, labels) = detect_headers(sheet, range);
    let data_top = range.start.row + names as u32;
    let data_left = range.start.col + labels as u32;
    if data_top > range.end.row || data_left > range.end.col {
        return (Vec::new(), by_rows.unwrap_or(false));
    }
    let data_rows = range.end.row - data_top + 1;
    let data_cols = range.end.col - data_left + 1;
    let by_rows = by_rows.unwrap_or(data_rows < data_cols);
    let mut out = Vec::new();
    let r = |a: CellRef, b: CellRef| DataRef::local(Range::new(a, b));
    if !by_rows {
        let categories =
            labels.then(|| r(CellRef::new(data_top, range.start.col), CellRef::new(range.end.row, range.start.col)));
        let mut cols: Vec<u32> = (data_left..=range.end.col).collect();
        // Scatter: the first numeric column holds the x values.
        let mut categories = categories;
        if chart_type.is_scatter() && !labels && cols.len() > 1 {
            let x = cols.remove(0);
            categories = Some(r(CellRef::new(data_top, x), CellRef::new(range.end.row, x)));
        }
        for c in cols {
            out.push(Series {
                name: if names {
                    SeriesName::Ref(r(CellRef::new(range.start.row, c), CellRef::new(range.start.row, c)))
                } else {
                    SeriesName::None
                },
                categories: categories.clone(),
                values: r(CellRef::new(data_top, c), CellRef::new(range.end.row, c)),
                color: None,
            });
        }
    } else {
        let categories =
            names.then(|| r(CellRef::new(range.start.row, data_left), CellRef::new(range.start.row, range.end.col)));
        let mut rows: Vec<u32> = (data_top..=range.end.row).collect();
        let mut categories = categories;
        if chart_type.is_scatter() && !names && rows.len() > 1 {
            let x = rows.remove(0);
            categories = Some(r(CellRef::new(x, data_left), CellRef::new(x, range.end.col)));
        }
        for row in rows {
            out.push(Series {
                name: if labels {
                    SeriesName::Ref(r(CellRef::new(row, range.start.col), CellRef::new(row, range.start.col)))
                } else {
                    SeriesName::None
                },
                categories: categories.clone(),
                values: r(CellRef::new(row, data_left), CellRef::new(row, range.end.col)),
                color: None,
            });
        }
    }
    (out, by_rows)
}

/// True if any series value cell holds a number (there is something to plot).
fn has_numbers(sheet: &Sheet, series: &[Series]) -> bool {
    series.iter().any(|s| s.values.range.cells().take(100_000).any(|at| is_number(sheet, at)))
}

/// A new chart for a selection. A single cell expands to its current region.
/// `None` if the range holds nothing to plot.
pub fn chart_for_range(sheet: &Sheet, selection: Range, chart_type: ChartType) -> Option<Chart> {
    let range = if selection.start == selection.end { current_region(sheet, selection.start) } else { selection };
    let (series, by_rows) = series_from_range(sheet, range, chart_type, None);
    if !has_numbers(sheet, &series) {
        return None;
    }
    let mut chart = Chart::new(chart_type, series);
    chart.source = Some(DataRef::local(range));
    chart.by_rows = by_rows;
    Some(chart)
}

// ---------------------------------------------------------------- editing

/// One change from the chart formatting pane or the Chart Design tab.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartPatch {
    Type(ChartType),
    ShowTitle(bool),
    /// `None` returns to the automatic title.
    Title(Option<String>),
    XTitle(Option<String>),
    YTitle(Option<String>),
    Legend(LegendPos),
    DataLabels(bool),
    /// Series index and colour (`None`: automatic).
    SeriesColor(usize, Option<String>),
    /// Rebuild the series from a new block of cells ("Select Data").
    Source(DataRef),
    /// Swap whether series come from rows or columns.
    SwitchRowCol,
}

impl Chart {
    /// A copy with one change applied. `sheet` is the chart's own sheet, used
    /// when series have to be rebuilt from cells. Returns `None` when the
    /// change is impossible (for example a data range with nothing to plot).
    pub fn patched(&self, patch: &ChartPatch, sheet: &Sheet) -> Option<Chart> {
        let mut c = self.clone();
        match patch {
            ChartPatch::Type(t) => {
                let was_scatter = c.chart_type.is_scatter();
                c.chart_type = *t;
                // Scatter charts read x values from the first column, so rebuild.
                if was_scatter != t.is_scatter() {
                    if let Some(src) = c.source.clone().filter(|s| s.sheet.is_none()) {
                        let (series, by_rows) = series_from_range(sheet, src.range, *t, Some(c.by_rows));
                        if !series.is_empty() {
                            c.series = keep_colors(&self.series, series);
                            c.by_rows = by_rows;
                        }
                    }
                }
                if t.is_round() && c.legend == LegendPos::None {
                    c.legend = LegendPos::Bottom;
                }
            }
            ChartPatch::ShowTitle(on) => c.show_title = *on,
            ChartPatch::Title(t) => {
                c.title = t.clone();
                c.show_title = true;
            }
            ChartPatch::XTitle(t) => c.x_title = t.clone(),
            ChartPatch::YTitle(t) => c.y_title = t.clone(),
            ChartPatch::Legend(p) => c.legend = *p,
            ChartPatch::DataLabels(on) => c.data_labels = *on,
            ChartPatch::SeriesColor(i, color) => c.series.get_mut(*i)?.color = color.clone(),
            ChartPatch::Source(src) => {
                if src.sheet.is_some() {
                    return None;
                }
                let (series, by_rows) = series_from_range(sheet, src.range, c.chart_type, None);
                if !has_numbers(sheet, &series) {
                    return None;
                }
                c.series = keep_colors(&self.series, series);
                c.by_rows = by_rows;
                c.source = Some(src.clone());
            }
            ChartPatch::SwitchRowCol => {
                let src = c.source.clone().filter(|s| s.sheet.is_none())?;
                let (series, by_rows) = series_from_range(sheet, src.range, c.chart_type, Some(!c.by_rows));
                if series.is_empty() {
                    return None;
                }
                c.series = keep_colors(&self.series, series);
                c.by_rows = by_rows;
                if c.legend == LegendPos::None && c.series.len() > 1 {
                    c.legend = LegendPos::Bottom;
                }
            }
        }
        Some(c)
    }
}

/// Custom series colours stay with the series position after a rebuild.
fn keep_colors(old: &[Series], mut new: Vec<Series>) -> Vec<Series> {
    for (n, o) in new.iter_mut().zip(old) {
        n.color = o.color.clone();
    }
    new
}

// ---------------------------------------------------------------- computing what to draw

/// One series with its current values.
#[derive(Debug, Clone, PartialEq)]
pub struct SeriesView {
    pub name: String,
    pub color: String,
    /// `None` for empty or non-numeric cells (drawn as gaps).
    pub values: Vec<Option<f64>>,
    /// Values formatted with their cells' number formats, for data labels.
    pub labels: Vec<String>,
    /// Scatter charts: x values (`None` for gaps).
    pub xs: Vec<Option<f64>>,
}

/// A numeric axis: from `min` to `max` in steps of `step`, with tick labels.
#[derive(Debug, Clone, PartialEq)]
pub struct AxisView {
    pub min: f64,
    pub max: f64,
    pub step: f64,
    pub labels: Vec<String>,
}

/// Everything the UI needs to draw a chart, with data read from the cells now.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartView {
    pub chart_type: ChartType,
    /// The title text, or `None` when hidden.
    pub title: Option<String>,
    pub x_title: Option<String>,
    pub y_title: Option<String>,
    pub legend: LegendPos,
    pub data_labels: bool,
    pub categories: Vec<String>,
    pub series: Vec<SeriesView>,
    /// Pie/doughnut: one colour per category.
    pub point_colors: Vec<String>,
    /// The value axis (vertical, or horizontal for bar charts). `None` for pie charts.
    pub value_axis: Option<AxisView>,
    /// Scatter charts: the numeric x axis.
    pub x_axis: Option<AxisView>,
}

/// Reads cells for a chart. `sheet_of(None)` is the chart's own sheet.
pub trait SheetLookup {
    fn sheet_of(&self, name: Option<&str>) -> Option<&Sheet>;
}

impl Workbook {
    fn lookup(&self, own: usize) -> impl SheetLookup + '_ {
        struct L<'a>(&'a Workbook, usize);
        impl SheetLookup for L<'_> {
            fn sheet_of(&self, name: Option<&str>) -> Option<&Sheet> {
                match name {
                    None => self.0.sheets.get(self.1).map(|s| &s.sheet),
                    Some(n) => self.0.sheets.iter().find(|s| s.name.eq_ignore_ascii_case(n)).map(|s| &s.sheet),
                }
            }
        }
        L(self, own)
    }
}

/// What to draw for drawing `id` on sheet `sheet`. `value_axis_px` is the
/// approximate length of the value axis in pixels (it decides how many
/// gridlines fit); `x_axis_px` the same for a scatter chart's x axis.
pub fn chart_view(book: &Workbook, sheet: usize, id: u32, value_axis_px: f64, x_axis_px: f64) -> Option<ChartView> {
    let chart = book.sheets.get(sheet)?.sheet.drawings().get(id)?.chart()?;
    Some(compute_view(chart, &book.lookup(sheet), value_axis_px, x_axis_px))
}

fn ref_cells(r: &DataRef) -> Vec<CellRef> {
    r.range.cells().take(100_000).collect()
}

/// Text of a series name.
pub fn series_name(s: &Series, index: usize, sheets: &dyn SheetLookup) -> String {
    match &s.name {
        SeriesName::Text(t) => t.clone(),
        SeriesName::Ref(r) => {
            let Some(sheet) = sheets.sheet_of(r.sheet.as_deref()) else { return format!("Series{}", index + 1) };
            let parts: Vec<String> =
                ref_cells(r).into_iter().map(|at| sheet.display(at)).filter(|t| !t.is_empty()).collect();
            if parts.is_empty() {
                format!("Series{}", index + 1)
            } else {
                parts.join(" ")
            }
        }
        SeriesName::None => format!("Series{}", index + 1),
    }
}

pub fn compute_view(chart: &Chart, sheets: &dyn SheetLookup, value_axis_px: f64, x_axis_px: f64) -> ChartView {
    let t = chart.chart_type;
    let mut series = Vec::new();
    let mut first_format: Option<NumberFormat> = None;
    let mut x_format: Option<NumberFormat> = None;
    let mut categories: Vec<String> = Vec::new();
    for (i, s) in chart.series.iter().enumerate() {
        let sheet = sheets.sheet_of(s.values.sheet.as_deref());
        let (mut values, mut labels) = (Vec::new(), Vec::new());
        for at in ref_cells(&s.values) {
            let (v, label) = match sheet {
                Some(sh) => match sh.value(at) {
                    Value::Number(n) if n.is_finite() => {
                        if first_format.is_none() {
                            first_format = Some(sh.style(at).number_format.clone());
                        }
                        (Some(*n), sh.display(at))
                    }
                    _ => (None, String::new()),
                },
                None => (None, String::new()),
            };
            values.push(v);
            labels.push(label);
        }
        let mut xs = Vec::new();
        if let Some(cats) = &s.categories {
            let cat_sheet = sheets.sheet_of(cats.sheet.as_deref());
            let cells = ref_cells(cats);
            if t.is_scatter() {
                for at in &cells {
                    xs.push(cat_sheet.and_then(|sh| match sh.value(*at) {
                        Value::Number(n) if n.is_finite() => {
                            if x_format.is_none() {
                                x_format = Some(sh.style(*at).number_format.clone());
                            }
                            Some(*n)
                        }
                        _ => None,
                    }));
                }
            }
            if categories.is_empty() {
                categories = cells.iter().map(|at| cat_sheet.map_or(String::new(), |sh| sh.display(*at))).collect();
            }
        }
        if t.is_scatter() && (xs.is_empty() || xs.iter().all(Option::is_none)) {
            xs = (1..=values.len()).map(|n| Some(n as f64)).collect();
        }
        series.push(SeriesView {
            name: series_name(s, i, sheets),
            color: s.color.clone().unwrap_or_else(|| auto_color(i).to_string()),
            values,
            labels,
            xs,
        });
    }
    let points = series.iter().map(|s| s.values.len()).max().unwrap_or(0);
    if categories.is_empty() {
        categories = (1..=points).map(|n| n.to_string()).collect();
    }
    categories.resize(points.max(categories.len()), String::new());

    let point_colors =
        if t.is_round() { (0..categories.len()).map(|i| auto_color(i).to_string()).collect() } else { Vec::new() };
    let value_axis = (!t.is_round()).then(|| {
        let (lo, hi) = value_extent(t, &series);
        if t.grouping() == Grouping::Percent {
            let scale = AxisScale { min: if lo < 0.0 { -1.0 } else { 0.0 }, max: 1.0, step: 0.1 };
            return scale.view(&NumberFormat::Percent { decimals: 0 });
        }
        let nf = first_format.clone().unwrap_or_default();
        auto_scale(lo, hi, max_intervals(value_axis_px), is_date_like(&nf)).view(&nf)
    });
    let x_axis = t.is_scatter().then(|| {
        let xs = series.iter().flat_map(|s| s.xs.iter().flatten().copied());
        let (lo, hi) = xs.fold((f64::INFINITY, f64::NEG_INFINITY), |(a, b), x| (a.min(x), b.max(x)));
        let (lo, hi) = if lo.is_finite() { (lo, hi) } else { (0.0, 1.0) };
        let nf = x_format.clone().unwrap_or_default();
        auto_scale(lo, hi, max_intervals(x_axis_px), is_date_like(&nf)).view(&nf)
    });

    let title = chart.show_title.then(|| match &chart.title {
        Some(text) => text.clone(),
        None if series.len() == 1 || (t.is_round() && !series.is_empty()) => series[0].name.clone(),
        None => "Chart Title".to_string(),
    });
    let axes = !t.is_round();
    ChartView {
        chart_type: t,
        title,
        x_title: chart.x_title.clone().filter(|_| axes),
        y_title: chart.y_title.clone().filter(|_| axes),
        legend: chart.legend,
        data_labels: chart.data_labels,
        categories,
        series,
        point_colors,
        value_axis,
        x_axis,
    }
}

fn is_date_like(nf: &NumberFormat) -> bool {
    matches!(nf, NumberFormat::Date | NumberFormat::DateTime)
}

fn max_intervals(px: f64) -> f64 {
    if px.is_finite() && px > 0.0 {
        (px / 27.0).clamp(2.0, 12.0)
    } else {
        7.0
    }
}

/// Smallest and largest value the value axis has to show (stacked charts:
/// the largest positive and negative stacks).
fn value_extent(t: ChartType, series: &[SeriesView]) -> (f64, f64) {
    let (mut lo, mut hi) = (f64::INFINITY, f64::NEG_INFINITY);
    match t.grouping() {
        Grouping::Standard => {
            for v in series.iter().flat_map(|s| s.values.iter().flatten()) {
                lo = lo.min(*v);
                hi = hi.max(*v);
            }
        }
        Grouping::Stacked | Grouping::Percent => {
            let n = series.iter().map(|s| s.values.len()).max().unwrap_or(0);
            for i in 0..n {
                let (mut pos, mut neg) = (0.0, 0.0);
                for s in series {
                    match s.values.get(i).copied().flatten() {
                        Some(v) if v >= 0.0 => pos += v,
                        Some(v) => neg += v,
                        None => {}
                    }
                }
                lo = lo.min(neg);
                hi = hi.max(pos);
            }
        }
    }
    if !lo.is_finite() {
        return (0.0, 0.0);
    }
    (lo, hi)
}

/// A computed axis range.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AxisScale {
    pub min: f64,
    pub max: f64,
    pub step: f64,
}

impl AxisScale {
    fn view(&self, nf: &NumberFormat) -> AxisView {
        let n = ((self.max - self.min) / self.step).round().max(0.0) as usize;
        let decimals = decimals_of(self.step);
        let labels = (0..=n)
            .map(|i| {
                let v = round_to(self.min + i as f64 * self.step, decimals);
                let v = if v == 0.0 { 0.0 } else { v }; // no "-0"
                match nf {
                    NumberFormat::General | NumberFormat::Text => format_number(v),
                    other => crate::format::format_value(&Value::Number(v), other),
                }
            })
            .collect();
        AxisView { min: self.min, max: self.max, step: self.step, labels }
    }
}

fn decimals_of(step: f64) -> i32 {
    if step >= 1.0 {
        0
    } else {
        (-step.log10().floor()) as i32 + 1
    }
}

fn round_to(v: f64, decimals: i32) -> f64 {
    let f = 10f64.powi(decimals.clamp(0, 15));
    (v * f).round() / f
}

/// Excel-like automatic axis scaling: zero is included unless the data sits
/// in a narrow band far from zero, there is about 5% headroom above the data,
/// and the step is a "nice" 1, 2 or 5 times a power of ten (whole days for dates).
pub fn auto_scale(lo: f64, hi: f64, max_intervals: f64, whole_numbers: bool) -> AxisScale {
    let (mut lo, mut hi) = (lo, hi);
    if lo == hi {
        if lo == 0.0 {
            return AxisScale { min: 0.0, max: 1.0, step: 0.2 };
        }
        if lo > 0.0 {
            lo = 0.0;
        } else {
            hi = 0.0;
        }
    }
    // Excel starts at zero unless the minimum is more than 5/6 of the maximum.
    let mut min = if lo >= 0.0 {
        if lo < hi * 5.0 / 6.0 {
            0.0
        } else {
            lo - (hi - lo) / 2.0
        }
    } else {
        lo
    };
    let mut max = if hi <= 0.0 {
        if hi > lo * 5.0 / 6.0 {
            0.0
        } else {
            hi + (hi - lo) / 2.0
        }
    } else {
        hi
    };
    let span = (max - min).max(f64::MIN_POSITIVE);
    let mut step = nice_step(span / max_intervals);
    if whole_numbers {
        step = step.max(1.0).round();
    }
    // Headroom: about 5% of the range beyond the data.
    if hi > 0.0 {
        max = ((hi + 0.05 * (hi - min)) / step).ceil() * step;
    } else {
        max = (max / step).ceil() * step;
    }
    if lo < 0.0 {
        min = ((lo - 0.05 * (max - lo)) / step).floor() * step;
    } else {
        min = (min / step).floor() * step;
    }
    // Recompute the step if headroom pushed the count of intervals too high.
    let intervals = ((max - min) / step).round();
    if intervals > max_intervals.ceil() + 1.0 {
        return auto_scale(lo, hi, max_intervals * 0.75, whole_numbers);
    }
    AxisScale { min, max, step }
}

fn nice_step(raw: f64) -> f64 {
    let exp = raw.log10().floor();
    let base = 10f64.powf(exp);
    let f = raw / base;
    let nice = if f <= 1.0 {
        1.0
    } else if f <= 2.0 {
        2.0
    } else if f <= 2.5 && exp >= 1.0 {
        2.5
    } else if f <= 5.0 {
        5.0
    } else {
        10.0
    };
    nice * base
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn rng(s: &str) -> Range {
        DataRef::parse(s).unwrap().range
    }

    fn sheet(cells: &[(&str, &str)]) -> Sheet {
        let mut s = Sheet::new();
        s.set_inputs(cells.iter().map(|(a, i)| (at(a), i.to_string())));
        s
    }

    /// Region\Q1\Q2\Q3 with three regions.
    fn table() -> Sheet {
        sheet(&[
            ("A1", "Region"),
            ("B1", "Q1"),
            ("C1", "Q2"),
            ("A2", "North"),
            ("B2", "10"),
            ("C2", "20"),
            ("A3", "South"),
            ("B3", "30"),
            ("C3", "40"),
            ("A4", "East"),
            ("B4", "50"),
            ("C4", "5"),
        ])
    }

    struct One<'a>(&'a Sheet);
    impl SheetLookup for One<'_> {
        fn sheet_of(&self, name: Option<&str>) -> Option<&Sheet> {
            name.is_none().then_some(self.0)
        }
    }

    #[test]
    fn parses_and_writes_references() {
        assert_eq!(DataRef::parse("A1:B5"), Some(DataRef::local(Range::new(at("A1"), at("B5")))));
        assert_eq!(DataRef::parse("=$B$2").unwrap().range, Range::single(at("B2")));
        let r = DataRef::parse("'My ''Q'' Sheet'!$A$1:$A$3").unwrap();
        assert_eq!(r.sheet.as_deref(), Some("My 'Q' Sheet"));
        assert_eq!(r.to_formula("x"), "'My ''Q'' Sheet'!$A$1:$A$3");
        assert_eq!(DataRef::local(Range::single(at("C3"))).to_formula("Sheet1"), "Sheet1!$C$3");
        assert_eq!(DataRef::parse("nonsense"), None);
        assert_eq!(DataRef::parse("!A1"), None);
        assert_eq!(quote_sheet_name("2024"), "'2024'");
    }

    #[test]
    fn detects_headers_and_series_like_excel() {
        let s = table();
        assert_eq!(detect_headers(&s, rng("A1:C4")), (true, true));
        let (series, by_rows) = series_from_range(&s, rng("A1:C4"), ChartType::ColumnClustered, None);
        // 3 rows x 2 columns of data: series are the (fewer) columns.
        assert!(!by_rows);
        assert_eq!(series.len(), 2);
        assert_eq!(series[0].name, SeriesName::Ref(DataRef::local(Range::single(at("B1")))));
        assert_eq!(series[0].values.range, rng("B2:B4"));
        assert_eq!(series[1].categories.as_ref().unwrap().range, rng("A2:A4"));

        let (series, by_rows) = series_from_range(&s, rng("A1:C4"), ChartType::ColumnClustered, Some(true));
        assert!(by_rows);
        assert_eq!(series.len(), 3);
        assert_eq!(series[2].values.range, rng("B4:C4"));
        assert_eq!(series[2].categories.as_ref().unwrap().range, rng("B1:C1"));

        // Numbers only: no headers, Series1..
        let nums = sheet(&[("A1", "1"), ("A2", "2"), ("B1", "3"), ("B2", "4"), ("A3", "5"), ("B3", "6")]);
        assert_eq!(detect_headers(&nums, rng("A1:B3")), (false, false));
        // Empty corner: first row and column are labels even when numeric (years).
        let years = sheet(&[("B1", "2023"), ("C1", "2024"), ("A2", "Tea"), ("B2", "5"), ("C2", "7")]);
        assert_eq!(detect_headers(&years, rng("A1:C2")), (true, true));
    }

    #[test]
    fn chart_from_one_cell_uses_the_current_region() {
        let mut s = table();
        s.set_input(at("E10"), "far away");
        assert_eq!(current_region(&s, at("B3")), rng("A1:C4"));
        let chart = chart_for_range(&s, Range::single(at("B3")), ChartType::Pie).unwrap();
        assert_eq!(chart.source.unwrap().range, rng("A1:C4"));
        assert_eq!(chart.legend, LegendPos::Bottom);
        assert!(chart_for_range(&s, Range::single(at("H20")), ChartType::Line).is_none());
    }

    #[test]
    fn view_reads_live_values() {
        let mut s = table();
        let chart = chart_for_range(&s, rng("A1:C4"), ChartType::ColumnClustered).unwrap();
        let v = compute_view(&chart, &One(&s), 200.0, 400.0);
        assert_eq!(v.categories, vec!["North", "South", "East"]);
        assert_eq!(v.series[0].name, "Q1");
        assert_eq!(v.series[0].values, vec![Some(10.0), Some(30.0), Some(50.0)]);
        assert_eq!(v.series[1].color, "#ED7D31");
        assert_eq!(v.title.as_deref(), Some("Chart Title"));
        let axis = v.value_axis.unwrap();
        assert_eq!((axis.min, axis.max, axis.step), (0.0, 60.0, 10.0));
        assert_eq!(axis.labels.first().map(String::as_str), Some("0"));

        s.set_input(at("B2"), "=B3*4");
        s.set_input(at("C3"), "text");
        let v = compute_view(&chart, &One(&s), 200.0, 400.0);
        assert_eq!(v.series[0].values[0], Some(120.0));
        assert_eq!(v.series[1].values[1], None);
    }

    #[test]
    fn single_series_title_and_formats() {
        let s =
            sheet(&[("A1", "Month"), ("B1", "Sales"), ("A2", "Jan"), ("B2", "₹1,500"), ("A3", "Feb"), ("B3", "2500")]);
        let chart = chart_for_range(&s, rng("A1:B3"), ChartType::Line).unwrap();
        assert_eq!(chart.legend, LegendPos::None);
        let v = compute_view(&chart, &One(&s), 200.0, 400.0);
        assert_eq!(v.title.as_deref(), Some("Sales"));
        assert_eq!(v.series[0].labels[0], "₹1,500");
        assert!(v.value_axis.unwrap().labels.iter().all(|l| l.starts_with('₹')));
    }

    #[test]
    fn stacked_and_percent_axes() {
        let s = table();
        let mut chart = chart_for_range(&s, rng("A1:C4"), ChartType::ColumnStacked).unwrap();
        let v = compute_view(&chart, &One(&s), 200.0, 400.0);
        assert_eq!(v.value_axis.unwrap().max, 80.0); // largest stack: 30+40 = 70
        chart.chart_type = ChartType::BarPercent;
        let axis = compute_view(&chart, &One(&s), 200.0, 400.0).value_axis.unwrap();
        assert_eq!((axis.min, axis.max), (0.0, 1.0));
        assert_eq!(axis.labels.last().map(String::as_str), Some("100%"));
    }

    #[test]
    fn scatter_uses_first_column_as_x() {
        let s = sheet(&[("A1", "1"), ("B1", "10"), ("A2", "2"), ("B2", "20"), ("A3", "4"), ("B3", "15")]);
        let chart = chart_for_range(&s, rng("A1:B3"), ChartType::Scatter).unwrap();
        assert_eq!(chart.series.len(), 1);
        let v = compute_view(&chart, &One(&s), 200.0, 400.0);
        assert_eq!(v.series[0].xs, vec![Some(1.0), Some(2.0), Some(4.0)]);
        assert_eq!(v.x_axis.unwrap().min, 0.0);
    }

    #[test]
    fn axis_scaling_matches_excel() {
        let s = auto_scale(10.0, 50.0, 7.4, false);
        assert_eq!((s.min, s.max, s.step), (0.0, 60.0, 10.0));
        let s = auto_scale(0.0, 100.0, 7.4, false);
        assert_eq!((s.min, s.max, s.step), (0.0, 120.0, 20.0));
        // A narrow band far from zero does not start at zero.
        let s = auto_scale(950.0, 1000.0, 7.4, false);
        assert!(s.min > 0.0 && s.min <= 950.0, "{s:?}");
        // Negative values.
        let s = auto_scale(-30.0, 40.0, 7.4, false);
        assert!(s.min <= -30.0 && s.max >= 40.0 && s.min < 0.0);
        assert_eq!(auto_scale(0.0, 0.0, 5.0, false), AxisScale { min: 0.0, max: 1.0, step: 0.2 });
        let s = auto_scale(0.1, 0.9, 10.0, false);
        assert_eq!(s.min, 0.0);
        assert_eq!(s.view(&NumberFormat::General).labels[1], "0.1");
    }

    #[test]
    fn patches_and_switch_row_col() {
        let s = table();
        let chart = chart_for_range(&s, rng("A1:C4"), ChartType::ColumnClustered).unwrap();
        let c = chart.patched(&ChartPatch::SeriesColor(1, Some("#FF0000".into())), &s).unwrap();
        assert_eq!(c.series[1].color.as_deref(), Some("#FF0000"));
        let switched = c.patched(&ChartPatch::SwitchRowCol, &s).unwrap();
        assert!(switched.by_rows);
        assert_eq!(switched.series.len(), 3);
        assert_eq!(switched.series[1].color.as_deref(), Some("#FF0000"));
        let back = switched.patched(&ChartPatch::SwitchRowCol, &s).unwrap();
        assert_eq!(back.series, c.series);
        let narrower = c.patched(&ChartPatch::Source(DataRef::parse("A1:B4").unwrap()), &s).unwrap();
        assert_eq!(narrower.series.len(), 1);
        assert!(c.patched(&ChartPatch::Source(DataRef::parse("H1:H4").unwrap()), &s).is_none());
        let titled = c.patched(&ChartPatch::Title(Some("Sales".into())), &s).unwrap();
        assert_eq!(compute_view(&titled, &One(&s), 200.0, 200.0).title.as_deref(), Some("Sales"));
        let hidden = titled.patched(&ChartPatch::ShowTitle(false), &s).unwrap();
        assert_eq!(compute_view(&hidden, &One(&s), 200.0, 200.0).title, None);
        let pie = c.patched(&ChartPatch::Type(ChartType::Pie), &s).unwrap();
        let v = compute_view(&pie, &One(&s), 200.0, 200.0);
        assert_eq!(v.point_colors.len(), 3);
        assert!(v.value_axis.is_none());
    }

    #[test]
    fn structural_edits_move_anchors_and_references() {
        let mut s = table();
        let chart = chart_for_range(&s, rng("A1:C4"), ChartType::ColumnClustered).unwrap();
        let anchor = Anchor { from: AnchorPoint::new(5, 1, 3.0, 4.0), to: AnchorPoint::new(15, 6, 0.0, 0.0) };
        let mut d = Drawings::default();
        let id = d.push("Chart 1".into(), anchor, DrawingContent::Chart(Box::new(chart)));
        s.set_drawings(d);

        let ins = s.with_inserted(Axis::Row, 1, 2);
        let moved = ins.drawings().get(id).unwrap();
        assert_eq!(moved.anchor.from.row, 7);
        assert_eq!(moved.anchor.from.row_off, 3.0);
        assert_eq!(moved.chart().unwrap().series[0].values.range, rng("B4:B6"));
        assert_eq!(moved.chart().unwrap().source.as_ref().unwrap().range, rng("A1:C6"));

        let del = s.with_deleted(Axis::Col, 1, 1);
        let c = del.drawings().get(id).unwrap();
        assert_eq!(c.chart().unwrap().series.len(), 1, "series in the deleted column is dropped");
        assert_eq!(c.anchor.from.col, 1);
        assert_eq!(c.anchor.from.col_off, 0.0, "corner in the deleted column snaps to the next one");
        assert_eq!(c.anchor.to.col, 5);

        let sorted = s.with_sorted_rows(rng("A2:C4"), 1, false);
        assert_eq!(sorted.drawings(), s.drawings());
    }

    #[test]
    fn set_drawings_is_undoable() {
        use office_core::history::History;
        let mut wb = Workbook::new();
        let mut history = History::default();
        let mut d = Drawings::default();
        let image = Image { mime: "image/png".into(), data: Arc::from(vec![1u8, 2, 3]) };
        let anchor = Anchor::from_pixels_uniform(10.0, 10.0, 100.0, 50.0, 100.0, 22.0);
        d.push("Picture 1".into(), anchor, DrawingContent::Picture(image));
        history.apply(&mut wb, Box::new(SetDrawings { sheet: 0, drawings: d.clone() }));
        assert_eq!(wb.sheets[0].sheet.drawings(), &d);
        history.undo(&mut wb);
        assert!(wb.sheets[0].sheet.drawings().is_empty());
        history.redo(&mut wb);
        assert_eq!(wb.sheets[0].sheet.drawings().items.len(), 1);
        assert_eq!(d.next_id(), 2);
        assert_eq!(anchor.to.col, 1);
        assert_eq!(anchor.to.col_off, 10.0);
    }

    #[test]
    fn reorders_and_removes() {
        let mut d = Drawings::default();
        let image = || DrawingContent::Picture(Image { mime: "image/png".into(), data: Arc::from(vec![0u8]) });
        let a = Anchor::from_pixels_uniform(0.0, 0.0, 10.0, 10.0, 100.0, 22.0);
        let one = d.push(d.default_name("Picture"), a, image());
        let two = d.push(d.default_name("Picture"), a, image());
        assert_eq!(d.items[1].name, "Picture 2");
        d.reorder(two, false);
        assert_eq!(d.items[0].id, two);
        d.reorder(two, true);
        assert_eq!(d.items[1].id, two);
        assert!(d.remove(one));
        assert!(!d.remove(one));
        assert_eq!(auto_color(7), "#9E480E");
    }
}
