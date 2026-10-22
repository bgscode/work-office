//! Bindings for charts and pictures (`sheets_engine::chart`). Thin: parse
//! arguments, call the engine, apply one `SetDrawings` op per edit, and write
//! JSON for the UI (types in `apps/web/src/sheets/charts/types.ts`).
//!
//! Anchors cross the boundary as 8 numbers:
//! `[fromRow, fromCol, fromRowOffPx, fromColOffPx, toRow, toCol, toRowOffPx, toColOffPx]`.

use std::sync::Arc;

use sheets_engine::chart::{
    self, Anchor, AnchorPoint, AxisView, ChartPatch, ChartType, DataRef, DrawingContent, Image, LegendPos, SeriesName,
    SetDrawings,
};
use sheets_engine::shape::{self, Shape};
use sheets_engine::{CellRef, Range};
use wasm_bindgen::prelude::*;

use crate::SheetsDocument;

fn anchor_from(v: &[f64]) -> Option<Anchor> {
    if v.len() != 8 || v.iter().any(|x| !x.is_finite()) {
        return None;
    }
    let idx = |x: f64, max: u32| (x.max(0.0) as u32).min(max - 1);
    let point = |o: usize| {
        AnchorPoint::new(idx(v[o], sheets_engine::MAX_ROWS), idx(v[o + 1], sheets_engine::MAX_COLS), v[o + 2], v[o + 3])
    };
    Some(Anchor { from: point(0), to: point(4) })
}

fn anchor_json(a: &Anchor) -> String {
    let p = |p: &AnchorPoint| format!("{},{},{},{}", p.row, p.col, num(p.row_off), num(p.col_off));
    format!("[{},{}]", p(&a.from), p(&a.to))
}

fn num(x: f64) -> String {
    if x.is_finite() {
        format!("{x}")
    } else {
        "null".into()
    }
}

/// A JSON string literal.
fn js(s: &str) -> String {
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

fn opt(s: &Option<String>) -> String {
    s.as_deref().map_or("null".into(), js)
}

fn list(items: impl IntoIterator<Item = String>) -> String {
    format!("[{}]", items.into_iter().collect::<Vec<_>>().join(","))
}

fn axis_json(a: &Option<AxisView>) -> String {
    match a {
        None => "null".into(),
        Some(a) => format!(
            r#"{{"min":{},"max":{},"step":{},"labels":{}}}"#,
            num(a.min),
            num(a.max),
            num(a.step),
            list(a.labels.iter().map(|l| js(l)))
        ),
    }
}

/// `A1:C4`, or `Sheet2!A1:C4` for another sheet.
fn ref_text(r: &DataRef) -> String {
    match &r.sheet {
        None => r.range.to_string(),
        Some(s) => format!("{}!{}", chart::quote_sheet_name(s), r.range),
    }
}

impl SheetsDocument {
    fn set_drawings(&mut self, sheet: usize, drawings: chart::Drawings) {
        self.apply(SetDrawings { sheet, drawings });
    }
}

#[wasm_bindgen]
impl SheetsDocument {
    /// Charts and pictures on a sheet, back to front, as JSON (`DrawingInfo[]`).
    #[wasm_bindgen(js_name = drawingsJson)]
    pub fn drawings_json(&self, sheet: usize) -> String {
        let book_sheet = self.sheet(sheet);
        list(book_sheet.drawings().items.iter().map(|d| {
            let head = format!(r#""id":{},"name":{},"anchor":{}"#, d.id, js(&d.name), anchor_json(&d.anchor));
            match &d.content {
                DrawingContent::Picture(img) => {
                    format!(r#"{{{head},"kind":"picture","mime":{},"bytes":{}}}"#, js(&img.mime), img.data.len())
                }
                DrawingContent::Shape(sh) => format!(
                    concat!(
                        r#"{{{},"kind":"shape","shape":{{"preset":{},"fill":{},"line":{},"lineWidth":{},"arrowEnd":{},"#,
                        r#""flipV":{},"text":{},"textColor":{},"fontSize":{},"bold":{},"italic":{},"align":{},"textTop":{}}}}}"#
                    ),
                    head,
                    js(&sh.preset),
                    opt(&sh.fill),
                    opt(&sh.line),
                    num(sh.line_width),
                    sh.arrow_end,
                    sh.flip_v,
                    js(&sh.text),
                    js(&sh.text_color),
                    num(sh.font_size),
                    sh.bold,
                    sh.italic,
                    js(sh.align.id()),
                    sh.text_top
                ),
                DrawingContent::Chart(c) => {
                    let series = list(c.series.iter().map(|s| {
                        let name = match &s.name {
                            SeriesName::Ref(r) => js(&format!("={}", ref_text(r))),
                            SeriesName::Text(t) => js(t),
                            SeriesName::None => "null".into(),
                        };
                        format!(
                            r#"{{"name":{name},"values":{},"categories":{},"color":{}}}"#,
                            js(&ref_text(&s.values)),
                            s.categories.as_ref().map_or("null".into(), |r| js(&ref_text(r))),
                            opt(&s.color)
                        )
                    }));
                    format!(
                        concat!(
                            r#"{{{},"kind":"chart","chart":{{"type":{},"showTitle":{},"title":{},"xTitle":{},"yTitle":{},"#,
                            r#""legend":{},"dataLabels":{},"source":{},"byRows":{},"series":{}}}}}"#
                        ),
                        head,
                        js(c.chart_type.id()),
                        c.show_title,
                        opt(&c.title),
                        opt(&c.x_title),
                        opt(&c.y_title),
                        js(c.legend.id()),
                        c.data_labels,
                        c.source.as_ref().map_or("null".into(), |r| js(&ref_text(r))),
                        c.by_rows,
                        series
                    )
                }
            }
        }))
    }

    /// What to draw for a chart, with its data read from the cells now
    /// (`ChartView` JSON), or "" if `id` is not a chart. `value_axis_px` /
    /// `x_axis_px`: approximate axis lengths, which decide how many gridlines fit.
    #[wasm_bindgen(js_name = chartViewJson)]
    pub fn chart_view_json(&self, sheet: usize, id: u32, value_axis_px: f64, x_axis_px: f64) -> String {
        let Some(v) = chart::chart_view(&self.book, sheet, id, value_axis_px, x_axis_px) else { return String::new() };
        let series = list(v.series.iter().map(|s| {
            format!(
                r#"{{"name":{},"color":{},"values":{},"labels":{},"xs":{}}}"#,
                js(&s.name),
                js(&s.color),
                list(s.values.iter().map(|x| x.map_or("null".into(), num))),
                list(s.labels.iter().map(|l| js(l))),
                list(s.xs.iter().map(|x| x.map_or("null".into(), num)))
            )
        }));
        format!(
            concat!(
                r#"{{"type":{},"title":{},"xTitle":{},"yTitle":{},"legend":{},"dataLabels":{},"categories":{},"#,
                r#""series":{},"pointColors":{},"valueAxis":{},"xAxis":{}}}"#
            ),
            js(v.chart_type.id()),
            opt(&v.title),
            opt(&v.x_title),
            opt(&v.y_title),
            js(v.legend.id()),
            v.data_labels,
            list(v.categories.iter().map(|c| js(c))),
            series,
            list(v.point_colors.iter().map(|c| js(c))),
            axis_json(&v.value_axis),
            axis_json(&v.x_axis)
        )
    }

    /// Insert a chart of `chart_type` (see `ChartType::id`) built from a
    /// rectangle (one cell expands to its surrounding data block), placed at
    /// `anchor`. Returns the new drawing id, or -1 if there is nothing to plot.
    #[wasm_bindgen(js_name = insertChart)]
    pub fn insert_chart(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        chart_type: &str,
        anchor: Vec<f64>,
    ) -> Result<i32, JsError> {
        let t =
            ChartType::from_id(chart_type).ok_or_else(|| JsError::new(&format!("unknown chart type {chart_type}")))?;
        let anchor = anchor_from(&anchor).ok_or_else(|| JsError::new("bad anchor"))?;
        let s = self.sheet(sheet);
        let range = Range::new(CellRef::new(top, left), CellRef::new(bottom, right));
        let Some(c) = chart::chart_for_range(s, range, t) else { return Ok(-1) };
        let mut drawings = s.drawings().clone();
        let name = drawings.default_name("Chart");
        let id = drawings.push(name, anchor, DrawingContent::Chart(Box::new(c)));
        self.set_drawings(sheet, drawings);
        Ok(id as i32)
    }

    /// The data block a chart inserted at this selection would use:
    /// [top, left, bottom, right], or empty if there is nothing to plot.
    #[wasm_bindgen(js_name = chartSourceRange)]
    pub fn chart_source_range(&self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32) -> Vec<u32> {
        let range = Range::new(CellRef::new(top, left), CellRef::new(bottom, right));
        match chart::chart_for_range(self.sheet(sheet), range, ChartType::ColumnClustered).and_then(|c| c.source) {
            Some(r) => vec![r.range.start.row, r.range.start.col, r.range.end.row, r.range.end.col],
            None => Vec::new(),
        }
    }

    /// Insert a picture (`mime`: image/png, image/jpeg, image/gif or image/bmp). Returns its id.
    #[wasm_bindgen(js_name = insertPicture)]
    pub fn insert_picture(&mut self, sheet: usize, bytes: &[u8], mime: &str, anchor: Vec<f64>) -> Result<u32, JsError> {
        let ext = mime.strip_prefix("image/").unwrap_or("");
        let mime =
            Image::mime_for_extension(ext).ok_or_else(|| JsError::new(&format!("unsupported image type {mime}")))?;
        let anchor = anchor_from(&anchor).ok_or_else(|| JsError::new("bad anchor"))?;
        let mut drawings = self.sheet(sheet).drawings().clone();
        let name = drawings.default_name("Picture");
        let image = Image { mime: mime.to_string(), data: Arc::from(bytes) };
        let id = drawings.push(name, anchor, DrawingContent::Picture(image));
        self.set_drawings(sheet, drawings);
        Ok(id)
    }

    /// Insert a shape (`preset`: a DrawingML preset from `SHAPE_PRESETS`, or
    /// "textBox") at `anchor`. Returns its id. `flip_v`: a line drawn from
    /// bottom-left to top-right.
    #[wasm_bindgen(js_name = insertShape)]
    pub fn insert_shape(&mut self, sheet: usize, preset: &str, anchor: Vec<f64>, flip_v: bool) -> Result<u32, JsError> {
        let shape = match preset {
            "textBox" => Shape::text_box(),
            p if shape::is_known_preset(p) => Shape { flip_v, ..Shape::new(p) },
            p => return Err(JsError::new(&format!("unknown shape {p}"))),
        };
        let anchor = anchor_from(&anchor).ok_or_else(|| JsError::new("bad anchor"))?;
        let mut drawings = self.sheet(sheet).drawings().clone();
        let name = drawings.default_name(if preset == "textBox" { "TextBox" } else { "Shape" });
        let id = drawings.push(name, anchor, DrawingContent::Shape(Box::new(shape)));
        self.set_drawings(sheet, drawings);
        Ok(id)
    }

    /// One Shape Format change (see `Shape::patched` for keys). Returns false
    /// when nothing changed or the value was not accepted.
    #[wasm_bindgen(js_name = formatShape)]
    pub fn format_shape(&mut self, sheet: usize, id: u32, key: &str, value: &str) -> bool {
        let mut drawings = self.sheet(sheet).drawings().clone();
        let Some(d) = drawings.get_mut(id) else { return false };
        let DrawingContent::Shape(sh) = &d.content else { return false };
        let Some(next) = sh.patched(key, value) else { return false };
        if **sh == next {
            return false;
        }
        d.content = DrawingContent::Shape(Box::new(next));
        self.set_drawings(sheet, drawings);
        true
    }

    /// The Shapes gallery: `[group, preset, label, group, preset, label, ...]`.
    #[wasm_bindgen(js_name = shapePresets)]
    pub fn shape_presets() -> Vec<String> {
        shape::SHAPE_PRESETS.iter().flat_map(|(g, p, l)| [g.to_string(), p.to_string(), l.to_string()]).collect()
    }

    /// The bytes of a picture (empty if `id` is not a picture).
    #[wasm_bindgen(js_name = pictureData)]
    pub fn picture_data(&self, sheet: usize, id: u32) -> Vec<u8> {
        match self.sheet(sheet).drawings().get(id).map(|d| &d.content) {
            Some(DrawingContent::Picture(img)) => img.data.to_vec(),
            _ => Vec::new(),
        }
    }

    /// Move or resize a drawing. Returns false if it doesn't exist.
    #[wasm_bindgen(js_name = setDrawingAnchor)]
    pub fn set_drawing_anchor(&mut self, sheet: usize, id: u32, anchor: Vec<f64>) -> bool {
        let Some(anchor) = anchor_from(&anchor) else { return false };
        let mut drawings = self.sheet(sheet).drawings().clone();
        let Some(d) = drawings.get_mut(id) else { return false };
        if d.anchor == anchor {
            return false;
        }
        d.anchor = anchor;
        self.set_drawings(sheet, drawings);
        true
    }

    #[wasm_bindgen(js_name = deleteDrawing)]
    pub fn delete_drawing(&mut self, sheet: usize, id: u32) -> bool {
        let mut drawings = self.sheet(sheet).drawings().clone();
        if !drawings.remove(id) {
            return false;
        }
        self.set_drawings(sheet, drawings);
        true
    }

    /// Bring to front (`to_front`) or send to back.
    #[wasm_bindgen(js_name = reorderDrawing)]
    pub fn reorder_drawing(&mut self, sheet: usize, id: u32, to_front: bool) -> bool {
        let mut drawings = self.sheet(sheet).drawings().clone();
        drawings.reorder(id, to_front);
        if &drawings == self.sheet(sheet).drawings() {
            return false;
        }
        self.set_drawings(sheet, drawings);
        true
    }

    /// One chart formatting change. `key`/`value`: type: a chart type id;
    /// showTitle|dataLabels: "1"/"0"; title: text ("" = automatic);
    /// xTitle|yTitle: text ("" hides the axis title); legend: none|right|bottom|top|left;
    /// seriesColor: "index:#RRGGBB" ("index:" = automatic); source: a range like
    /// "A1:C5"; switch: "" (swap rows/columns). Returns false if nothing changed
    /// (for example a data range with no numbers).
    #[wasm_bindgen(js_name = formatChart)]
    pub fn format_chart(&mut self, sheet: usize, id: u32, key: &str, value: &str) -> Result<bool, JsError> {
        let text = |v: &str| (!v.is_empty()).then(|| v.to_string());
        let bad = || JsError::new(&format!("bad chart format {key}={value}"));
        let patch = match key {
            "type" => ChartPatch::Type(ChartType::from_id(value).ok_or_else(bad)?),
            "showTitle" => ChartPatch::ShowTitle(value == "1"),
            "title" => ChartPatch::Title(text(value)),
            "xTitle" => ChartPatch::XTitle(text(value)),
            "yTitle" => ChartPatch::YTitle(text(value)),
            "legend" => ChartPatch::Legend(LegendPos::from_id(value).ok_or_else(bad)?),
            "dataLabels" => ChartPatch::DataLabels(value == "1"),
            "seriesColor" => {
                let (i, color) = value.split_once(':').ok_or_else(bad)?;
                ChartPatch::SeriesColor(i.parse().map_err(|_| bad())?, text(color))
            }
            "source" => ChartPatch::Source(DataRef::parse(value).ok_or_else(bad)?),
            "switch" => ChartPatch::SwitchRowCol,
            _ => return Err(bad()),
        };
        let s = self.sheet(sheet);
        let mut drawings = s.drawings().clone();
        let Some(d) = drawings.get_mut(id) else { return Ok(false) };
        let DrawingContent::Chart(c) = &d.content else { return Ok(false) };
        let Some(next) = c.patched(&patch, s) else { return Ok(false) };
        if **c == next {
            return Ok(false);
        }
        d.content = DrawingContent::Chart(Box::new(next));
        self.set_drawings(sheet, drawings);
        Ok(true)
    }
}
