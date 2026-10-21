//! Charts and pictures in `.xlsx` files (DrawingML).
//!
//! ```text
//! xl/worksheets/sheetN.xml               <drawing r:id="..."/>
//! xl/worksheets/_rels/sheetN.xml.rels    -> ../drawings/drawingD.xml
//! xl/drawings/drawingD.xml               <xdr:twoCellAnchor> per object: from/to cells + a chart frame or a picture
//! xl/drawings/_rels/drawingD.xml.rels    -> ../charts/chartC.xml, ../media/imageM.png
//! xl/charts/chartC.xml                   <c:chartSpace>: chart type, series references (+ cached values), titles, legend
//! xl/media/imageM.png                    picture bytes
//! ```
//!
//! Written so Excel opens the charts with its Microsoft 365 look (fonts,
//! colours, gridlines). Read back: bar/column, line, pie, doughnut, area and
//! scatter charts (other chart types are skipped), and PNG/JPEG/GIF/BMP pictures.

use std::collections::HashMap;
use std::io::{Cursor, Read};
use std::sync::Arc;

use office_core::{OfficeError, Result};
use sheets_engine::chart::{
    self, Anchor, AnchorPoint, Chart, ChartType, DataRef, Drawing, DrawingContent, Drawings, Grouping, Image,
    LegendPos, Series, SeriesName, SheetLookup, EMU_PER_PX,
};
use sheets_engine::{Sheet, Workbook};
use zip::ZipArchive;

use super::{attr, escape, relationships, walk, XmlEvent, NS_PKG_REL, NS_REL};

const NS_XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_C: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CT_CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

/// Text colours of Excel's default chart style.
const TEXT_GREY: &str = "595959";
const LINE_GREY: &str = "D9D9D9";
/// Default row height in pixels, for anchors that only give a size.
const ROW_PX: f64 = 20.0;

// ---------------------------------------------------------------- writing

/// Collects drawing, chart and media parts while sheets are written.
#[derive(Default)]
pub(super) struct DrawingWriter {
    drawings: usize,
    charts: usize,
    images: usize,
    pub(super) parts: Vec<(String, String)>,
    pub(super) media: Vec<(String, Arc<[u8]>)>,
    content_types: Vec<String>,
    extensions: Vec<&'static str>,
}

impl DrawingWriter {
    /// Write the drawing parts for sheet `id` (1-based) and return its
    /// worksheet XML with a `<drawing>` element added. Sheets without
    /// drawings are returned unchanged.
    pub(super) fn add_sheet(&mut self, id: usize, book: &Workbook, index: usize, worksheet_xml: String) -> String {
        let ws = &book.sheets[index];
        let drawings = ws.sheet.drawings();
        if drawings.is_empty() {
            return worksheet_xml;
        }
        self.drawings += 1;
        let d = self.drawings;
        let lookup = Lookup { book, own: index };
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}">"#
        );
        let mut rels =
            format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{NS_PKG_REL}">"#);
        for (n, item) in drawings.items.iter().enumerate() {
            let rid = format!("rId{}", n + 1);
            let shape_id = n + 2;
            let name = escape(&item.name);
            match &item.content {
                DrawingContent::Chart(c) => {
                    self.charts += 1;
                    let file = format!("chart{}.xml", self.charts);
                    self.parts.push((format!("xl/charts/{file}"), chart_xml(c, &ws.name, &lookup)));
                    self.content_types
                        .push(format!(r#"<Override PartName="/xl/charts/{file}" ContentType="{CT_CHART}"/>"#));
                    rels.push_str(&format!(
                        r#"<Relationship Id="{rid}" Type="{NS_REL}/chart" Target="../charts/{file}"/>"#
                    ));
                    xml.push_str("<xdr:twoCellAnchor>");
                    xml.push_str(&anchor_xml(&item.anchor));
                    xml.push_str(&format!(
                        concat!(
                            r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{id}" name="{name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"#,
                            r#"<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>"#,
                            r#"<a:graphic><a:graphicData uri="{c}"><c:chart xmlns:c="{c}" xmlns:r="{r}" r:id="{rid}"/></a:graphicData></a:graphic>"#,
                            r#"</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"#
                        ),
                        id = shape_id,
                        name = name,
                        c = NS_C,
                        r = NS_REL,
                        rid = rid
                    ));
                }
                DrawingContent::Picture(image) => {
                    self.images += 1;
                    let ext = image.extension();
                    let file = format!("image{}.{ext}", self.images);
                    self.media.push((format!("xl/media/{file}"), image.data.clone()));
                    if !self.extensions.contains(&ext) {
                        self.extensions.push(ext);
                    }
                    rels.push_str(&format!(
                        r#"<Relationship Id="{rid}" Type="{NS_REL}/image" Target="../media/{file}"/>"#
                    ));
                    let (cx, cy) = anchor_size_emu(&ws.sheet, &item.anchor);
                    xml.push_str(r#"<xdr:twoCellAnchor editAs="oneCell">"#);
                    xml.push_str(&anchor_xml(&item.anchor));
                    xml.push_str(&format!(
                        concat!(
                            r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id}" name="{name}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>"#,
                            r#"<xdr:blipFill><a:blip xmlns:r="{r}" r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"#,
                            r#"<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>"#,
                            r#"</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#
                        ),
                        id = shape_id,
                        name = name,
                        r = NS_REL,
                        rid = rid,
                        cx = cx,
                        cy = cy
                    ));
                }
            }
        }
        xml.push_str("</xdr:wsDr>");
        rels.push_str("</Relationships>");
        self.parts.push((format!("xl/drawings/drawing{d}.xml"), xml));
        self.parts.push((format!("xl/drawings/_rels/drawing{d}.xml.rels"), rels));
        self.content_types
            .push(format!(r#"<Override PartName="/xl/drawings/drawing{d}.xml" ContentType="{CT_DRAWING}"/>"#));

        // The worksheet points at its drawing through its own relationships part.
        let sheet_rels = format!("xl/worksheets/_rels/sheet{id}.xml.rels");
        let rel =
            format!(r#"<Relationship Id="rIdDrawing1" Type="{NS_REL}/drawing" Target="../drawings/drawing{d}.xml"/>"#);
        self.parts.push((
            sheet_rels,
            format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{NS_PKG_REL}">{rel}</Relationships>"#),
        ));
        insert_drawing_element(worksheet_xml)
    }

    /// Add the collected parts to the package, registering their content types.
    /// Sheet relationship parts that another feature already wrote are merged.
    pub(super) fn finish(self, parts: &mut Vec<(String, String)>) -> Vec<(String, Arc<[u8]>)> {
        let mut defaults = String::new();
        for ext in &self.extensions {
            let mime = match *ext {
                "jpeg" => "image/jpeg",
                "gif" => "image/gif",
                "bmp" => "image/bmp",
                _ => "image/png",
            };
            defaults.push_str(&format!(r#"<Default Extension="{ext}" ContentType="{mime}"/>"#));
        }
        if let Some((_, ct)) = parts.iter_mut().find(|(name, _)| name == "[Content_Types].xml") {
            // Defaults must come before Overrides.
            let pos = ct.find("<Override").unwrap_or_else(|| ct.rfind("</Types>").unwrap_or(ct.len()));
            ct.insert_str(pos, &defaults);
            let end = ct.rfind("</Types>").unwrap_or(ct.len());
            ct.insert_str(end, &self.content_types.concat());
        }
        for (name, xml) in self.parts {
            match parts.iter_mut().find(|(n, _)| *n == name) {
                Some((_, existing)) if name.ends_with(".rels") => {
                    // Merge our relationship into a rels part another writer created.
                    let rel_start = xml.find("<Relationship ").unwrap_or(xml.len());
                    let rel_end = xml.rfind("</Relationships>").unwrap_or(xml.len());
                    let end = existing.rfind("</Relationships>").unwrap_or(existing.len());
                    existing.insert_str(end, &xml[rel_start..rel_end]);
                }
                _ => parts.push((name, xml)),
            }
        }
        self.media
    }
}

/// Insert `<drawing r:id="rIdDrawing1"/>` where the schema wants it (after
/// most worksheet elements, before legacy drawings, tables and extensions),
/// declaring the relationships namespace on the root if needed.
fn insert_drawing_element(mut xml: String) -> String {
    if let Some(root) = xml.find("<worksheet") {
        let tag_end = xml[root..].find('>').map(|i| root + i).unwrap_or(xml.len());
        if !xml[root..tag_end].contains("xmlns:r=") {
            xml.insert_str(tag_end, &format!(r#" xmlns:r="{NS_REL}""#));
        }
    }
    let element = r#"<drawing r:id="rIdDrawing1"/>"#;
    let before = [
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
    let pos = before.iter().filter_map(|tag| xml.find(tag)).min().unwrap_or(xml.len());
    xml.insert_str(pos, element);
    xml
}

fn emu(px: f64) -> i64 {
    (px * EMU_PER_PX).round() as i64
}

fn anchor_xml(a: &Anchor) -> String {
    let point = |tag: &str, p: &AnchorPoint| {
        format!(
            "<xdr:{tag}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{tag}>",
            p.col,
            emu(p.col_off),
            p.row,
            emu(p.row_off)
        )
    };
    format!("{}{}", point("from", &a.from), point("to", &a.to))
}

/// Approximate size of an anchored object in EMU (for picture extents).
fn anchor_size_emu(sheet: &Sheet, a: &Anchor) -> (i64, i64) {
    let x = |col: u32, off: f64| (0..col).map(|c| sheet.col_width(c)).sum::<f64>() + off;
    let w = x(a.to.col, a.to.col_off) - x(a.from.col, a.from.col_off);
    let h = (a.to.row as f64 * ROW_PX + a.to.row_off) - (a.from.row as f64 * ROW_PX + a.from.row_off);
    (emu(w.max(1.0)), emu(h.max(1.0)))
}

struct Lookup<'a> {
    book: &'a Workbook,
    own: usize,
}

impl SheetLookup for Lookup<'_> {
    fn sheet_of(&self, name: Option<&str>) -> Option<&Sheet> {
        match name {
            None => self.book.sheets.get(self.own).map(|s| &s.sheet),
            Some(n) => self.book.sheets.iter().find(|s| s.name.eq_ignore_ascii_case(n)).map(|s| &s.sheet),
        }
    }
}

/// `<c:txPr>` with a font size (hundredths of a point) and colour.
fn text_props(size: u32, color: &str, rot: Option<i32>) -> String {
    let rot = rot.map_or(String::new(), |r| format!(r#" rot="{r}" vert="horz""#));
    format!(
        concat!(
            r#"<c:txPr><a:bodyPr{rot}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{size}" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">"#,
            r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>"#
        ),
        rot = rot,
        size = size,
        color = color
    )
}

/// A `<c:title>` with fixed text, or an automatic one when `text` is `None`.
fn title_xml(text: Option<&str>, size: u32, rot: Option<i32>) -> String {
    let rot_attr = rot.map_or(String::new(), |r| format!(r#" rot="{r}" vert="horz""#));
    let tx = text.map_or(String::new(), |t| {
        format!(
            concat!(
                r#"<c:tx><c:rich><a:bodyPr{rot}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{size}" b="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:defRPr></a:pPr>"#,
                r#"<a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx>"#
            ),
            rot = rot_attr,
            size = size,
            color = TEXT_GREY,
            text = escape(t)
        )
    });
    format!(
        r#"<c:title>{tx}<c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{}</c:title>"#,
        text_props(size, TEXT_GREY, rot)
    )
}

fn line_props(color: &str, width_emu: u32) -> String {
    format!(
        r#"<c:spPr><a:noFill/><a:ln w="{width_emu}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:round/></a:ln></c:spPr>"#
    )
}

fn str_ref(r: &DataRef, own: &str, texts: &[String]) -> String {
    let mut cache = format!(r#"<c:strCache><c:ptCount val="{}"/>"#, texts.len());
    for (i, t) in texts.iter().enumerate() {
        if !t.is_empty() {
            cache.push_str(&format!(r#"<c:pt idx="{i}"><c:v>{}</c:v></c:pt>"#, escape(t)));
        }
    }
    cache.push_str("</c:strCache>");
    format!("<c:strRef><c:f>{}</c:f>{cache}</c:strRef>", escape(&r.to_formula(own)))
}

fn num_ref(r: &DataRef, own: &str, values: &[Option<f64>]) -> String {
    let mut cache = format!(r#"<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{}"/>"#, values.len());
    for (i, v) in values.iter().enumerate() {
        if let Some(v) = v {
            cache.push_str(&format!(r#"<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>"#));
        }
    }
    cache.push_str("</c:numCache>");
    format!("<c:numRef><c:f>{}</c:f>{cache}</c:numRef>", escape(&r.to_formula(own)))
}

fn data_labels_xml(show: bool, pos: Option<&str>) -> String {
    if !show {
        return String::from(
            r#"<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#,
        );
    }
    let pos = pos.map_or(String::new(), |p| format!(r#"<c:dLblPos val="{p}"/>"#));
    format!(
        concat!(
            r#"<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{}{}"#,
            r#"<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>"#
        ),
        text_props(900, "404040", None),
        pos
    )
}

/// The `chartN.xml` part for one chart.
pub(super) fn chart_xml(chart: &Chart, own: &str, sheets: &dyn SheetLookup) -> String {
    let t = chart.chart_type;
    let view = chart::compute_view(chart, sheets, 200.0, 400.0);
    let mut series_xml = String::new();
    for (i, (s, v)) in chart.series.iter().zip(&view.series).enumerate() {
        let color = v.color.trim_start_matches('#');
        let mut x = format!(r#"<c:ser><c:idx val="{i}"/><c:order val="{i}"/>"#);
        match &s.name {
            SeriesName::Ref(r) => {
                x.push_str(&format!("<c:tx>{}</c:tx>", str_ref(r, own, std::slice::from_ref(&v.name))))
            }
            SeriesName::Text(text) => x.push_str(&format!("<c:tx><c:v>{}</c:v></c:tx>", escape(text))),
            SeriesName::None => {}
        }
        let solid = format!(r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#);
        let marker = |symbol: &str| {
            if symbol == "none" {
                r#"<c:marker><c:symbol val="none"/></c:marker>"#.to_string()
            } else {
                format!(
                    r#"<c:marker><c:symbol val="{symbol}"/><c:size val="5"/><c:spPr>{solid}<a:ln w="9525">{solid}</a:ln></c:spPr></c:marker>"#
                )
            }
        };
        match t {
            ChartType::Line | ChartType::LineMarkers => {
                x.push_str(&format!(r#"<c:spPr><a:ln w="28575" cap="rnd">{solid}<a:round/></a:ln></c:spPr>"#));
                x.push_str(&marker(if t == ChartType::LineMarkers { "circle" } else { "none" }));
            }
            ChartType::Scatter => {
                x.push_str(r#"<c:spPr><a:ln w="25400" cap="rnd"><a:noFill/><a:round/></a:ln></c:spPr>"#);
                x.push_str(&marker("circle"));
            }
            ChartType::ScatterLines => {
                x.push_str(&format!(r#"<c:spPr><a:ln w="19050" cap="rnd">{solid}<a:round/></a:ln></c:spPr>"#));
                x.push_str(&marker("none"));
            }
            ChartType::Pie | ChartType::Doughnut => {
                x.push_str(
                    r#"<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>"#,
                );
                // One colour per slice.
                for (p, c) in view.point_colors.iter().enumerate() {
                    x.push_str(&format!(
                        r#"<c:dPt><c:idx val="{p}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="{}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>"#,
                        c.trim_start_matches('#')
                    ));
                }
            }
            ChartType::Area | ChartType::AreaStacked => {
                x.push_str(&format!(r#"<c:spPr>{solid}<a:ln><a:noFill/></a:ln></c:spPr>"#));
            }
            _ => {
                x.push_str(&format!(
                    r#"<c:spPr>{solid}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>"#
                ));
            }
        }
        if chart.data_labels {
            let pos = match t {
                ChartType::ColumnClustered | ChartType::BarClustered => Some("outEnd"),
                ChartType::Line | ChartType::LineMarkers | ChartType::Scatter | ChartType::ScatterLines => Some("r"),
                ChartType::Pie => Some("bestFit"),
                ChartType::ColumnStacked | ChartType::ColumnPercent | ChartType::BarStacked | ChartType::BarPercent => {
                    Some("ctr")
                }
                _ => None,
            };
            x.push_str(&data_labels_xml(true, pos));
        }
        let (cat_tag, val_tag) = if t.is_scatter() { ("c:xVal", "c:yVal") } else { ("c:cat", "c:val") };
        if let Some(cats) = &s.categories {
            if t.is_scatter() {
                x.push_str(&format!("<{cat_tag}>{}</{cat_tag}>", num_ref(cats, own, &v.xs)));
            } else {
                x.push_str(&format!("<{cat_tag}>{}</{cat_tag}>", str_ref(cats, own, &view.categories)));
            }
        }
        x.push_str(&format!("<{val_tag}>{}</{val_tag}>", num_ref(&s.values, own, &v.values)));
        if matches!(t, ChartType::Line | ChartType::LineMarkers | ChartType::Scatter | ChartType::ScatterLines) {
            x.push_str(r#"<c:smooth val="0"/>"#);
        }
        x.push_str("</c:ser>");
        series_xml.push_str(&x);
    }

    let grouping = match t.grouping() {
        Grouping::Standard if matches!(t, ChartType::ColumnClustered | ChartType::BarClustered) => "clustered",
        Grouping::Standard => "standard",
        Grouping::Stacked => "stacked",
        Grouping::Percent => "percentStacked",
    };
    let chart_dlbls = data_labels_xml(false, None);
    let axes = r#"<c:axId val="500000001"/><c:axId val="500000002"/>"#;
    let group = match t {
        ChartType::ColumnClustered
        | ChartType::ColumnStacked
        | ChartType::ColumnPercent
        | ChartType::BarClustered
        | ChartType::BarStacked
        | ChartType::BarPercent => {
            let dir = if matches!(t, ChartType::BarClustered | ChartType::BarStacked | ChartType::BarPercent) {
                "bar"
            } else {
                "col"
            };
            let (gap, overlap) = if grouping == "clustered" { (219, -27) } else { (150, 100) };
            format!(
                r#"<c:barChart><c:barDir val="{dir}"/><c:grouping val="{grouping}"/><c:varyColors val="0"/>{series_xml}{chart_dlbls}<c:gapWidth val="{gap}"/><c:overlap val="{overlap}"/>{axes}</c:barChart>"#
            )
        }
        ChartType::Line | ChartType::LineMarkers => format!(
            r#"<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{series_xml}{chart_dlbls}<c:marker val="1"/><c:smooth val="0"/>{axes}</c:lineChart>"#
        ),
        ChartType::Pie => format!(
            r#"<c:pieChart><c:varyColors val="1"/>{series_xml}{chart_dlbls}<c:firstSliceAng val="0"/></c:pieChart>"#
        ),
        ChartType::Doughnut => format!(
            r#"<c:doughnutChart><c:varyColors val="1"/>{series_xml}{chart_dlbls}<c:firstSliceAng val="0"/><c:holeSize val="75"/></c:doughnutChart>"#
        ),
        ChartType::Area | ChartType::AreaStacked => format!(
            r#"<c:areaChart><c:grouping val="{grouping}"/><c:varyColors val="0"/>{series_xml}{chart_dlbls}{axes}</c:areaChart>"#
        ),
        ChartType::Scatter | ChartType::ScatterLines => format!(
            r#"<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{series_xml}{chart_dlbls}{axes}</c:scatterChart>"#
        ),
    };

    let bar = matches!(t, ChartType::BarClustered | ChartType::BarStacked | ChartType::BarPercent);
    let (cat_pos, val_pos) = if bar { ("l", "b") } else { ("b", "l") };
    let axis_text = text_props(900, TEXT_GREY, None);
    let gridlines = format!("<c:majorGridlines>{}</c:majorGridlines>", line_props(LINE_GREY, 9525));
    let axis_title = |title: &Option<String>, vertical: bool| {
        title.as_deref().map_or(String::new(), |s| title_xml(Some(s), 1000, vertical.then_some(-5_400_000)))
    };
    // x_title belongs to the category axis (horizontal unless this is a bar chart).
    let cat_title = axis_title(&chart.x_title, bar);
    let val_title = axis_title(&chart.y_title, !bar);
    let val_fmt = if t.grouping() == Grouping::Percent { "0%" } else { "General" };
    let axes_xml = if t.is_round() {
        String::new()
    } else if t.is_scatter() {
        format!(
            concat!(
                r#"<c:valAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>{grid}{xt}"#,
                r#"<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>{line}{text}"#,
                r#"<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>"#,
                r#"<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>{grid}{yt}"#,
                r#"<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{text}"#,
                r#"<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>"#
            ),
            grid = gridlines,
            xt = cat_title,
            yt = val_title,
            line = line_props(LINE_GREY, 9525),
            text = axis_text
        )
    } else {
        let between = if matches!(t, ChartType::Area | ChartType::AreaStacked) { "midCat" } else { "between" };
        format!(
            concat!(
                r#"<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{cat_pos}"/>{ct}"#,
                r#"<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>{line}{text}"#,
                r#"<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>"#,
                r#"<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{val_pos}"/>{grid}{vt}"#,
                r#"<c:numFmt formatCode="{val_fmt}" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{text}"#,
                r#"<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="{between}"/></c:valAx>"#
            ),
            cat_pos = cat_pos,
            val_pos = val_pos,
            ct = cat_title,
            vt = val_title,
            grid = gridlines,
            line = line_props(LINE_GREY, 9525),
            text = axis_text,
            val_fmt = val_fmt,
            between = between
        )
    };

    let title = if chart.show_title {
        format!(r#"{}<c:autoTitleDeleted val="0"/>"#, title_xml(chart.title.as_deref(), 1400, None))
    } else {
        r#"<c:autoTitleDeleted val="1"/>"#.to_string()
    };
    let legend = match chart.legend {
        LegendPos::None => String::new(),
        pos => format!(
            r#"<c:legend><c:legendPos val="{}"/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>{}</c:legend>"#,
            match pos {
                LegendPos::Right => "r",
                LegendPos::Top => "t",
                LegendPos::Left => "l",
                _ => "b",
            },
            text_props(900, TEXT_GREY, None)
        ),
    };
    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<c:chartSpace xmlns:c="{c}" xmlns:a="{a}" xmlns:r="{r}"><c:date1904 val="0"/><c:roundedCorners val="0"/>"#,
            r#"<c:chart>{title}<c:plotArea><c:layout/>{group}{axes}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>{legend}"#,
            r#"<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>"#,
            r#"<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill><a:round/></a:ln></c:spPr>"#,
            r#"<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>"#,
            r#"<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>"#,
            r#"</c:chartSpace>"#
        ),
        c = NS_C,
        a = NS_A,
        r = NS_REL,
        title = title,
        group = group,
        axes = axes_xml,
        legend = legend,
        line = LINE_GREY
    )
}

// ---------------------------------------------------------------- reading

fn part_bytes(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<Option<Vec<u8>>> {
    let mut file = match zip.by_name(name) {
        Ok(f) => f,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(e) => return Err(OfficeError::Parse(e.to_string())),
    };
    let mut out = Vec::new();
    file.read_to_end(&mut out).map_err(|e| OfficeError::Io(e.to_string()))?;
    Ok(Some(out))
}

fn part_text(zip: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<Option<String>> {
    Ok(part_bytes(zip, name)?.map(|b| String::from_utf8_lossy(&b).into_owned()))
}

/// `xl/worksheets/sheet1.xml` -> `xl/worksheets/_rels/sheet1.xml.rels`.
fn rels_path(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

/// Resolve a relationship target relative to the part that references it.
fn resolve(from_part: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') {
        return abs.to_string();
    }
    let mut dir: Vec<&str> = from_part.split('/').collect();
    dir.pop();
    for seg in target.split('/') {
        match seg {
            ".." => {
                dir.pop();
            }
            "." | "" => {}
            s => dir.push(s),
        }
    }
    dir.join("/")
}

/// Read the charts and pictures of one worksheet part into `sheet`.
/// Problems with drawings never fail the whole file: they are skipped.
pub(super) fn read_sheet_drawings(
    zip: &mut ZipArchive<Cursor<&[u8]>>,
    sheet_part: &str,
    sheet_name: &str,
    sheet: &mut Sheet,
) -> Result<()> {
    let Some(rels_xml) = part_text(zip, &rels_path(sheet_part))? else { return Ok(()) };
    let rels = relationships(&rels_xml);
    let mut items = Vec::new();
    for rel in rels.iter().filter(|r| r.kind.ends_with("/drawing")) {
        let drawing_part = resolve(sheet_part, &rel.target);
        let Some(xml) = part_text(zip, &drawing_part)? else { continue };
        let drawing_rels: HashMap<String, String> = part_text(zip, &rels_path(&drawing_part))?
            .map(|x| relationships(&x).into_iter().map(|r| (r.id, resolve(&drawing_part, &r.target))).collect())
            .unwrap_or_default();
        for anchor in parse_anchors(&xml, sheet) {
            let content = match &anchor.target {
                Target::Chart(rid) => {
                    let Some(path) = drawing_rels.get(rid) else { continue };
                    let Some(chart_xml) = part_text(zip, path)? else { continue };
                    match parse_chart(&chart_xml, sheet_name) {
                        Some(c) => DrawingContent::Chart(Box::new(c)),
                        None => continue,
                    }
                }
                Target::Picture(rid) => {
                    let Some(path) = drawing_rels.get(rid) else { continue };
                    let ext = path.rsplit('.').next().unwrap_or("");
                    let Some(mime) = Image::mime_for_extension(ext) else { continue };
                    let Some(bytes) = part_bytes(zip, path)? else { continue };
                    DrawingContent::Picture(Image { mime: mime.to_string(), data: Arc::from(bytes) })
                }
            };
            items.push((anchor.name, anchor.anchor, content));
        }
    }
    if !items.is_empty() {
        let mut drawings = Drawings::default();
        for (name, anchor, content) in items {
            let id = drawings.next_id();
            let name = if name.is_empty() {
                let kind = if matches!(content, DrawingContent::Chart(_)) { "Chart" } else { "Picture" };
                format!("{kind} {id}")
            } else {
                name
            };
            drawings.items.push(Drawing { id, name, anchor, content });
        }
        sheet.set_drawings(drawings);
    }
    Ok(())
}

enum Target {
    Chart(String),
    Picture(String),
}

struct ParsedAnchor {
    anchor: Anchor,
    name: String,
    target: Target,
}

/// Anchors from a drawing part. Two-cell anchors are used as they are;
/// one-cell and absolute anchors are converted using the sheet's column widths.
fn parse_anchors(xml: &str, sheet: &Sheet) -> Vec<ParsedAnchor> {
    #[derive(Default)]
    struct Cur {
        kind: String,
        from: [f64; 4],
        to: [f64; 4],
        pos: (f64, f64),
        ext: (f64, f64),
        name: String,
        chart: Option<String>,
        picture: Option<String>,
    }
    let mut out = Vec::new();
    let mut cur: Option<Cur> = None;
    let mut corner = "";
    let mut field = "";
    let mut depth_in_group = 0;
    walk(xml, |ev| match ev {
        XmlEvent::Start { name: n @ ("twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor"), .. } => {
            cur = Some(Cur { kind: n.to_string(), ..Cur::default() });
        }
        XmlEvent::Start { name: "grpSp", .. } => depth_in_group += 1,
        XmlEvent::End { name: "grpSp" } => depth_in_group -= 1,
        XmlEvent::Start { name: n @ ("from" | "to"), .. } => corner = if *n == "from" { "from" } else { "to" },
        XmlEvent::End { name: "from" | "to" } => corner = "",
        XmlEvent::Start { name: n @ ("col" | "colOff" | "row" | "rowOff"), .. } if !corner.is_empty() => {
            field = match *n {
                "col" => "col",
                "colOff" => "colOff",
                "row" => "row",
                _ => "rowOff",
            }
        }
        XmlEvent::End { name: "col" | "colOff" | "row" | "rowOff" } => field = "",
        XmlEvent::Text(text) if !field.is_empty() => {
            if let (Some(c), Ok(v)) = (cur.as_mut(), text.trim().parse::<f64>()) {
                let slot = match field {
                    "col" => 0,
                    "colOff" => 1,
                    "row" => 2,
                    _ => 3,
                };
                if corner == "from" {
                    c.from[slot] = v;
                } else {
                    c.to[slot] = v;
                }
            }
        }
        XmlEvent::Start { name: "pos", attrs } => {
            if let Some(c) = cur.as_mut() {
                c.pos = (attr(attrs, "x").parse().unwrap_or(0.0), attr(attrs, "y").parse().unwrap_or(0.0));
            }
        }
        XmlEvent::Start { name: "ext", attrs } if depth_in_group == 0 => {
            if let Some(c) = cur.as_mut() {
                if c.ext == (0.0, 0.0) {
                    if let (Ok(cx), Ok(cy)) = (attr(attrs, "cx").parse::<f64>(), attr(attrs, "cy").parse::<f64>()) {
                        c.ext = (cx, cy);
                    }
                }
            }
        }
        XmlEvent::Start { name: "cNvPr", attrs } => {
            if let Some(c) = cur.as_mut() {
                if c.name.is_empty() {
                    c.name = attr(attrs, "name").to_string();
                }
            }
        }
        XmlEvent::Start { name: "chart", attrs } => {
            if let Some(c) = cur.as_mut() {
                c.chart = Some(attr(attrs, "id").to_string());
            }
        }
        XmlEvent::Start { name: "blip", attrs } => {
            if let Some(c) = cur.as_mut() {
                c.picture = Some(attr(attrs, "embed").to_string());
            }
        }
        XmlEvent::End { name: "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" } => {
            let Some(c) = cur.take() else { return };
            // Grouped shapes and other objects are skipped.
            if depth_in_group > 0 {
                return;
            }
            let target = match (c.chart, c.picture) {
                (Some(r), _) => Target::Chart(r),
                (None, Some(r)) => Target::Picture(r),
                _ => return,
            };
            let px = |e: f64| e / EMU_PER_PX;
            let point = |v: [f64; 4]| AnchorPoint::new(v[2] as u32, v[0] as u32, px(v[3]), px(v[1]));
            let anchor = match c.kind.as_str() {
                "twoCellAnchor" => Anchor { from: point(c.from), to: point(c.to) },
                "oneCellAnchor" => {
                    let from = point(c.from);
                    let to = point_at(sheet, from, px(c.ext.0), px(c.ext.1));
                    Anchor { from, to }
                }
                _ => {
                    let origin = AnchorPoint::new(0, 0, 0.0, 0.0);
                    let from = point_at(sheet, origin, px(c.pos.0), px(c.pos.1));
                    let to = point_at(sheet, from, px(c.ext.0), px(c.ext.1));
                    Anchor { from, to }
                }
            };
            out.push(ParsedAnchor { anchor, name: c.name, target });
        }
        _ => {}
    });
    out
}

/// The anchor point `dx`, `dy` pixels right of and below `from`.
fn point_at(sheet: &Sheet, from: AnchorPoint, dx: f64, dy: f64) -> AnchorPoint {
    let mut col = from.col;
    let mut x = from.col_off + dx;
    while x >= sheet.col_width(col) && col + 1 < sheets_engine::MAX_COLS {
        x -= sheet.col_width(col);
        col += 1;
    }
    let total_y = from.row_off + dy;
    let rows = (total_y / ROW_PX).floor();
    AnchorPoint::new(from.row + rows as u32, col, total_y - rows * ROW_PX, x)
}

/// Read a `chartN.xml` part. `None` for chart types we don't support.
pub(super) fn parse_chart(xml: &str, own_sheet: &str) -> Option<Chart> {
    #[derive(Default)]
    struct Ser {
        name_ref: Option<String>,
        name_text: Option<String>,
        cat: Option<String>,
        val: Option<String>,
        color: Option<String>,
        no_marker: bool,
        no_line: bool,
        labels: bool,
    }
    let mut stack: Vec<String> = Vec::new();
    let mut group: Option<String> = None;
    let (mut bar_dir, mut grouping) = (String::from("col"), String::from("clustered"));
    let mut series: Vec<Ser> = Vec::new();
    let mut title_seen = false;
    let mut title_text = String::new();
    let mut auto_title_deleted = false;
    let mut legend = LegendPos::None;
    let mut group_labels = false;
    let mut axis_titles: Vec<(String, String, String)> = Vec::new(); // (axis kind, axPos, title)
    let mut text = String::new();
    let mut in_ser = false;

    let has = |stack: &[String], tail: &[&str]| {
        stack.len() >= tail.len() && stack[stack.len() - tail.len()..].iter().zip(tail).all(|(a, b)| a == b)
    };

    walk(xml, |ev| match ev {
        XmlEvent::Start { name, attrs } => {
            let name = name.to_string();
            let val = attr(attrs, "val").to_string();
            let parent = stack.last().map(String::as_str).unwrap_or("");
            match name.as_str() {
                "barChart" | "lineChart" | "pieChart" | "doughnutChart" | "areaChart" | "scatterChart"
                | "bar3DChart" | "line3DChart" | "pie3DChart" | "area3DChart" | "radarChart" | "bubbleChart"
                | "stockChart" | "surfaceChart" | "ofPieChart" => {
                    if group.is_none() {
                        group = Some(name.clone());
                    }
                }
                "barDir" => bar_dir = val.clone(),
                "grouping" if group.is_some() && parent.ends_with("Chart") => grouping = val.clone(),
                "ser" if group.as_deref().is_some_and(|g| parent == g) => {
                    in_ser = true;
                    series.push(Ser::default());
                }
                "title" if parent == "chart" => title_seen = true,
                "autoTitleDeleted" => auto_title_deleted = val == "1" || val == "true",
                "legendPos" => {
                    legend = match val.as_str() {
                        "r" | "tr" => LegendPos::Right,
                        "t" => LegendPos::Top,
                        "l" => LegendPos::Left,
                        _ => LegendPos::Bottom,
                    }
                }
                "legend" => {
                    if legend == LegendPos::None {
                        legend = LegendPos::Right; // Excel's default when legendPos is missing
                    }
                }
                "showVal" if val == "1" || val == "true" => {
                    if in_ser {
                        if let Some(s) = series.last_mut() {
                            s.labels = true;
                        }
                    } else if stack.iter().any(|s| s.ends_with("Chart")) {
                        group_labels = true;
                    }
                }
                "srgbClr" if in_ser => {
                    let fill = has(&stack, &["ser", "spPr", "solidFill"]);
                    let line = has(&stack, &["ser", "spPr", "ln", "solidFill"])
                        || has(&stack, &["ser", "marker", "spPr", "solidFill"]);
                    if let Some(s) = series.last_mut() {
                        if (fill || line) && s.color.is_none() {
                            s.color = Some(format!("#{}", val.to_ascii_uppercase()));
                        }
                    }
                }
                "symbol" if in_ser && val == "none" => {
                    if let Some(s) = series.last_mut() {
                        s.no_marker = true;
                    }
                }
                "noFill" if in_ser && has(&stack, &["ser", "spPr", "ln"]) => {
                    if let Some(s) = series.last_mut() {
                        s.no_line = true;
                    }
                }
                "axPos" => {
                    if let Some(kind) = stack.last() {
                        axis_titles.push((kind.clone(), val.clone(), String::new()));
                    }
                }
                _ => {}
            }
            stack.push(name);
            text.clear();
        }
        XmlEvent::Text(t) => text.push_str(t),
        XmlEvent::End { name } => {
            let name = name.to_string();
            match name.as_str() {
                "f" if in_ser => {
                    let f = text.trim().to_string();
                    if let Some(s) = series.last_mut() {
                        if has(&stack, &["tx", "strRef", "f"]) {
                            s.name_ref = Some(f);
                        } else if has(&stack, &["cat", "strRef", "f"])
                            || has(&stack, &["cat", "numRef", "f"])
                            || has(&stack, &["xVal", "numRef", "f"])
                            || has(&stack, &["xVal", "strRef", "f"])
                            || has(&stack, &["cat", "multiLvlStrRef", "f"])
                        {
                            s.cat = Some(f);
                        } else if has(&stack, &["val", "numRef", "f"]) || has(&stack, &["yVal", "numRef", "f"]) {
                            s.val = Some(f);
                        }
                    }
                }
                "v" if in_ser && has(&stack, &["ser", "tx", "v"]) => {
                    if let Some(s) = series.last_mut() {
                        s.name_text = Some(text.clone());
                    }
                }
                "t" => {
                    // Rich text runs: the chart title or an axis title.
                    let in_rich = stack.iter().any(|s| s == "rich");
                    if in_rich && stack.iter().any(|s| s == "title") {
                        let axis = stack.iter().rev().find(|s| *s == "catAx" || *s == "valAx" || *s == "dateAx");
                        match axis {
                            Some(_) => {
                                if let Some(last) = axis_titles.last_mut() {
                                    last.2.push_str(&text);
                                }
                            }
                            None if stack.len() >= 2 && stack.iter().any(|s| s == "chart") => {
                                title_text.push_str(&text)
                            }
                            None => {}
                        }
                    }
                }
                "ser" => in_ser = false,
                _ => {}
            }
            stack.pop();
            text.clear();
        }
    });

    let group = group?;
    let markers_off = !series.is_empty() && series.iter().all(|s| s.no_marker);
    let chart_type = match group.as_str() {
        "barChart" | "bar3DChart" => match (bar_dir.as_str(), grouping.as_str()) {
            ("bar", "stacked") => ChartType::BarStacked,
            ("bar", "percentStacked") => ChartType::BarPercent,
            ("bar", _) => ChartType::BarClustered,
            (_, "stacked") => ChartType::ColumnStacked,
            (_, "percentStacked") => ChartType::ColumnPercent,
            _ => ChartType::ColumnClustered,
        },
        "lineChart" | "line3DChart" if markers_off => ChartType::Line,
        "lineChart" | "line3DChart" => ChartType::LineMarkers,
        "pieChart" | "pie3DChart" | "ofPieChart" => ChartType::Pie,
        "doughnutChart" => ChartType::Doughnut,
        "areaChart" | "area3DChart" if grouping == "stacked" || grouping == "percentStacked" => ChartType::AreaStacked,
        "areaChart" | "area3DChart" => ChartType::Area,
        "scatterChart" if series.iter().all(|s| s.no_line) => ChartType::Scatter,
        "scatterChart" => ChartType::ScatterLines,
        _ => return None,
    };
    let local = |f: &str| -> Option<DataRef> {
        let mut r = DataRef::parse(f)?;
        if r.sheet.as_deref().is_some_and(|s| s.eq_ignore_ascii_case(own_sheet)) {
            r.sheet = None;
        }
        Some(r)
    };
    let mut out_series = Vec::new();
    let mut any_labels = group_labels;
    for s in &series {
        let Some(values) = s.val.as_deref().and_then(local) else { continue };
        any_labels |= s.labels;
        out_series.push(Series {
            name: match (&s.name_ref, &s.name_text) {
                (Some(f), _) => local(f).map_or(SeriesName::None, SeriesName::Ref),
                (None, Some(t)) => SeriesName::Text(t.clone()),
                _ => SeriesName::None,
            },
            categories: s.cat.as_deref().and_then(local),
            values,
            // Palette colours are written explicitly; read them back as automatic.
            color: if chart_type.is_round() {
                None
            } else {
                s.color.clone().filter(|c| !c.eq_ignore_ascii_case(chart::auto_color(out_series.len())))
            },
        });
    }
    if out_series.is_empty() {
        return None;
    }
    let mut c = Chart::new(chart_type, out_series);
    c.legend = legend;
    c.data_labels = any_labels;
    c.show_title = title_seen || (!auto_title_deleted && c.series.len() == 1);
    c.title = (!title_text.is_empty()).then_some(title_text);
    for (kind, pos, title) in axis_titles.into_iter().filter(|(_, _, t)| !t.is_empty()) {
        let is_category =
            if chart_type.is_scatter() { pos == "b" || pos == "t" } else { kind == "catAx" || kind == "dateAx" };
        if is_category {
            c.x_title = Some(title);
        } else {
            c.y_title = Some(title);
        }
    }
    c.source = infer_source(&c.series);
    c.by_rows = c.series.first().is_some_and(|s| s.values.range.rows() == 1 && s.values.range.cols() > 1);
    Some(c)
}

/// The block of cells covering all series names, categories and values, if
/// they are on the chart's own sheet and form exactly that block (so "Switch
/// Row/Column" can rebuild them).
fn infer_source(series: &[Series]) -> Option<DataRef> {
    let mut refs: Vec<&DataRef> = Vec::new();
    for s in series {
        refs.push(&s.values);
        if let Some(c) = &s.categories {
            refs.push(c);
        }
        if let SeriesName::Ref(r) = &s.name {
            refs.push(r);
        }
    }
    if refs.iter().any(|r| r.sheet.is_some()) {
        return None;
    }
    let top = refs.iter().map(|r| r.range.start.row).min()?;
    let left = refs.iter().map(|r| r.range.start.col).min()?;
    let bottom = refs.iter().map(|r| r.range.end.row).max()?;
    let right = refs.iter().map(|r| r.range.end.col).max()?;
    let cells: usize = refs.iter().map(|r| (r.range.rows() * r.range.cols()) as usize).sum();
    let area = ((bottom - top + 1) * (right - left + 1)) as usize;
    // Categories are shared by every series, so they may be counted several
    // times; the top-left corner cell is usually not referenced at all.
    (cells + 1 >= area && area <= 100_000).then(|| {
        DataRef::local(sheets_engine::Range::new(
            sheets_engine::CellRef::new(top, left),
            sheets_engine::CellRef::new(bottom, right),
        ))
    })
}

#[cfg(test)]
mod tests {
    use super::super::{read, write};
    use super::*;
    use sheets_engine::chart::{chart_for_range, ChartPatch};
    use sheets_engine::{CellRef, Range};

    fn table_book() -> Workbook {
        let mut book = Workbook::new();
        book.sheets[0].name = "Sales Data".into();
        let s = &mut book.sheets[0].sheet;
        let cells = [
            ("A1", "Region"),
            ("B1", "Q1"),
            ("C1", "Q2"),
            ("A2", "North"),
            ("B2", "10"),
            ("C2", "20"),
            ("A3", "South & West"),
            ("B3", "30"),
            ("C3", "40"),
        ];
        s.set_inputs(cells.iter().map(|(a, v)| (CellRef::parse(a).unwrap(), v.to_string())));
        book
    }

    fn add_chart(book: &mut Workbook, t: ChartType) -> u32 {
        let s = &book.sheets[0].sheet;
        let chart = chart_for_range(s, Range::new(CellRef::new(0, 0), CellRef::new(2, 2)), t).unwrap();
        let mut d = s.drawings().clone();
        let anchor = Anchor { from: AnchorPoint::new(4, 1, 5.0, 10.0), to: AnchorPoint::new(18, 7, 0.0, 0.0) };
        let id = d.push(d.default_name("Chart"), anchor, DrawingContent::Chart(Box::new(chart)));
        book.sheets[0].sheet.set_drawings(d);
        id
    }

    fn unzip(bytes: &[u8], name: &str) -> String {
        let mut zip = ZipArchive::new(Cursor::new(bytes)).unwrap();
        part_text(&mut zip, name).unwrap().unwrap_or_default()
    }

    #[test]
    fn writes_excel_drawing_parts() {
        let mut book = table_book();
        add_chart(&mut book, ChartType::ColumnClustered);
        let bytes = write(&book).unwrap();
        let sheet = unzip(&bytes, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#));
        assert!(sheet.ends_with(r#"</sheetData><drawing r:id="rIdDrawing1"/></worksheet>"#), "{sheet}");
        let types = unzip(&bytes, "[Content_Types].xml");
        assert!(types.contains("/xl/drawings/drawing1.xml"));
        assert!(types.contains("/xl/charts/chart1.xml"));
        assert!(unzip(&bytes, "xl/worksheets/_rels/sheet1.xml.rels").contains("../drawings/drawing1.xml"));
        let drawing = unzip(&bytes, "xl/drawings/drawing1.xml");
        assert!(drawing.contains("<xdr:from><xdr:col>1</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>47625</xdr:rowOff></xdr:from>"));
        let chart = unzip(&bytes, "xl/charts/chart1.xml");
        assert!(chart.contains(r#"<c:barDir val="col"/><c:grouping val="clustered"/>"#));
        assert!(chart.contains("<c:f>'Sales Data'!$B$2:$B$3</c:f>"));
        assert!(chart.contains("<c:v>South &amp; West</c:v>"));
        assert!(chart.contains(r#"<c:legendPos val="b"/>"#));
        // Well-formed XML.
        for part in [&sheet, &drawing, &chart, &types] {
            let mut depth = 0i32;
            walk(part, |ev| match ev {
                XmlEvent::Start { .. } => depth += 1,
                XmlEvent::End { .. } => depth -= 1,
                _ => {}
            });
            assert_eq!(depth, 0);
        }
    }

    #[test]
    fn round_trips_every_chart_type() {
        for t in ChartType::ALL {
            let mut book = table_book();
            let id = add_chart(&mut book, t);
            let s = &book.sheets[0].sheet;
            let mut d = s.drawings().clone();
            let chart = d.get(id).unwrap().chart().unwrap().clone();
            let chart = chart.patched(&ChartPatch::Title(Some("Revenue <2024>".into())), s).unwrap();
            let chart = chart.patched(&ChartPatch::Legend(LegendPos::Right), s).unwrap();
            let chart = chart.patched(&ChartPatch::DataLabels(true), s).unwrap();
            let chart = chart.patched(&ChartPatch::XTitle(Some("Region".into())), s).unwrap();
            let chart = chart.patched(&ChartPatch::YTitle(Some("Amount".into())), s).unwrap();
            let chart = chart.patched(&ChartPatch::SeriesColor(0, Some("#112233".into())), s).unwrap();
            d.get_mut(id).unwrap().content = DrawingContent::Chart(Box::new(chart.clone()));
            book.sheets[0].sheet.set_drawings(d);

            let back = read(&write(&book).unwrap()).unwrap();
            let drawings = back.sheets[0].sheet.drawings();
            assert_eq!(drawings.items.len(), 1, "{t:?}");
            let item = &drawings.items[0];
            assert_eq!(item.name, "Chart 1");
            assert_eq!(item.anchor, book.sheets[0].sheet.drawings().items[0].anchor, "{t:?}");
            let got = item.chart().unwrap();
            assert_eq!(got.chart_type, t);
            assert_eq!(
                got.series,
                if t.is_round() {
                    chart.series.iter().map(|s| Series { color: None, ..s.clone() }).collect::<Vec<_>>()
                } else {
                    chart.series.clone()
                },
                "{t:?}"
            );
            assert_eq!(got.title.as_deref(), Some("Revenue <2024>"));
            assert!(got.show_title);
            assert_eq!(got.legend, LegendPos::Right);
            assert!(got.data_labels, "{t:?}");
            if !t.is_round() {
                assert_eq!(got.x_title.as_deref(), Some("Region"), "{t:?}");
                assert_eq!(got.y_title.as_deref(), Some("Amount"), "{t:?}");
            }
            assert_eq!(got.source, chart.source, "{t:?}");
        }
    }

    #[test]
    fn hidden_title_and_no_legend_round_trip() {
        let mut book = table_book();
        let id = add_chart(&mut book, ChartType::Line);
        let s = &book.sheets[0].sheet;
        let mut d = s.drawings().clone();
        let c = d.get(id).unwrap().chart().unwrap().clone();
        let c = c
            .patched(&ChartPatch::ShowTitle(false), s)
            .unwrap()
            .patched(&ChartPatch::Legend(LegendPos::None), s)
            .unwrap();
        d.get_mut(id).unwrap().content = DrawingContent::Chart(Box::new(c));
        book.sheets[0].sheet.set_drawings(d);
        let back = read(&write(&book).unwrap()).unwrap();
        let got = back.sheets[0].sheet.drawings().items[0].chart().unwrap().clone();
        assert!(!got.show_title);
        assert_eq!(got.legend, LegendPos::None);
        assert_eq!(got.title, None);
    }

    #[test]
    fn round_trips_pictures() {
        let mut book = table_book();
        let png: Vec<u8> = vec![0x89, b'P', b'N', b'G', 1, 2, 3];
        let mut d = Drawings::default();
        let a = Anchor::from_pixels_uniform(120.0, 30.0, 200.0, 100.0, 100.0, 22.0);
        d.push(
            "Picture 1".into(),
            a,
            DrawingContent::Picture(Image { mime: "image/png".into(), data: Arc::from(png.clone()) }),
        );
        let jpg = Image { mime: "image/jpeg".into(), data: Arc::from(vec![0xFF, 0xD8, 9]) };
        d.push("Logo".into(), a, DrawingContent::Picture(jpg));
        book.sheets[0].sheet.set_drawings(d);
        let bytes = write(&book).unwrap();
        assert!(unzip(&bytes, "[Content_Types].xml").contains(r#"<Default Extension="png" ContentType="image/png"/>"#));
        let back = read(&bytes).unwrap();
        let items = &back.sheets[0].sheet.drawings().items;
        assert_eq!(items.len(), 2);
        match &items[0].content {
            DrawingContent::Picture(img) => {
                assert_eq!(&*img.data, &png[..]);
                assert_eq!(img.mime, "image/png");
            }
            _ => panic!("expected a picture"),
        }
        assert_eq!(items[1].name, "Logo");
        assert_eq!(items[0].anchor, a);
    }

    #[test]
    fn reads_excel_written_chart_xml() {
        // Trimmed from a chart saved by Excel for Microsoft 365.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>
<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>
<c:marker><c:symbol val="none"/></c:marker>
<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val><c:smooth val="0"/></c:ser>
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>
<c:smooth val="0"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>
<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/></c:catAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>
<c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:p><a:r><a:t>Rupees</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title></c:valAx>
</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>"#;
        let c = parse_chart(xml, "Sheet1").unwrap();
        assert_eq!(c.chart_type, ChartType::Line);
        assert_eq!(c.series.len(), 1);
        assert_eq!(c.series[0].values, DataRef::parse("B2:B4").unwrap());
        assert_eq!(c.series[0].color, None, "theme colours fall back to the palette");
        assert!(c.show_title, "single series with automatic title");
        assert_eq!(c.title, None);
        assert_eq!(c.y_title.as_deref(), Some("Rupees"));
        assert_eq!(c.x_title, None);
        assert_eq!(c.legend, LegendPos::None);
        assert_eq!(c.source, DataRef::parse("A1:B4"));
        assert!(parse_chart(
            "<c:chartSpace><c:chart><c:plotArea><c:radarChart/></c:plotArea></c:chart></c:chartSpace>",
            "S"
        )
        .is_none());
    }

    #[test]
    fn resolves_relative_targets() {
        assert_eq!(resolve("xl/worksheets/sheet1.xml", "../drawings/drawing1.xml"), "xl/drawings/drawing1.xml");
        assert_eq!(resolve("xl/drawings/drawing1.xml", "/xl/media/image1.png"), "xl/media/image1.png");
        assert_eq!(rels_path("xl/drawings/drawing1.xml"), "xl/drawings/_rels/drawing1.xml.rels");
        let xml = insert_drawing_element(r#"<worksheet xmlns="x"><sheetData/><extLst/></worksheet>"#.to_string());
        assert!(xml.contains(r#"<sheetData/><drawing r:id="rIdDrawing1"/><extLst/>"#));
    }
}
