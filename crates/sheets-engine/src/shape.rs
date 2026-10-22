//! Shapes from Insert → Shapes: rectangles, ovals, arrows, stars, lines, text
//! boxes... Each is a DrawingML preset geometry (`prst` in `.xlsx`, the same
//! names Excel uses) with a fill, an outline and optional text. Shapes float
//! over the sheet like charts and pictures (see `chart.rs` for anchoring and
//! the undoable `SetDrawings` op).

/// Preset geometries offered in the Shapes gallery, in Excel's groups.
/// (group, preset, label)
pub const SHAPE_PRESETS: &[(&str, &str, &str)] = &[
    ("Lines", "line", "Line"),
    ("Lines", "straightConnector1", "Arrow"),
    ("Rectangles", "rect", "Rectangle"),
    ("Rectangles", "roundRect", "Rounded Rectangle"),
    ("Rectangles", "snip1Rect", "Snipped Rectangle"),
    ("Basic Shapes", "ellipse", "Oval"),
    ("Basic Shapes", "triangle", "Triangle"),
    ("Basic Shapes", "rtTriangle", "Right Triangle"),
    ("Basic Shapes", "parallelogram", "Parallelogram"),
    ("Basic Shapes", "trapezoid", "Trapezoid"),
    ("Basic Shapes", "diamond", "Diamond"),
    ("Basic Shapes", "pentagon", "Pentagon"),
    ("Basic Shapes", "hexagon", "Hexagon"),
    ("Basic Shapes", "octagon", "Octagon"),
    ("Basic Shapes", "plus", "Cross"),
    ("Basic Shapes", "donut", "Donut"),
    ("Basic Shapes", "heart", "Heart"),
    ("Basic Shapes", "smileyFace", "Smiley Face"),
    ("Basic Shapes", "can", "Cylinder"),
    ("Basic Shapes", "cube", "Cube"),
    ("Block Arrows", "rightArrow", "Right Arrow"),
    ("Block Arrows", "leftArrow", "Left Arrow"),
    ("Block Arrows", "upArrow", "Up Arrow"),
    ("Block Arrows", "downArrow", "Down Arrow"),
    ("Block Arrows", "leftRightArrow", "Left-Right Arrow"),
    ("Block Arrows", "chevron", "Chevron"),
    ("Block Arrows", "homePlate", "Pentagon Arrow"),
    ("Stars and Banners", "star5", "5-Point Star"),
    ("Stars and Banners", "star6", "6-Point Star"),
    ("Stars and Banners", "star8", "8-Point Star"),
    ("Stars and Banners", "irregularSeal1", "Explosion"),
    ("Callouts", "wedgeRectCallout", "Rectangular Callout"),
    ("Callouts", "wedgeRoundRectCallout", "Rounded Callout"),
    ("Callouts", "wedgeEllipseCallout", "Oval Callout"),
    ("Callouts", "cloudCallout", "Thought Bubble"),
    ("Flowchart", "flowChartProcess", "Process"),
    ("Flowchart", "flowChartDecision", "Decision"),
    ("Flowchart", "flowChartTerminator", "Terminator"),
    ("Flowchart", "flowChartDocument", "Document"),
    ("Flowchart", "flowChartInputOutput", "Data"),
];

/// Whether `preset` is a line (no fill, no text; drawn corner to corner).
pub fn is_line(preset: &str) -> bool {
    matches!(preset, "line" | "straightConnector1")
}

pub fn is_known_preset(preset: &str) -> bool {
    SHAPE_PRESETS.iter().any(|(_, p, _)| *p == preset)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TextAlign {
    Left,
    #[default]
    Center,
    Right,
}

impl TextAlign {
    pub fn id(self) -> &'static str {
        match self {
            TextAlign::Left => "left",
            TextAlign::Center => "center",
            TextAlign::Right => "right",
        }
    }

    pub fn from_id(s: &str) -> Option<TextAlign> {
        Some(match s {
            "left" => TextAlign::Left,
            "center" => TextAlign::Center,
            "right" => TextAlign::Right,
            _ => return None,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Shape {
    /// DrawingML preset geometry, e.g. `rect`, `ellipse`, `rightArrow`.
    pub preset: String,
    /// Fill `#RRGGBB`; `None` is No Fill.
    pub fill: Option<String>,
    /// Outline `#RRGGBB`; `None` is No Outline.
    pub line: Option<String>,
    /// Outline width in points.
    pub line_width: f64,
    /// Lines only: an arrowhead at the end.
    pub arrow_end: bool,
    /// Lines only: drawn bottom-left to top-right instead of top-left to bottom-right.
    pub flip_v: bool,
    pub text: String,
    pub text_color: String,
    /// Points.
    pub font_size: f64,
    pub bold: bool,
    pub italic: bool,
    pub align: TextAlign,
    /// Text boxes put text at the top; shapes centre it vertically.
    pub text_top: bool,
}

impl Shape {
    /// A shape with Excel's default style (Office theme: blue fill, darker
    /// blue outline, white 11pt text; lines in the accent colour).
    pub fn new(preset: &str) -> Shape {
        let line = is_line(preset);
        Shape {
            preset: preset.to_string(),
            fill: (!line).then(|| "#4472C4".to_string()),
            line: Some(if line { "#4472C4" } else { "#2F528F" }.to_string()),
            line_width: 1.0,
            arrow_end: preset == "straightConnector1",
            flip_v: false,
            text: String::new(),
            text_color: "#FFFFFF".into(),
            font_size: 11.0,
            bold: false,
            italic: false,
            align: TextAlign::Center,
            text_top: false,
        }
    }

    /// Insert → Text Box: white, thin dark outline, black text at the top left.
    pub fn text_box() -> Shape {
        Shape {
            fill: Some("#FFFFFF".into()),
            line: Some("#BFBFBF".into()),
            line_width: 0.75,
            text_color: "#000000".into(),
            align: TextAlign::Left,
            text_top: true,
            ..Shape::new("rect")
        }
    }

    /// One formatting change from the Shape Format tab. Returns `None` for an
    /// unknown key or a bad value. Keys: `fill` / `line` (`#RRGGBB`, "" = none),
    /// `lineWidth` (points), `text`, `textColor`, `fontSize`, `bold`, `italic`
    /// ("1"/"0"), `align` (left|center|right), `preset`, `arrowEnd`.
    pub fn patched(&self, key: &str, value: &str) -> Option<Shape> {
        let mut s = self.clone();
        let color = |v: &str| -> Option<Option<String>> {
            if v.is_empty() {
                return Some(None);
            }
            let hex = v.strip_prefix('#')?;
            (hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()))
                .then(|| Some(format!("#{}", hex.to_uppercase())))
        };
        let flag = |v: &str| v == "1" || v == "true";
        match key {
            "fill" => s.fill = color(value)?,
            "line" => s.line = color(value)?,
            "lineWidth" => s.line_width = value.parse::<f64>().ok().filter(|w| (0.0..=20.0).contains(w))?,
            "text" => s.text = value.chars().take(10_000).collect(),
            "textColor" => s.text_color = color(value)?.unwrap_or_else(|| "#000000".into()),
            "fontSize" => s.font_size = value.parse::<f64>().ok().filter(|p| (1.0..=409.0).contains(p))?,
            "bold" => s.bold = flag(value),
            "italic" => s.italic = flag(value),
            "align" => s.align = TextAlign::from_id(value)?,
            "arrowEnd" => s.arrow_end = flag(value),
            "preset" if is_known_preset(value) && is_line(value) == is_line(&self.preset) => {
                s.preset = value.to_string()
            }
            _ => return None,
        }
        Some(s)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_follow_excel() {
        let s = Shape::new("ellipse");
        assert_eq!(s.fill.as_deref(), Some("#4472C4"));
        assert_eq!(s.text_color, "#FFFFFF");
        let l = Shape::new("straightConnector1");
        assert!(l.fill.is_none() && l.arrow_end);
        let t = Shape::text_box();
        assert!(t.text_top && t.align == TextAlign::Left);
    }

    #[test]
    fn patches_validate_values() {
        let s = Shape::new("rect");
        assert_eq!(s.patched("fill", "#ff0000").unwrap().fill.as_deref(), Some("#FF0000"));
        assert!(s.patched("fill", "").unwrap().fill.is_none());
        assert!(s.patched("fill", "red").is_none());
        assert!(s.patched("fontSize", "0").is_none());
        assert_eq!(s.patched("preset", "star5").unwrap().preset, "star5");
        // A filled shape can't become a line (it would lose its fill and text).
        assert!(s.patched("preset", "line").is_none());
        assert!(s.patched("nope", "1").is_none());
    }
}
