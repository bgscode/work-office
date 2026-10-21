//! Cell borders: line styles, the four edges of a cell, and Excel's
//! Home → Borders presets applied to a range.
//!
//! Like Excel, every cell stores its own four edges. A line between two cells
//! can therefore be stored on either of them. When a preset sets or clears an
//! edge on the outside of the range, the facing edge of the neighbouring cell
//! is cleared too, so the line you just chose is the one you see (and "No
//! Border" really removes the line).

use crate::cell::{CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::sheet::Sheet;
use crate::style::Style;

/// The four edges of a cell. `None` means no line on that side.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct Borders {
    pub top: Option<BorderLine>,
    pub right: Option<BorderLine>,
    pub bottom: Option<BorderLine>,
    pub left: Option<BorderLine>,
}

impl Borders {
    pub const NONE: Borders = Borders { top: None, right: None, bottom: None, left: None };

    pub fn is_empty(&self) -> bool {
        *self == Borders::NONE
    }

    pub fn side(&self, side: Side) -> Option<&BorderLine> {
        match side {
            Side::Top => self.top.as_ref(),
            Side::Right => self.right.as_ref(),
            Side::Bottom => self.bottom.as_ref(),
            Side::Left => self.left.as_ref(),
        }
    }

    pub fn side_mut(&mut self, side: Side) -> &mut Option<BorderLine> {
        match side {
            Side::Top => &mut self.top,
            Side::Right => &mut self.right,
            Side::Bottom => &mut self.bottom,
            Side::Left => &mut self.left,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Side {
    Top,
    Right,
    Bottom,
    Left,
}

impl Side {
    pub const ALL: [Side; 4] = [Side::Top, Side::Right, Side::Bottom, Side::Left];

    /// The facing side of the neighbouring cell.
    pub fn opposite(self) -> Side {
        match self {
            Side::Top => Side::Bottom,
            Side::Right => Side::Left,
            Side::Bottom => Side::Top,
            Side::Left => Side::Right,
        }
    }

    /// The neighbouring cell across this side, if it is inside the sheet.
    fn neighbour(self, at: CellRef) -> Option<CellRef> {
        match self {
            Side::Top => at.row.checked_sub(1).map(|r| CellRef::new(r, at.col)),
            Side::Bottom => (at.row + 1 < MAX_ROWS).then(|| CellRef::new(at.row + 1, at.col)),
            Side::Left => at.col.checked_sub(1).map(|c| CellRef::new(at.row, c)),
            Side::Right => (at.col + 1 < MAX_COLS).then(|| CellRef::new(at.row, at.col + 1)),
        }
    }
}

/// One border line: its style and colour (`None` = automatic, black).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct BorderLine {
    pub style: BorderStyle,
    /// `#RRGGBB`.
    pub color: Option<String>,
}

impl BorderLine {
    pub fn new(style: BorderStyle, color: Option<String>) -> BorderLine {
        BorderLine { style, color }
    }
}

/// Excel's border line styles. `name()` is the `.xlsx` attribute value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BorderStyle {
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl BorderStyle {
    pub const ALL: [BorderStyle; 13] = [
        BorderStyle::Thin,
        BorderStyle::Medium,
        BorderStyle::Thick,
        BorderStyle::Dashed,
        BorderStyle::Dotted,
        BorderStyle::Double,
        BorderStyle::Hair,
        BorderStyle::MediumDashed,
        BorderStyle::DashDot,
        BorderStyle::MediumDashDot,
        BorderStyle::DashDotDot,
        BorderStyle::MediumDashDotDot,
        BorderStyle::SlantDashDot,
    ];

    pub fn name(self) -> &'static str {
        match self {
            BorderStyle::Thin => "thin",
            BorderStyle::Medium => "medium",
            BorderStyle::Thick => "thick",
            BorderStyle::Dashed => "dashed",
            BorderStyle::Dotted => "dotted",
            BorderStyle::Double => "double",
            BorderStyle::Hair => "hair",
            BorderStyle::MediumDashed => "mediumDashed",
            BorderStyle::DashDot => "dashDot",
            BorderStyle::MediumDashDot => "mediumDashDot",
            BorderStyle::DashDotDot => "dashDotDot",
            BorderStyle::MediumDashDotDot => "mediumDashDotDot",
            BorderStyle::SlantDashDot => "slantDashDot",
        }
    }

    /// `None` for "none" and unknown names.
    pub fn from_name(name: &str) -> Option<BorderStyle> {
        BorderStyle::ALL.into_iter().find(|s| s.name() == name)
    }
}

/// The entries of Excel's Home → Borders menu.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderPreset {
    Bottom,
    Top,
    Left,
    Right,
    /// Remove every border in (and around) the range.
    None,
    /// Every edge of every cell.
    All,
    /// The outline of the range.
    Outside,
    /// The outline of the range, thick.
    ThickOutside,
    /// The lines between cells, not the outline.
    Inside,
    InsideHorizontal,
    InsideVertical,
    BottomDouble,
    ThickBottom,
    TopBottom,
    TopThickBottom,
    TopDoubleBottom,
}

impl BorderPreset {
    pub fn from_name(name: &str) -> Option<BorderPreset> {
        Some(match name {
            "bottom" => BorderPreset::Bottom,
            "top" => BorderPreset::Top,
            "left" => BorderPreset::Left,
            "right" => BorderPreset::Right,
            "none" => BorderPreset::None,
            "all" => BorderPreset::All,
            "outside" => BorderPreset::Outside,
            "thickOutside" => BorderPreset::ThickOutside,
            "inside" => BorderPreset::Inside,
            "insideHorizontal" => BorderPreset::InsideHorizontal,
            "insideVertical" => BorderPreset::InsideVertical,
            "bottomDouble" => BorderPreset::BottomDouble,
            "thickBottom" => BorderPreset::ThickBottom,
            "topBottom" => BorderPreset::TopBottom,
            "topThickBottom" => BorderPreset::TopThickBottom,
            "topDoubleBottom" => BorderPreset::TopDoubleBottom,
            _ => return None,
        })
    }
}

/// What a preset does to one kind of edge.
#[derive(Clone)]
enum Edit {
    Keep,
    Set(BorderLine),
    Clear,
}

/// Per-edge plan for a preset: the four outline edges and the inner lines.
struct Plan {
    top: Edit,
    bottom: Edit,
    left: Edit,
    right: Edit,
    inner_h: Edit,
    inner_v: Edit,
}

impl Plan {
    fn new(preset: BorderPreset, line: &BorderLine) -> Plan {
        use BorderPreset as P;
        let set = || Edit::Set(line.clone());
        let with = |style| Edit::Set(BorderLine::new(style, line.color.clone()));
        let mut p = Plan {
            top: Edit::Keep,
            bottom: Edit::Keep,
            left: Edit::Keep,
            right: Edit::Keep,
            inner_h: Edit::Keep,
            inner_v: Edit::Keep,
        };
        match preset {
            P::Bottom => p.bottom = set(),
            P::Top => p.top = set(),
            P::Left => p.left = set(),
            P::Right => p.right = set(),
            P::None => {
                p = Plan {
                    top: Edit::Clear,
                    bottom: Edit::Clear,
                    left: Edit::Clear,
                    right: Edit::Clear,
                    inner_h: Edit::Clear,
                    inner_v: Edit::Clear,
                }
            }
            P::All => p = Plan { top: set(), bottom: set(), left: set(), right: set(), inner_h: set(), inner_v: set() },
            P::Outside => {
                (p.top, p.bottom, p.left, p.right) = (set(), set(), set(), set());
            }
            P::ThickOutside => {
                let thick = || with(BorderStyle::Thick);
                (p.top, p.bottom, p.left, p.right) = (thick(), thick(), thick(), thick());
            }
            P::Inside => (p.inner_h, p.inner_v) = (set(), set()),
            P::InsideHorizontal => p.inner_h = set(),
            P::InsideVertical => p.inner_v = set(),
            P::BottomDouble => p.bottom = with(BorderStyle::Double),
            P::ThickBottom => p.bottom = with(BorderStyle::Thick),
            P::TopBottom => (p.top, p.bottom) = (set(), set()),
            P::TopThickBottom => (p.top, p.bottom) = (set(), with(BorderStyle::Thick)),
            P::TopDoubleBottom => (p.top, p.bottom) = (set(), with(BorderStyle::Double)),
        }
        p
    }

    /// What happens to `side` of the cell at `at` inside `range`.
    fn edit(&self, range: Range, at: CellRef, side: Side) -> &Edit {
        match side {
            Side::Top if at.row == range.start.row => &self.top,
            Side::Bottom if at.row == range.end.row => &self.bottom,
            Side::Left if at.col == range.start.col => &self.left,
            Side::Right if at.col == range.end.col => &self.right,
            Side::Top | Side::Bottom => &self.inner_h,
            Side::Left | Side::Right => &self.inner_v,
        }
    }
}

impl Sheet {
    /// New styles for applying a Borders preset to `range`. `line` is the
    /// style and colour chosen in the menu; presets with their own style
    /// (thick outside, double bottom, …) keep that style and use the colour.
    /// Cells just outside the range are included when their facing edge is
    /// cleared. Apply the result with `ops::SetStyles`.
    pub fn border_edits(&self, range: Range, preset: BorderPreset, line: &BorderLine) -> Vec<(CellRef, Style)> {
        let plan = Plan::new(preset, line);
        let mut out: Vec<(CellRef, Style)> = Vec::new();
        for at in range.cells() {
            let mut style = self.style(at).clone();
            for side in Side::ALL {
                match plan.edit(range, at, side) {
                    Edit::Keep => {}
                    Edit::Set(l) => *style.borders.side_mut(side) = Some(l.clone()),
                    Edit::Clear => *style.borders.side_mut(side) = None,
                }
            }
            if style != *self.style(at) {
                out.push((at, style));
            }
        }
        // Neighbours across the outline: clear their facing edge when we touched ours.
        let outline = [
            (Side::Top, &plan.top),
            (Side::Bottom, &plan.bottom),
            (Side::Left, &plan.left),
            (Side::Right, &plan.right),
        ];
        for (side, edit) in outline {
            if matches!(edit, Edit::Keep) {
                continue;
            }
            for at in edge_cells(range, side) {
                let Some(n) = side.neighbour(at) else { continue };
                let facing = side.opposite();
                if self.style(n).borders.side(facing).is_some() {
                    let mut style = self.style(n).clone();
                    *style.borders.side_mut(facing) = None;
                    out.push((n, style));
                }
            }
        }
        out
    }
}

/// The cells of `range` along one of its outline edges.
fn edge_cells(range: Range, side: Side) -> Vec<CellRef> {
    let (s, e) = (range.start, range.end);
    match side {
        Side::Top => (s.col..=e.col).map(|c| CellRef::new(s.row, c)).collect(),
        Side::Bottom => (s.col..=e.col).map(|c| CellRef::new(e.row, c)).collect(),
        Side::Left => (s.row..=e.row).map(|r| CellRef::new(r, s.col)).collect(),
        Side::Right => (s.row..=e.row).map(|r| CellRef::new(r, e.col)).collect(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn range(s: &str) -> Range {
        let (a, b) = s.split_once(':').unwrap_or((s, s));
        Range::new(at(a), at(b))
    }

    fn thin() -> BorderLine {
        BorderLine::new(BorderStyle::Thin, None)
    }

    fn apply(sheet: &mut Sheet, r: &str, preset: BorderPreset, line: &BorderLine) {
        for (cell, style) in sheet.border_edits(range(r), preset, line) {
            sheet.set_style(cell, style);
        }
    }

    fn sides(sheet: &Sheet, cell: &str) -> String {
        let b = &sheet.style(at(cell)).borders;
        Side::ALL
            .into_iter()
            .map(|s| match b.side(s) {
                Some(l) if l.style == BorderStyle::Thin => '1',
                Some(l) if l.style == BorderStyle::Thick => '3',
                Some(l) if l.style == BorderStyle::Double => '=',
                Some(_) => '?',
                None => '.',
            })
            .collect()
    }

    #[test]
    fn style_names_round_trip() {
        for s in BorderStyle::ALL {
            assert_eq!(BorderStyle::from_name(s.name()), Some(s));
        }
        assert_eq!(BorderStyle::from_name("none"), None);
        assert_eq!(Side::Top.opposite(), Side::Bottom);
        assert_eq!(BorderPreset::from_name("thickOutside"), Some(BorderPreset::ThickOutside));
        assert_eq!(BorderPreset::from_name("diagonal"), None);
    }

    #[test]
    fn all_borders_sets_every_edge() {
        let mut s = Sheet::new();
        apply(&mut s, "A1:B2", BorderPreset::All, &thin());
        // Sides are listed top, right, bottom, left.
        for cell in ["A1", "B1", "A2", "B2"] {
            assert_eq!(sides(&s, cell), "1111", "{cell}");
        }
        assert!(s.style(at("C1")).borders.is_empty());
    }

    #[test]
    fn outside_borders_only_touch_the_outline() {
        let mut s = Sheet::new();
        apply(&mut s, "A1:C3", BorderPreset::Outside, &thin());
        assert_eq!(sides(&s, "A1"), "1..1");
        assert_eq!(sides(&s, "B1"), "1...");
        assert_eq!(sides(&s, "C3"), ".11.");
        assert_eq!(sides(&s, "B2"), "....");
        apply(&mut s, "A1:C3", BorderPreset::ThickOutside, &thin());
        assert_eq!(sides(&s, "A1"), "3..3");
        assert_eq!(sides(&s, "C3"), ".33.");
    }

    #[test]
    fn inside_borders_skip_the_outline() {
        let mut s = Sheet::new();
        apply(&mut s, "A1:B2", BorderPreset::Inside, &thin());
        assert_eq!(sides(&s, "A1"), ".11.");
        assert_eq!(sides(&s, "B2"), "1..1");
        let mut s = Sheet::new();
        apply(&mut s, "A1:B2", BorderPreset::InsideHorizontal, &thin());
        assert_eq!(sides(&s, "A1"), "..1.");
        assert_eq!(sides(&s, "A2"), "1...");
    }

    #[test]
    fn bottom_presets_and_colour() {
        let mut s = Sheet::new();
        let red = BorderLine::new(BorderStyle::Thin, Some("#FF0000".into()));
        apply(&mut s, "A1:B1", BorderPreset::BottomDouble, &red);
        let line = s.style(at("B1")).borders.bottom.clone().unwrap();
        assert_eq!(line, BorderLine::new(BorderStyle::Double, Some("#FF0000".into())));
        apply(&mut s, "A1", BorderPreset::TopThickBottom, &thin());
        assert_eq!(sides(&s, "A1"), "1.3.");
        apply(&mut s, "A5", BorderPreset::Left, &BorderLine::new(BorderStyle::Dashed, None));
        assert_eq!(s.style(at("A5")).borders.left.as_ref().unwrap().style, BorderStyle::Dashed);
    }

    #[test]
    fn setting_an_edge_replaces_the_neighbours_facing_line() {
        let mut s = Sheet::new();
        apply(&mut s, "B2", BorderPreset::Bottom, &thin());
        // Thick top on B3 is the same line: B2's bottom goes, B3's top shows.
        apply(&mut s, "B3", BorderPreset::Top, &BorderLine::new(BorderStyle::Thick, None));
        assert_eq!(sides(&s, "B2"), "....");
        assert_eq!(sides(&s, "B3"), "3...");
    }

    #[test]
    fn no_border_also_clears_lines_drawn_by_neighbours() {
        let mut s = Sheet::new();
        apply(&mut s, "A1:C3", BorderPreset::All, &thin());
        apply(&mut s, "B2", BorderPreset::None, &thin());
        assert_eq!(sides(&s, "B2"), "....");
        assert_eq!(sides(&s, "B1"), "11.1"); // bottom gone
        assert_eq!(sides(&s, "A2"), "1.11"); // right gone
        assert_eq!(sides(&s, "C2"), "111."); // left gone
        assert_eq!(sides(&s, "B3"), ".111"); // top gone
                                             // Keeps other formatting.
        s.set_style(at("E5"), Style { bold: true, ..Style::default() });
        apply(&mut s, "E5", BorderPreset::All, &thin());
        apply(&mut s, "E5", BorderPreset::None, &thin());
        assert!(s.style(at("E5")).bold);
        assert!(s.style(at("E5")).borders.is_empty());
    }

    #[test]
    fn edges_of_the_sheet_have_no_neighbours() {
        let s = Sheet::new();
        let edits = s.border_edits(range("A1"), BorderPreset::Outside, &thin());
        assert_eq!(edits.len(), 1);
        let last = CellRef::new(MAX_ROWS - 1, MAX_COLS - 1);
        let edits = s.border_edits(Range::new(last, last), BorderPreset::None, &thin());
        assert!(edits.is_empty());
    }
}
