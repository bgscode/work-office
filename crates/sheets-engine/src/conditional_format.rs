//! Conditional formatting (Home → Conditional Formatting).
//!
//! A [`CfRule`] applies to one or more ranges and either *highlights* cells
//! that match a condition with a [`CfFormat`] (Excel's "dxf"), or paints a
//! colour scale or data bar from each cell's number. Rules are kept in
//! priority order (index 0 first). For a cell, the first matching rule sets a
//! property (fill, font colour, bold...) and later rules can only fill in
//! properties that are still unset, as in Excel; "Stop If True" ends the
//! search for that cell.
//!
//! [`evaluate`] works on a block of cells (the visible part of the grid), so
//! rules that need statistics over their whole range (Top 10, Above Average,
//! Duplicates, colour scales, data bars) compute them once per call.

use std::collections::HashMap;

use crate::cell::{CellRef, Range};
use crate::data_tools::{self, json_str};
use crate::date;
use crate::format;
use crate::formula::{self, refs};
use crate::sheet::Sheet;
use crate::value::{compare, Value};

/// The differential format a highlight rule applies. `None` leaves the cell's own formatting.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct CfFormat {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike: Option<bool>,
    /// Text colour `#RRGGBB`.
    pub color: Option<String>,
    /// Fill colour `#RRGGBB`.
    pub fill: Option<String>,
}

impl CfFormat {
    /// Excel's presets in the highlight dialogs, by name.
    pub fn preset(name: &str) -> Option<CfFormat> {
        let f = |fill: Option<&str>, color: Option<&str>| CfFormat {
            fill: fill.map(String::from),
            color: color.map(String::from),
            ..CfFormat::default()
        };
        Some(match name {
            "lightRedDarkRed" => f(Some("#FFC7CE"), Some("#9C0006")),
            "yellowDarkYellow" => f(Some("#FFEB9C"), Some("#9C5700")),
            "greenDarkGreen" => f(Some("#C6EFCE"), Some("#006100")),
            "lightRed" => f(Some("#FFC7CE"), None),
            "redText" => f(None, Some("#9C0006")),
            _ => return None,
        })
    }

    /// Parse `fill=#FFC7CE;color=#9C0006;b=1;i=0` (the UI's custom format), or a preset name.
    pub fn parse(spec: &str) -> CfFormat {
        if let Some(p) = CfFormat::preset(spec) {
            return p;
        }
        let mut f = CfFormat::default();
        for part in spec.split(';') {
            let Some((k, v)) = part.split_once('=') else { continue };
            let on = Some(v == "1");
            let color = (!v.is_empty()).then(|| v.to_string());
            match k {
                "b" => f.bold = on,
                "i" => f.italic = on,
                "u" => f.underline = on,
                "s" => f.strike = on,
                "color" => f.color = color,
                "fill" => f.fill = color,
                _ => {}
            }
        }
        f
    }

    fn merge_into(&self, out: &mut CfFormat) {
        let take = |slot: &mut Option<bool>, v: Option<bool>| {
            if slot.is_none() {
                *slot = v;
            }
        };
        take(&mut out.bold, self.bold);
        take(&mut out.italic, self.italic);
        take(&mut out.underline, self.underline);
        take(&mut out.strike, self.strike);
        if out.color.is_none() {
            out.color.clone_from(&self.color);
        }
        if out.fill.is_none() {
            out.fill.clone_from(&self.fill);
        }
    }

    /// JSON with the same keys as a cell style (`b`, `i`, `u`, `s`, `color`, `fill`).
    pub fn to_json(&self) -> String {
        let mut parts = Vec::new();
        for (k, v) in [("b", self.bold), ("i", self.italic), ("u", self.underline), ("s", self.strike)] {
            if let Some(on) = v {
                parts.push(format!("\"{k}\":{}", on as u8));
            }
        }
        if let Some(c) = &self.color {
            parts.push(format!("\"color\":{}", json_str(c)));
        }
        if let Some(c) = &self.fill {
            parts.push(format!("\"fill\":{}", json_str(c)));
        }
        format!("{{{}}}", parts.join(","))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellOp {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    Greater,
    Less,
    GreaterEqual,
    LessEqual,
}

impl CellOp {
    pub fn name(self) -> &'static str {
        match self {
            CellOp::Between => "between",
            CellOp::NotBetween => "notBetween",
            CellOp::Equal => "equal",
            CellOp::NotEqual => "notEqual",
            CellOp::Greater => "greaterThan",
            CellOp::Less => "lessThan",
            CellOp::GreaterEqual => "greaterThanOrEqual",
            CellOp::LessEqual => "lessThanOrEqual",
        }
    }

    pub fn parse(s: &str) -> Option<CellOp> {
        Some(match s {
            "between" => CellOp::Between,
            "notBetween" => CellOp::NotBetween,
            "equal" => CellOp::Equal,
            "notEqual" => CellOp::NotEqual,
            "greaterThan" => CellOp::Greater,
            "lessThan" => CellOp::Less,
            "greaterThanOrEqual" => CellOp::GreaterEqual,
            "lessThanOrEqual" => CellOp::LessEqual,
            _ => return None,
        })
    }

    fn symbol(self) -> &'static str {
        match self {
            CellOp::Between => "between",
            CellOp::NotBetween => "not between",
            CellOp::Equal => "=",
            CellOp::NotEqual => "<>",
            CellOp::Greater => ">",
            CellOp::Less => "<",
            CellOp::GreaterEqual => ">=",
            CellOp::LessEqual => "<=",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextOp {
    Contains,
    NotContains,
    BeginsWith,
    EndsWith,
}

impl TextOp {
    /// The `cfRule type=` for this operator in `.xlsx`.
    pub fn rule_type(self) -> &'static str {
        match self {
            TextOp::Contains => "containsText",
            TextOp::NotContains => "notContainsText",
            TextOp::BeginsWith => "beginsWith",
            TextOp::EndsWith => "endsWith",
        }
    }

    pub fn parse(s: &str) -> Option<TextOp> {
        Some(match s {
            "containsText" => TextOp::Contains,
            "notContainsText" | "notContains" => TextOp::NotContains,
            "beginsWith" => TextOp::BeginsWith,
            "endsWith" => TextOp::EndsWith,
            _ => return None,
        })
    }
}

/// "A Date Occurring" periods. Weeks start on Sunday, as in Excel.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimePeriod {
    Yesterday,
    Today,
    Tomorrow,
    Last7Days,
    LastWeek,
    ThisWeek,
    NextWeek,
    LastMonth,
    ThisMonth,
    NextMonth,
}

impl TimePeriod {
    pub const ALL: [TimePeriod; 10] = [
        TimePeriod::Yesterday,
        TimePeriod::Today,
        TimePeriod::Tomorrow,
        TimePeriod::Last7Days,
        TimePeriod::LastWeek,
        TimePeriod::ThisWeek,
        TimePeriod::NextWeek,
        TimePeriod::LastMonth,
        TimePeriod::ThisMonth,
        TimePeriod::NextMonth,
    ];

    pub fn name(self) -> &'static str {
        match self {
            TimePeriod::Yesterday => "yesterday",
            TimePeriod::Today => "today",
            TimePeriod::Tomorrow => "tomorrow",
            TimePeriod::Last7Days => "last7Days",
            TimePeriod::LastWeek => "lastWeek",
            TimePeriod::ThisWeek => "thisWeek",
            TimePeriod::NextWeek => "nextWeek",
            TimePeriod::LastMonth => "lastMonth",
            TimePeriod::ThisMonth => "thisMonth",
            TimePeriod::NextMonth => "nextMonth",
        }
    }

    pub fn parse(s: &str) -> Option<TimePeriod> {
        TimePeriod::ALL.into_iter().find(|p| p.name() == s)
    }

    /// Does the date serial `day` fall in this period, given today's serial?
    pub fn contains(self, day: f64, today: f64) -> bool {
        let (d, t) = (day.floor(), today.floor());
        // Serial 1 (1900-01-01) was a Sunday; weekday 0 = Sunday.
        let week_start = |s: f64| s - (s - 1.0).rem_euclid(7.0);
        let month = |s: f64| {
            let (y, m, _) = date::ymd_from_serial(s);
            y * 12 + m as i64
        };
        match self {
            TimePeriod::Yesterday => d == t - 1.0,
            TimePeriod::Today => d == t,
            TimePeriod::Tomorrow => d == t + 1.0,
            TimePeriod::Last7Days => d > t - 7.0 && d <= t,
            TimePeriod::LastWeek => week_start(d) == week_start(t) - 7.0,
            TimePeriod::ThisWeek => week_start(d) == week_start(t),
            TimePeriod::NextWeek => week_start(d) == week_start(t) + 7.0,
            TimePeriod::LastMonth => month(d) == month(t) - 1,
            TimePeriod::ThisMonth => month(d) == month(t),
            TimePeriod::NextMonth => month(d) == month(t) + 1,
        }
    }
}

/// A colour-scale / data-bar end point ("conditional format value object").
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Cfvo {
    Min,
    Max,
    Number(f64),
    Percent(f64),
    Percentile(f64),
}

impl Cfvo {
    pub fn kind(&self) -> &'static str {
        match self {
            Cfvo::Min => "min",
            Cfvo::Max => "max",
            Cfvo::Number(_) => "num",
            Cfvo::Percent(_) => "percent",
            Cfvo::Percentile(_) => "percentile",
        }
    }

    pub fn value(&self) -> Option<f64> {
        match self {
            Cfvo::Number(v) | Cfvo::Percent(v) | Cfvo::Percentile(v) => Some(*v),
            _ => None,
        }
    }

    pub fn parse(kind: &str, val: &str) -> Cfvo {
        let v = val.trim().parse::<f64>().unwrap_or(0.0);
        match kind {
            "max" | "autoMax" => Cfvo::Max,
            "num" => Cfvo::Number(v),
            "percent" => Cfvo::Percent(v),
            "percentile" => Cfvo::Percentile(v),
            _ => Cfvo::Min,
        }
    }

    /// The actual number this point stands for, given the sorted numbers of the range.
    fn resolve(&self, sorted: &[f64]) -> f64 {
        let (lo, hi) = (sorted[0], sorted[sorted.len() - 1]);
        match *self {
            Cfvo::Min => lo,
            Cfvo::Max => hi,
            Cfvo::Number(v) => v,
            Cfvo::Percent(p) => lo + (hi - lo) * p / 100.0,
            Cfvo::Percentile(p) => percentile(sorted, p / 100.0),
        }
    }
}

/// Excel's PERCENTILE.INC on sorted data.
fn percentile(sorted: &[f64], k: f64) -> f64 {
    let pos = k.clamp(0.0, 1.0) * (sorted.len() - 1) as f64;
    let i = pos.floor() as usize;
    let frac = pos - i as f64;
    match sorted.get(i + 1) {
        Some(next) => sorted[i] + frac * (next - sorted[i]),
        None => sorted[i],
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum CfKind {
    /// "Cell Value" comparisons; `f1`/`f2` are formula text without `=`.
    CellIs {
        op: CellOp,
        f1: String,
        f2: String,
    },
    Text {
        op: TextOp,
        text: String,
    },
    TimePeriod(TimePeriod),
    /// Duplicate (or, with `unique`, unique) values in the rule's range.
    Duplicates {
        unique: bool,
    },
    Top10 {
        bottom: bool,
        percent: bool,
        rank: u32,
    },
    Average {
        below: bool,
        equal: bool,
    },
    /// "Use a formula": relative to the top-left cell of the first range.
    Expression {
        formula: String,
    },
    /// 2 or 3 points with their `#RRGGBB` colours, lowest first.
    ColorScale {
        points: Vec<(Cfvo, String)>,
    },
    DataBar {
        min: Cfvo,
        max: Cfvo,
        color: String,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub struct CfRule {
    pub ranges: Vec<Range>,
    pub kind: CfKind,
    /// Used by highlight rules (ignored by colour scales and data bars).
    pub format: CfFormat,
    pub stop_if_true: bool,
}

impl CfRule {
    pub fn new(ranges: Vec<Range>, kind: CfKind, format: CfFormat) -> CfRule {
        CfRule { ranges, kind, format, stop_if_true: false }
    }

    fn origin(&self) -> CellRef {
        self.ranges.first().map_or(CellRef::new(0, 0), |r| r.start)
    }

    /// Evaluate formula text (no `=`) as seen from `at`.
    fn eval_at(&self, sheet: &Sheet, f: &str, at: CellRef) -> Value {
        let o = self.origin();
        let shifted = refs::shift_input(&format!("={f}"), at.row as i64 - o.row as i64, at.col as i64 - o.col as i64);
        match formula::parse(&shifted[1..]) {
            Ok(e) => formula::eval(&e, sheet),
            Err(e) => Value::Error(e),
        }
    }

    /// What the Manage Rules dialog shows in its "Rule" column.
    pub fn describe(&self) -> String {
        match &self.kind {
            CfKind::CellIs { op, f1, f2 } => match op {
                CellOp::Between | CellOp::NotBetween => {
                    format!("Cell Value {} ={f1} and ={f2}", op.symbol())
                }
                _ => format!("Cell Value {} {}", op.symbol(), shown_formula(f1)),
            },
            CfKind::Text { op, text } => match op {
                TextOp::Contains => format!("Cell Value contains '{text}'"),
                TextOp::NotContains => format!("Cell Value does not contain '{text}'"),
                TextOp::BeginsWith => format!("Cell Value begins with '{text}'"),
                TextOp::EndsWith => format!("Cell Value ends with '{text}'"),
            },
            CfKind::TimePeriod(p) => format!("Date Occurring: {}", period_label(*p)),
            CfKind::Duplicates { unique: false } => "Duplicate Values".into(),
            CfKind::Duplicates { unique: true } => "Unique Values".into(),
            CfKind::Top10 { bottom, percent, rank } => {
                format!("{} {rank}{}", if *bottom { "Bottom" } else { "Top" }, if *percent { "%" } else { "" })
            }
            CfKind::Average { below, equal } => {
                let side = if *below { "Below" } else { "Above" };
                if *equal {
                    format!("Equal To or {side} Average")
                } else {
                    format!("{side} Average")
                }
            }
            CfKind::Expression { formula } => format!("Formula: ={formula}"),
            CfKind::ColorScale { .. } => "Graded Color Scale".into(),
            CfKind::DataBar { .. } => "Data Bar".into(),
        }
    }

    /// JSON for the Manage Rules dialog.
    pub fn to_json(&self) -> String {
        let extra = match &self.kind {
            CfKind::ColorScale { points } => {
                format!(",\"scale\":[{}]", points.iter().map(|(_, c)| json_str(c)).collect::<Vec<_>>().join(","))
            }
            CfKind::DataBar { color, .. } => format!(",\"bar\":{}", json_str(color)),
            _ => String::new(),
        };
        format!(
            "{{\"description\":{},\"ranges\":{},\"format\":{},\"stopIfTrue\":{}{extra}}}",
            json_str(&self.describe()),
            json_str(&data_tools::format_sqref(&self.ranges)),
            self.format.to_json(),
            self.stop_if_true
        )
    }
}

fn shown_formula(f: &str) -> String {
    if f.parse::<f64>().is_ok() || f.starts_with('"') {
        f.to_string()
    } else {
        format!("={f}")
    }
}

fn period_label(p: TimePeriod) -> &'static str {
    match p {
        TimePeriod::Yesterday => "Yesterday",
        TimePeriod::Today => "Today",
        TimePeriod::Tomorrow => "Tomorrow",
        TimePeriod::Last7Days => "In the last 7 days",
        TimePeriod::LastWeek => "Last week",
        TimePeriod::ThisWeek => "This week",
        TimePeriod::NextWeek => "Next week",
        TimePeriod::LastMonth => "Last month",
        TimePeriod::ThisMonth => "This month",
        TimePeriod::NextMonth => "Next month",
    }
}

/// What conditional formatting does to one cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CfResult {
    pub format: CfFormat,
    /// Data bar: length as a fraction of the cell width (0..=1) and colour.
    pub bar: Option<(f64, String)>,
}

impl CfResult {
    pub fn to_json(&self) -> String {
        let mut s = self.format.to_json();
        if let Some((len, color)) = &self.bar {
            s.pop();
            if s.len() > 1 {
                s.push(',');
            }
            s.push_str(&format!("\"bar\":{:.4},\"barColor\":{}}}", len, json_str(color)));
        }
        s
    }
}

/// Statistics a rule needs over its whole range.
#[derive(Default)]
struct Stats {
    /// Numbers in the range, ascending.
    numbers: Vec<f64>,
    /// Their average (Above/Below Average rules).
    mean: f64,
    /// Lower-cased displayed text -> how many cells show it.
    counts: HashMap<String, u32>,
}

fn stats(sheet: &Sheet, rule: &CfRule) -> Stats {
    let mut s = Stats::default();
    // Rules that only look at their own cell need nothing from the range.
    if matches!(
        rule.kind,
        CfKind::CellIs { .. } | CfKind::Text { .. } | CfKind::TimePeriod(_) | CfKind::Expression { .. }
    ) {
        return s;
    }
    let need_counts = matches!(rule.kind, CfKind::Duplicates { .. });
    for r in &rule.ranges {
        // Only cells that hold something matter; avoid walking huge empty areas.
        let (rows, cols) = sheet.extent();
        if r.start.row >= rows || r.start.col >= cols {
            continue;
        }
        let clipped = Range::new(r.start, CellRef::new(r.end.row.min(rows - 1), r.end.col.min(cols - 1)));
        for at in clipped.cells() {
            match sheet.value(at) {
                Value::Empty => continue,
                Value::Number(n) => s.numbers.push(*n),
                _ => {}
            }
            if need_counts {
                *s.counts.entry(sheet.display(at).to_lowercase()).or_default() += 1;
            }
        }
    }
    s.numbers.sort_by(f64::total_cmp);
    if !s.numbers.is_empty() {
        s.mean = s.numbers.iter().sum::<f64>() / s.numbers.len() as f64;
    }
    s
}

/// Per-rule statistics over each rule's whole range (ranks, averages,
/// duplicates, scale bounds), computed once and reused for every block the
/// grid paints until the sheet changes. See [`evaluate_prepared`].
pub struct CfPrepared {
    stats: Vec<Stats>,
}

pub fn prepare(sheet: &Sheet) -> CfPrepared {
    CfPrepared { stats: sheet.data_tools().cond_formats.iter().map(|r| stats(sheet, r)).collect() }
}

fn lerp_color(a: &str, b: &str, t: f64) -> String {
    let rgb = |s: &str| {
        let h = s.trim_start_matches('#');
        let p = |i: usize| u8::from_str_radix(h.get(i..i + 2).unwrap_or("00"), 16).unwrap_or(0) as f64;
        (p(0), p(2), p(4))
    };
    let (a, b) = (rgb(a), rgb(b));
    let mix = |x: f64, y: f64| (x + (y - x) * t.clamp(0.0, 1.0)).round() as u8;
    format!("#{:02X}{:02X}{:02X}", mix(a.0, b.0), mix(a.1, b.1), mix(a.2, b.2))
}

fn matches(sheet: &Sheet, rule: &CfRule, st: &Stats, at: CellRef, today: f64) -> bool {
    let value = sheet.value(at);
    match &rule.kind {
        CfKind::CellIs { op, f1, f2 } => {
            let a = rule.eval_at(sheet, f1, at);
            let ord = |x: &Value| compare(value, x);
            use std::cmp::Ordering::*;
            if matches!(value, Value::Error(_)) {
                return false;
            }
            match op {
                CellOp::Equal => ord(&a) == Equal,
                CellOp::NotEqual => ord(&a) != Equal,
                CellOp::Greater => ord(&a) == Greater,
                CellOp::Less => ord(&a) == Less,
                CellOp::GreaterEqual => ord(&a) != Less,
                CellOp::LessEqual => ord(&a) != Greater,
                CellOp::Between | CellOp::NotBetween => {
                    let b = rule.eval_at(sheet, f2, at);
                    let (lo, hi) = if compare(&a, &b) == Greater { (b, a) } else { (a, b) };
                    let inside = ord(&lo) != Less && ord(&hi) != Greater;
                    inside == (*op == CellOp::Between)
                }
            }
        }
        CfKind::Text { op, text } => {
            if value.is_empty() && *op != TextOp::NotContains {
                return false;
            }
            let shown = sheet.display(at).to_lowercase();
            let t = text.to_lowercase();
            match op {
                TextOp::Contains => shown.contains(&t),
                TextOp::NotContains => !shown.contains(&t),
                TextOp::BeginsWith => shown.starts_with(&t),
                TextOp::EndsWith => shown.ends_with(&t),
            }
        }
        CfKind::TimePeriod(p) => matches!(value, Value::Number(n) if p.contains(*n, today)),
        CfKind::Duplicates { unique } => {
            if value.is_empty() {
                return false;
            }
            let n = st.counts.get(&sheet.display(at).to_lowercase()).copied().unwrap_or(0);
            (n > 1) != *unique
        }
        CfKind::Top10 { bottom, percent, rank } => {
            let Value::Number(x) = value else { return false };
            let len = st.numbers.len();
            if len == 0 {
                return false;
            }
            let n = if *percent { (len as f64 * *rank as f64 / 100.0).floor() as usize } else { *rank as usize };
            let n = n.clamp(1, len);
            if *bottom {
                *x <= st.numbers[n - 1]
            } else {
                *x >= st.numbers[len - n]
            }
        }
        CfKind::Average { below, equal } => {
            let Value::Number(x) = value else { return false };
            if st.numbers.is_empty() {
                return false;
            }
            let mean = st.mean;
            match (below, equal) {
                (false, false) => *x > mean,
                (false, true) => *x >= mean,
                (true, false) => *x < mean,
                (true, true) => *x <= mean,
            }
        }
        CfKind::Expression { formula } => rule.eval_at(sheet, formula, at).as_bool().unwrap_or(false),
        CfKind::ColorScale { .. } | CfKind::DataBar { .. } => matches!(value, Value::Number(_)),
    }
}

/// Conditional formatting for every cell of `block`, row-major; `None` where no rule applies.
pub fn evaluate(sheet: &Sheet, block: Range) -> Vec<Option<CfResult>> {
    evaluate_prepared(sheet, &prepare(sheet), block)
}

/// [`evaluate`] with statistics from [`prepare`] on the same, unchanged sheet.
pub fn evaluate_prepared(sheet: &Sheet, prepared: &CfPrepared, block: Range) -> Vec<Option<CfResult>> {
    let width = block.cols() as usize;
    let mut out: Vec<Option<CfResult>> = vec![None; block.rows() as usize * width];
    let mut stopped = vec![false; out.len()];
    let rules = &sheet.data_tools().cond_formats;
    if rules.is_empty() {
        return out;
    }
    let today = date::now_serial();
    for (rule, st) in rules.iter().zip(&prepared.stats) {
        if !rule.ranges.iter().any(|r| data_tools::intersects(r, &block)) {
            continue;
        }
        for r in rule.ranges.iter().filter(|r| data_tools::intersects(r, &block)) {
            let (top, bottom) = (r.start.row.max(block.start.row), r.end.row.min(block.end.row));
            let (left, right) = (r.start.col.max(block.start.col), r.end.col.min(block.end.col));
            for row in top..=bottom {
                for col in left..=right {
                    let i = (row - block.start.row) as usize * width + (col - block.start.col) as usize;
                    let at = CellRef::new(row, col);
                    if stopped[i] || !matches(sheet, rule, st, at, today) {
                        continue;
                    }
                    let slot = out[i].get_or_insert_with(CfResult::default);
                    match &rule.kind {
                        CfKind::ColorScale { points } if !st.numbers.is_empty() && points.len() >= 2 => {
                            let Value::Number(x) = sheet.value(at) else { continue };
                            if slot.format.fill.is_none() {
                                slot.format.fill = Some(scale_color(points, &st.numbers, *x));
                            }
                        }
                        CfKind::DataBar { min, max, color } if !st.numbers.is_empty() => {
                            let Value::Number(x) = sheet.value(at) else { continue };
                            if slot.bar.is_none() {
                                let (lo, hi) = (min.resolve(&st.numbers), max.resolve(&st.numbers));
                                // Excel draws the smallest value as a short bar, never as nothing.
                                let t = if hi > lo { ((x - lo) / (hi - lo)).clamp(0.0, 1.0) } else { 1.0 };
                                slot.bar = Some((0.1 + 0.9 * t, color.clone()));
                            }
                        }
                        CfKind::ColorScale { .. } | CfKind::DataBar { .. } => {}
                        _ => rule.format.merge_into(&mut slot.format),
                    }
                    if rule.stop_if_true {
                        stopped[i] = true;
                    }
                }
            }
        }
    }
    out
}

fn scale_color(points: &[(Cfvo, String)], sorted: &[f64], x: f64) -> String {
    let stops: Vec<(f64, &str)> = points.iter().map(|(p, c)| (p.resolve(sorted), c.as_str())).collect();
    if x <= stops[0].0 {
        return stops[0].1.to_string();
    }
    for w in stops.windows(2) {
        let ((a, ca), (b, cb)) = (w[0], w[1]);
        if x <= b {
            let t = if b > a { (x - a) / (b - a) } else { 1.0 };
            return lerp_color(ca, cb, t);
        }
    }
    stops[stops.len() - 1].1.to_string()
}

/// Build a rule from the UI's description. `kind` and `args`:
/// - `cellIs` [op, f1, f2]: op as in `CellOp::name`; f1/f2 typed values or `=formula`
/// - `text` [containsText|notContainsText|beginsWith|endsWith, text]
/// - `timePeriod` [period]
/// - `duplicates` / `unique`
/// - `top10` [top|bottom, items|percent, rank]
/// - `average` [above|below|aboveOrEqual|belowOrEqual]
/// - `expression` [formula]
/// - `colorScale` [#min, #mid, #max] or [#min, #max]
/// - `dataBar` [#color]
pub fn rule_from_spec(ranges: Vec<Range>, kind: &str, args: &[String], format: &str) -> Option<CfRule> {
    let arg = |i: usize| args.get(i).map(String::as_str).unwrap_or("");
    let kind = match kind {
        "cellIs" => CfKind::CellIs { op: CellOp::parse(arg(0))?, f1: criterion(arg(1)), f2: criterion(arg(2)) },
        "text" => CfKind::Text { op: TextOp::parse(arg(0))?, text: arg(1).to_string() },
        "timePeriod" => CfKind::TimePeriod(TimePeriod::parse(arg(0))?),
        "duplicates" => CfKind::Duplicates { unique: false },
        "unique" => CfKind::Duplicates { unique: true },
        "top10" => CfKind::Top10 {
            bottom: arg(0) == "bottom",
            percent: arg(1) == "percent",
            rank: arg(2).trim().parse::<u32>().ok().filter(|n| (1..=1000).contains(n))?,
        },
        "average" => CfKind::Average { below: arg(0).starts_with("below"), equal: arg(0).ends_with("OrEqual") },
        "expression" => CfKind::Expression { formula: arg(0).trim_start_matches('=').to_string() },
        "colorScale" => {
            let colors: Vec<String> = args.iter().filter(|c| !c.is_empty()).cloned().collect();
            let points = match colors.len() {
                2 => vec![(Cfvo::Min, colors[0].clone()), (Cfvo::Max, colors[1].clone())],
                3 => vec![
                    (Cfvo::Min, colors[0].clone()),
                    (Cfvo::Percentile(50.0), colors[1].clone()),
                    (Cfvo::Max, colors[2].clone()),
                ],
                _ => return None,
            };
            CfKind::ColorScale { points }
        }
        "dataBar" => CfKind::DataBar { min: Cfvo::Min, max: Cfvo::Max, color: arg(0).to_string() },
        _ => return None,
    };
    Some(CfRule::new(ranges, kind, CfFormat::parse(format)))
}

/// A value typed in a highlight dialog as rule formula text: `5` stays `5`,
/// `=A1` becomes `A1`, a date becomes its serial, other text is quoted.
pub fn criterion(typed: &str) -> String {
    let t = typed.trim();
    if t.is_empty() {
        return String::new();
    }
    if let Some(f) = t.strip_prefix('=') {
        return f.to_string();
    }
    if t.parse::<f64>().is_ok() {
        return t.to_string();
    }
    if let Some((n, _)) = format::parse_formatted(t) {
        return n.to_string();
    }
    format!("\"{}\"", t.replace('"', "\"\""))
}

impl Sheet {
    /// Data tools with `rule` added at the top of the priority list (as Excel does).
    pub fn with_cf_rule(&self, rule: CfRule) -> crate::data_tools::DataTools {
        let mut tools = self.data_tools().clone();
        tools.cond_formats.insert(0, rule);
        tools
    }

    /// Data tools with conditional formatting removed from `area` (`None`: the whole sheet).
    pub fn with_cf_cleared(&self, area: Option<Range>) -> crate::data_tools::DataTools {
        let mut tools = self.data_tools().clone();
        tools.cond_formats = match area {
            None => Vec::new(),
            Some(area) => tools
                .cond_formats
                .into_iter()
                .filter_map(|r| {
                    let ranges: Vec<Range> = r.ranges.iter().flat_map(|x| data_tools::subtract(*x, &area)).collect();
                    (!ranges.is_empty()).then_some(CfRule { ranges, ..r })
                })
                .collect(),
        };
        tools
    }

    /// Data tools keeping only the rules at `order` (old indices), in that order.
    /// This is how the Manage Rules dialog deletes and reorders rules in one step.
    pub fn with_cf_order(&self, order: &[usize]) -> crate::data_tools::DataTools {
        let mut tools = self.data_tools().clone();
        let old = std::mem::take(&mut tools.cond_formats);
        tools.cond_formats = order.iter().filter_map(|i| old.get(*i).cloned()).collect();
        tools
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn range(s: &str) -> Vec<Range> {
        data_tools::parse_sqref(s)
    }

    fn sheet_with(values: &[(&str, &str)], rules: Vec<CfRule>) -> Sheet {
        let mut s = Sheet::new();
        for (a, v) in values {
            s.set_input(at(a), v);
        }
        let mut tools = s.data_tools().clone();
        tools.cond_formats = rules;
        s.set_data_tools(tools);
        s
    }

    fn fills(s: &Sheet, block: &str) -> Vec<Option<String>> {
        evaluate(s, range(block)[0]).into_iter().map(|r| r.and_then(|r| r.format.fill)).collect()
    }

    #[test]
    fn highlight_cell_values() {
        let vals = [("A1", "5"), ("A2", "10"), ("A3", "15"), ("A4", "abc")];
        let gt =
            rule_from_spec(range("A1:A4"), "cellIs", &["greaterThan".into(), "9".into()], "lightRedDarkRed").unwrap();
        let s = sheet_with(&vals, vec![gt]);
        let red = Some("#FFC7CE".to_string());
        // Text sorts after numbers, so "abc" > 9 like in Excel.
        assert_eq!(fills(&s, "A1:A4"), vec![None, red.clone(), red.clone(), red.clone()]);
        let between =
            rule_from_spec(range("A1:A4"), "cellIs", &["between".into(), "6".into(), "12".into()], "greenDarkGreen")
                .unwrap();
        let s = sheet_with(&vals, vec![between]);
        assert_eq!(fills(&s, "A1:A3"), vec![None, Some("#C6EFCE".into()), None]);
        let eq = rule_from_spec(range("A1:A4"), "cellIs", &["equal".into(), "ABC".into()], "redText").unwrap();
        let s = sheet_with(&vals, vec![eq]);
        let r = evaluate(&s, range("A4")[0]);
        assert_eq!(r[0].as_ref().unwrap().format.color.as_deref(), Some("#9C0006"));
    }

    #[test]
    fn relative_formulas_follow_the_cell() {
        let vals = [("A1", "5"), ("A2", "10"), ("B1", "4"), ("B2", "11")];
        let rule = rule_from_spec(range("A1:A2"), "cellIs", &["greaterThan".into(), "=B1".into()], "lightRed").unwrap();
        let s = sheet_with(&vals, vec![rule]);
        assert_eq!(fills(&s, "A1:A2"), vec![Some("#FFC7CE".into()), None]);
        let expr = rule_from_spec(range("A1:A2"), "expression", &["=$B1>10".into()], "lightRed").unwrap();
        let s = sheet_with(&vals, vec![expr]);
        assert_eq!(fills(&s, "A1:A2"), vec![None, Some("#FFC7CE".into())]);
    }

    #[test]
    fn text_duplicates_and_dates() {
        let vals = [("A1", "Apple pie"), ("A2", "banana"), ("A3", "apple PIE"), ("A4", "")];
        let s = sheet_with(
            &vals,
            vec![rule_from_spec(range("A1:A4"), "text", &["containsText".into(), "APPLE".into()], "lightRed").unwrap()],
        );
        assert_eq!(fills(&s, "A1:A4"), vec![Some("#FFC7CE".into()), None, Some("#FFC7CE".into()), None]);
        let s = sheet_with(&vals, vec![rule_from_spec(range("A1:A4"), "duplicates", &[], "lightRed").unwrap()]);
        assert_eq!(fills(&s, "A1:A4"), vec![Some("#FFC7CE".into()), None, Some("#FFC7CE".into()), None]);
        let s = sheet_with(&vals, vec![rule_from_spec(range("A1:A4"), "unique", &[], "lightRed").unwrap()]);
        assert_eq!(fills(&s, "A1:A4"), vec![None, Some("#FFC7CE".into()), None, None]);

        let today = date::serial_from_ymd(2025, 8, 15); // a Friday
        assert!(TimePeriod::Today.contains(today + 0.5, today));
        assert!(TimePeriod::Yesterday.contains(today - 1.0, today));
        assert!(TimePeriod::ThisWeek.contains(date::serial_from_ymd(2025, 8, 10), today)); // Sunday
        assert!(!TimePeriod::ThisWeek.contains(date::serial_from_ymd(2025, 8, 9), today));
        assert!(TimePeriod::LastWeek.contains(date::serial_from_ymd(2025, 8, 9), today));
        assert!(TimePeriod::NextMonth.contains(date::serial_from_ymd(2025, 9, 30), today));
        assert!(TimePeriod::Last7Days.contains(today - 6.0, today));
        assert!(!TimePeriod::Last7Days.contains(today - 7.0, today));
    }

    #[test]
    fn top_bottom_and_average() {
        let vals: Vec<(String, String)> = (1..=10).map(|i| (format!("A{i}"), (i * 10).to_string())).collect();
        let vals: Vec<(&str, &str)> = vals.iter().map(|(a, b)| (a.as_str(), b.as_str())).collect();
        let hit = |s: &Sheet| {
            fills(s, "A1:A10").iter().enumerate().filter(|(_, f)| f.is_some()).map(|(i, _)| i + 1).collect::<Vec<_>>()
        };
        let top3 =
            rule_from_spec(range("A1:A10"), "top10", &["top".into(), "items".into(), "3".into()], "lightRed").unwrap();
        assert_eq!(hit(&sheet_with(&vals, vec![top3])), vec![8, 9, 10]);
        let bottom20 =
            rule_from_spec(range("A1:A10"), "top10", &["bottom".into(), "percent".into(), "20".into()], "lightRed")
                .unwrap();
        assert_eq!(hit(&sheet_with(&vals, vec![bottom20])), vec![1, 2]);
        let above = rule_from_spec(range("A1:A10"), "average", &["above".into()], "lightRed").unwrap();
        assert_eq!(hit(&sheet_with(&vals, vec![above])), vec![6, 7, 8, 9, 10]);
        let below = rule_from_spec(range("A1:A10"), "average", &["belowOrEqual".into()], "lightRed").unwrap();
        assert_eq!(hit(&sheet_with(&vals, vec![below])), vec![1, 2, 3, 4, 5]);
    }

    #[test]
    fn scales_and_bars() {
        let vals = [("A1", "0"), ("A2", "50"), ("A3", "100"), ("A4", "x")];
        let scale =
            rule_from_spec(range("A1:A4"), "colorScale", &["#F8696B".into(), "#FFEB84".into(), "#63BE7B".into()], "")
                .unwrap();
        let s = sheet_with(&vals, vec![scale]);
        assert_eq!(
            fills(&s, "A1:A4"),
            vec![Some("#F8696B".into()), Some("#FFEB84".into()), Some("#63BE7B".into()), None]
        );
        let two = rule_from_spec(range("A1:A3"), "colorScale", &["#000000".into(), "#FFFFFF".into()], "").unwrap();
        assert_eq!(fills(&sheet_with(&vals, vec![two]), "A2"), vec![Some("#808080".into())]);
        let bar = rule_from_spec(range("A1:A3"), "dataBar", &["#638EC6".into()], "").unwrap();
        let r = evaluate(&sheet_with(&vals, vec![bar]), range("A1:A3")[0]);
        let lens: Vec<f64> = r.iter().map(|x| x.as_ref().unwrap().bar.as_ref().unwrap().0).collect();
        assert_eq!(lens, vec![0.1, 0.55, 1.0]);
        assert_eq!(r[1].as_ref().unwrap().to_json(), r##"{"bar":0.5500,"barColor":"#638EC6"}"##);
    }

    #[test]
    fn priority_and_stop_if_true() {
        let vals = [("A1", "20")];
        let first =
            CfRule::new(range("A1"), CfKind::Average { below: false, equal: true }, CfFormat::parse("fill=#111111"));
        let second = CfRule::new(
            range("A1"),
            CfKind::CellIs { op: CellOp::Greater, f1: "10".into(), f2: String::new() },
            CfFormat::parse("fill=#222222;b=1;color=#333333"),
        );
        let s = sheet_with(&vals, vec![first.clone(), second.clone()]);
        let r = evaluate(&s, range("A1")[0])[0].clone().unwrap();
        assert_eq!(r.format.fill.as_deref(), Some("#111111"));
        assert_eq!(r.format.bold, Some(true));
        let stop = CfRule { stop_if_true: true, ..first };
        let s = sheet_with(&vals, vec![stop, second]);
        assert_eq!(evaluate(&s, range("A1")[0])[0].clone().unwrap().format.bold, None);
    }

    #[test]
    fn sheet_level_editing() {
        let s = sheet_with(&[], vec![]);
        let r1 = rule_from_spec(range("A1:A10"), "duplicates", &[], "lightRed").unwrap();
        let r2 = rule_from_spec(range("B1:B10"), "unique", &[], "redText").unwrap();
        let mut s = s;
        s.set_data_tools(s.with_cf_rule(r1));
        s.set_data_tools(s.with_cf_rule(r2));
        assert_eq!(s.data_tools().cond_formats[0].describe(), "Unique Values");
        let cleared = s.with_cf_cleared(Some(range("A1:A10")[0]));
        assert_eq!(cleared.cond_formats.len(), 1);
        let reordered = s.with_cf_order(&[1, 0]);
        assert_eq!(reordered.cond_formats[0].describe(), "Duplicate Values");
        assert!(s.with_cf_cleared(None).cond_formats.is_empty());
    }

    #[test]
    fn descriptions_and_specs() {
        let r = rule_from_spec(range("A1"), "cellIs", &["greaterThan".into(), "100".into()], "").unwrap();
        assert_eq!(r.describe(), "Cell Value > 100");
        let r = rule_from_spec(range("A1"), "cellIs", &["equal".into(), "Yes".into()], "").unwrap();
        assert_eq!(r.describe(), "Cell Value = \"Yes\"");
        assert_eq!(criterion("15-08-2025"), "45884");
        assert_eq!(CfFormat::parse("b=1;fill=#FFFF00").to_json(), r##"{"b":1,"fill":"#FFFF00"}"##);
        assert!(rule_from_spec(range("A1"), "top10", &["top".into(), "items".into(), "0".into()], "").is_none());
        let avg = rule_from_spec(range("A1"), "average", &["below".into()], "").unwrap();
        assert_eq!(avg.describe(), "Below Average");
    }
}
