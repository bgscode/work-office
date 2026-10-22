//! Data validation (Data → Data Validation): which entries a cell accepts,
//! the in-cell dropdown list, the input message shown when the cell is
//! selected and the error alert shown for invalid entries.
//!
//! Criteria are stored the way `.xlsx` stores them: `formula1`/`formula2`
//! hold formula text without the leading `=` (`10`, `$A$1`, `"Yes,No"`,
//! `$D$1:$D$5`). They are evaluated against the sheet, relative references
//! being relative to the top-left cell of the rule's first range.

use crate::cell::{CellRef, Range};
use crate::data_tools::{self, json_str};
use crate::format;
use crate::formula::{self, refs, Expr};
use crate::sheet::Sheet;
use crate::value::Value;

#[derive(Debug, Clone, PartialEq)]
pub struct Validation {
    pub ranges: Vec<Range>,
    pub kind: ValidationKind,
    pub operator: ValidationOp,
    pub formula1: String,
    pub formula2: String,
    /// "Ignore blank": an empty cell (or a blank list source cell) is always fine.
    pub allow_blank: bool,
    /// Show the dropdown arrow for List rules.
    pub in_cell_dropdown: bool,
    pub show_input: bool,
    pub input_title: String,
    pub input_message: String,
    pub show_error: bool,
    pub error_style: ErrorStyle,
    pub error_title: String,
    pub error_message: String,
}

impl Default for Validation {
    fn default() -> Self {
        Validation {
            ranges: Vec::new(),
            kind: ValidationKind::Any,
            operator: ValidationOp::Between,
            formula1: String::new(),
            formula2: String::new(),
            allow_blank: true,
            in_cell_dropdown: true,
            show_input: true,
            input_title: String::new(),
            input_message: String::new(),
            show_error: true,
            error_style: ErrorStyle::Stop,
            error_title: String::new(),
            error_message: String::new(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValidationKind {
    Any,
    Whole,
    Decimal,
    List,
    Date,
    Time,
    TextLength,
    Custom,
}

impl ValidationKind {
    /// Names used in `.xlsx` (`type=`) and by the UI.
    pub fn name(self) -> &'static str {
        match self {
            ValidationKind::Any => "none",
            ValidationKind::Whole => "whole",
            ValidationKind::Decimal => "decimal",
            ValidationKind::List => "list",
            ValidationKind::Date => "date",
            ValidationKind::Time => "time",
            ValidationKind::TextLength => "textLength",
            ValidationKind::Custom => "custom",
        }
    }

    pub fn parse(s: &str) -> ValidationKind {
        match s {
            "whole" => ValidationKind::Whole,
            "decimal" => ValidationKind::Decimal,
            "list" => ValidationKind::List,
            "date" => ValidationKind::Date,
            "time" => ValidationKind::Time,
            "textLength" => ValidationKind::TextLength,
            "custom" => ValidationKind::Custom,
            _ => ValidationKind::Any,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValidationOp {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    Greater,
    Less,
    GreaterEqual,
    LessEqual,
}

impl ValidationOp {
    pub fn name(self) -> &'static str {
        match self {
            ValidationOp::Between => "between",
            ValidationOp::NotBetween => "notBetween",
            ValidationOp::Equal => "equal",
            ValidationOp::NotEqual => "notEqual",
            ValidationOp::Greater => "greaterThan",
            ValidationOp::Less => "lessThan",
            ValidationOp::GreaterEqual => "greaterThanOrEqual",
            ValidationOp::LessEqual => "lessThanOrEqual",
        }
    }

    pub fn parse(s: &str) -> ValidationOp {
        match s {
            "notBetween" => ValidationOp::NotBetween,
            "equal" => ValidationOp::Equal,
            "notEqual" => ValidationOp::NotEqual,
            "greaterThan" => ValidationOp::Greater,
            "lessThan" => ValidationOp::Less,
            "greaterThanOrEqual" => ValidationOp::GreaterEqual,
            "lessThanOrEqual" => ValidationOp::LessEqual,
            _ => ValidationOp::Between,
        }
    }

    fn needs_two(self) -> bool {
        matches!(self, ValidationOp::Between | ValidationOp::NotBetween)
    }

    /// Does `x` satisfy this operator against `a` (and `b` for between)?
    pub fn test(self, x: f64, a: f64, b: f64) -> bool {
        let (lo, hi) = if a <= b { (a, b) } else { (b, a) };
        match self {
            ValidationOp::Between => x >= lo && x <= hi,
            ValidationOp::NotBetween => x < lo || x > hi,
            ValidationOp::Equal => x == a,
            ValidationOp::NotEqual => x != a,
            ValidationOp::Greater => x > a,
            ValidationOp::Less => x < a,
            ValidationOp::GreaterEqual => x >= a,
            ValidationOp::LessEqual => x <= a,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorStyle {
    /// Retry / Cancel: the entry is never accepted.
    Stop,
    /// "Continue?" Yes / No / Cancel.
    Warning,
    /// OK / Cancel: OK keeps the entry.
    Information,
}

impl ErrorStyle {
    pub fn name(self) -> &'static str {
        match self {
            ErrorStyle::Stop => "stop",
            ErrorStyle::Warning => "warning",
            ErrorStyle::Information => "information",
        }
    }

    pub fn parse(s: &str) -> ErrorStyle {
        match s {
            "warning" => ErrorStyle::Warning,
            "information" => ErrorStyle::Information,
            _ => ErrorStyle::Stop,
        }
    }
}

/// Excel's message when a rule has no custom error text.
pub const DEFAULT_ERROR: &str = "This value doesn't match the data validation restrictions defined for this cell.";

impl Validation {
    pub fn applies_to(&self, at: CellRef) -> bool {
        self.ranges.iter().any(|r| data_tools::contains(r, at))
    }

    fn origin(&self) -> CellRef {
        self.ranges.first().map_or(CellRef::new(0, 0), |r| r.start)
    }

    /// Parse `formula` as seen from cell `at`.
    fn expr_at(&self, formula: &str, at: CellRef) -> Option<Expr> {
        let f = formula.trim();
        if f.is_empty() {
            return None;
        }
        let o = self.origin();
        let shifted = refs::shift_input(&format!("={f}"), at.row as i64 - o.row as i64, at.col as i64 - o.col as i64);
        formula::parse(&shifted[1..]).ok()
    }

    /// A criterion as a number: formula results, or typed dates/times like Excel's dialog accepts.
    fn bound(&self, sheet: &Sheet, formula: &str, at: CellRef) -> Option<f64> {
        let f = formula.trim();
        if let Some((n, _)) = format::parse_formatted(f) {
            return Some(n);
        }
        if let Some(t) = parse_time(f) {
            return Some(t);
        }
        formula::eval(&self.expr_at(f, at)?, sheet).as_number().ok()
    }

    /// The items of a List rule, as shown in the dropdown.
    pub fn list_items(&self, sheet: &Sheet, at: CellRef) -> Vec<String> {
        let f = self.formula1.trim();
        if let Some(inner) = f.strip_prefix('"').and_then(|s| s.strip_suffix('"')) {
            return inner
                .replace("\"\"", "\"")
                .split(',')
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty())
                .collect();
        }
        match self.expr_at(f, at) {
            Some(Expr::Range(r)) => r.cells().map(|c| sheet.display(c)).filter(|s| !s.is_empty()).collect(),
            Some(Expr::Ref(c)) => vec![sheet.display(c)].into_iter().filter(|s| !s.is_empty()).collect(),
            // A bare comma list typed without quotes (what the dialog's Source box holds).
            _ if !f.is_empty() => f.split(',').map(|s| s.trim().to_string()).filter(|s| !s.is_empty()).collect(),
            _ => Vec::new(),
        }
    }

    /// Would typing `input` into `at` be accepted?
    pub fn accepts(&self, sheet: &Sheet, at: CellRef, input: &str) -> bool {
        if input.is_empty() {
            return true;
        }
        let value = match input.strip_prefix('=') {
            Some(src) => {
                formula::parse(src).map_or(Value::Error(crate::ErrorKind::Parse), |e| formula::eval(&e, sheet))
            }
            None if sheet.style(at).number_format == crate::NumberFormat::Text => Value::Text(input.to_string()),
            None => Value::from_input(input).or_time(input),
        };
        if value.is_empty() {
            return self.allow_blank;
        }
        let bounds = |sheet: &Sheet| -> Option<(f64, f64)> {
            let a = self.bound(sheet, &self.formula1, at)?;
            let b = if self.operator.needs_two() { self.bound(sheet, &self.formula2, at)? } else { a };
            Some((a, b))
        };
        match self.kind {
            ValidationKind::Any => true,
            ValidationKind::List => {
                let text = value.display().to_lowercase();
                // Lists compare what is displayed: "5" matches a list item "5".
                let shown = format::format_value(&value, &sheet.style(at).number_format).to_lowercase();
                self.list_items(sheet, at).iter().any(|i| {
                    let i = i.to_lowercase();
                    i == text || i == shown
                })
            }
            ValidationKind::Whole | ValidationKind::Decimal | ValidationKind::Date | ValidationKind::Time => {
                let Value::Number(n) = value else { return false };
                if self.kind == ValidationKind::Whole && n.fract() != 0.0 {
                    return false;
                }
                if self.kind == ValidationKind::Date && n < 0.0 {
                    return false;
                }
                let x = if self.kind == ValidationKind::Time { n.fract() } else { n };
                bounds(sheet).is_some_and(|(a, b)| self.operator.test(x, a, b))
            }
            ValidationKind::TextLength => {
                let len = value.display().chars().count() as f64;
                bounds(sheet).is_some_and(|(a, b)| self.operator.test(len, a, b))
            }
            ValidationKind::Custom => {
                // The formula is evaluated with the new value in place.
                let mut probe = sheet.clone();
                probe.set_input(at, input);
                self.expr_at(&self.formula1, at).is_some_and(|e| formula::eval(&e, &probe).as_bool().unwrap_or(false))
            }
        }
    }

    /// JSON for the Data Validation dialog and the error alert.
    pub fn to_json(&self) -> String {
        format!(
            concat!(
                "{{\"kind\":{},\"operator\":{},\"formula1\":{},\"formula2\":{},\"allowBlank\":{},",
                "\"dropdown\":{},\"showInput\":{},\"inputTitle\":{},\"inputMessage\":{},\"showError\":{},",
                "\"errorStyle\":{},\"errorTitle\":{},\"errorMessage\":{},\"ranges\":{}}}"
            ),
            json_str(self.kind.name()),
            json_str(self.operator.name()),
            json_str(&self.formula1),
            json_str(&self.formula2),
            self.allow_blank,
            self.in_cell_dropdown,
            self.show_input,
            json_str(&self.input_title),
            json_str(&self.input_message),
            self.show_error,
            json_str(self.error_style.name()),
            json_str(&self.error_title),
            json_str(&self.error_message),
            json_str(&data_tools::format_sqref(&self.ranges)),
        )
    }
}

/// `14:30` or `2:30 PM` as a fraction of a day.
fn parse_time(s: &str) -> Option<f64> {
    let lower = s.trim().to_ascii_lowercase();
    let (body, pm) = match (lower.strip_suffix("pm"), lower.strip_suffix("am")) {
        (Some(b), _) => (b.trim().to_string(), Some(true)),
        (_, Some(b)) => (b.trim().to_string(), Some(false)),
        _ => (lower.clone(), None),
    };
    let mut parts = body.split(':');
    let h: f64 = parts.next()?.trim().parse().ok()?;
    let m: f64 = parts.next()?.trim().parse().ok()?;
    let sec: f64 = parts.next().map_or(Some(0.0), |p| p.trim().parse().ok())?;
    if parts.next().is_some() || !(0.0..60.0).contains(&m) || !(0.0..60.0).contains(&sec) {
        return None;
    }
    let h = match pm {
        Some(true) if h < 12.0 => h + 12.0,
        Some(false) if h == 12.0 => 0.0,
        _ => h,
    };
    (0.0..24.0).contains(&h).then(|| (h * 3600.0 + m * 60.0 + sec) / 86400.0)
}

trait OrTime {
    fn or_time(self, input: &str) -> Value;
}

impl OrTime for Value {
    /// Typed times ("14:30") count as numbers for Time rules.
    fn or_time(self, input: &str) -> Value {
        match (&self, parse_time(input)) {
            (Value::Text(_), Some(t)) => Value::Number(t),
            _ => self,
        }
    }
}

impl Sheet {
    /// The validation rule for a cell (the first one listed wins on overlap).
    pub fn validation_at(&self, at: CellRef) -> Option<&Validation> {
        self.data_tools().validations.iter().find(|v| v.applies_to(at))
    }

    /// The rule `input` breaks if typed into `at`, or `None` when it's allowed.
    pub fn check_input(&self, at: CellRef, input: &str) -> Option<&Validation> {
        self.validation_at(at).filter(|v| v.show_error && !v.accepts(self, at, input))
    }

    /// Data tools with `rule` set for `area`, replacing whatever validation
    /// those cells had. A rule of kind Any with no input message clears it.
    pub fn with_validation(&self, area: Range, rule: Option<Validation>) -> crate::data_tools::DataTools {
        let mut tools = self.data_tools().clone();
        tools.validations = tools
            .validations
            .into_iter()
            .filter_map(|v| {
                let ranges: Vec<Range> = v.ranges.iter().flat_map(|r| data_tools::subtract(*r, &area)).collect();
                (!ranges.is_empty()).then_some(Validation { ranges, ..v })
            })
            .collect();
        if let Some(rule) = rule {
            let keep =
                rule.kind != ValidationKind::Any || !rule.input_message.is_empty() || !rule.input_title.is_empty();
            if keep {
                tools.validations.insert(0, Validation { ranges: vec![area], ..rule });
            }
        }
        tools
    }

    /// Cells in `area` whose current value breaks their validation rule ("Circle Invalid Data").
    pub fn invalid_cells(&self, area: Range) -> Vec<CellRef> {
        area.cells()
            .filter(|at| {
                let input = self.input(*at);
                !input.is_empty() && self.validation_at(*at).is_some_and(|v| !v.accepts(self, *at, &input))
            })
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn rule(kind: ValidationKind, op: ValidationOp, f1: &str, f2: &str) -> Validation {
        Validation {
            ranges: vec![Range::new(at("A1"), at("A10"))],
            kind,
            operator: op,
            formula1: f1.into(),
            formula2: f2.into(),
            ..Validation::default()
        }
    }

    #[test]
    fn whole_and_decimal_numbers() {
        let s = Sheet::new();
        let whole = rule(ValidationKind::Whole, ValidationOp::Between, "1", "10");
        assert!(whole.accepts(&s, at("A1"), "5"));
        assert!(!whole.accepts(&s, at("A1"), "5.5"));
        assert!(!whole.accepts(&s, at("A1"), "11"));
        assert!(!whole.accepts(&s, at("A1"), "abc"));
        assert!(whole.accepts(&s, at("A1"), ""));
        let dec = rule(ValidationKind::Decimal, ValidationOp::Greater, "0.5", "");
        assert!(dec.accepts(&s, at("A1"), "0.75"));
        assert!(!dec.accepts(&s, at("A1"), "0.5"));
        assert!(dec.accepts(&s, at("A1"), "=1/1"));
    }

    #[test]
    fn bounds_can_reference_cells() {
        let mut s = Sheet::new();
        s.set_input(at("C1"), "100");
        let r = rule(ValidationKind::Decimal, ValidationOp::LessEqual, "$C$1", "");
        assert!(r.accepts(&s, at("A5"), "100"));
        assert!(!r.accepts(&s, at("A5"), "101"));
    }

    #[test]
    fn lists_from_text_and_ranges() {
        let mut s = Sheet::new();
        for (a, v) in [("D1", "Red"), ("D2", "Green"), ("D3", ""), ("D4", "Blue")] {
            s.set_input(at(a), v);
        }
        let inline = rule(ValidationKind::List, ValidationOp::Between, "\"Yes,No, Maybe\"", "");
        assert_eq!(inline.list_items(&s, at("A1")), vec!["Yes", "No", "Maybe"]);
        assert!(inline.accepts(&s, at("A1"), "yes"));
        assert!(!inline.accepts(&s, at("A1"), "Perhaps"));
        let range = rule(ValidationKind::List, ValidationOp::Between, "$D$1:$D$4", "");
        assert_eq!(range.list_items(&s, at("A3")), vec!["Red", "Green", "Blue"]);
        assert!(range.accepts(&s, at("A3"), "Blue"));
        assert!(!range.accepts(&s, at("A3"), "Pink"));
    }

    #[test]
    fn dates_times_and_text_length() {
        let s = Sheet::new();
        let date = rule(ValidationKind::Date, ValidationOp::GreaterEqual, "01-01-2025", "");
        assert!(date.accepts(&s, at("A1"), "15-08-2025"));
        assert!(!date.accepts(&s, at("A1"), "31-12-2024"));
        let time = rule(ValidationKind::Time, ValidationOp::Between, "09:00", "17:00");
        assert!(time.accepts(&s, at("A1"), "14:30"));
        assert!(!time.accepts(&s, at("A1"), "18:00"));
        let len = rule(ValidationKind::TextLength, ValidationOp::LessEqual, "6", "");
        assert!(len.accepts(&s, at("A1"), "नमस्ते")); // 6 characters, not 18 bytes
        assert!(len.accepts(&s, at("A1"), "hello"));
        assert!(!len.accepts(&s, at("A1"), "toolong"));
    }

    #[test]
    fn custom_formula_rules() {
        let s = Sheet::new();
        let r = rule(ValidationKind::Custom, ValidationOp::Between, "A1>B1", "");
        let mut s2 = s.clone();
        s2.set_input(at("B2"), "10");
        assert!(r.accepts(&s2, at("A2"), "11"));
        assert!(!r.accepts(&s2, at("A2"), "9"));
    }

    #[test]
    fn sheet_level_rules() {
        let mut s = Sheet::new();
        let area = Range::new(at("A1"), at("A10"));
        let tools = s.with_validation(area, Some(rule(ValidationKind::Whole, ValidationOp::Greater, "0", "")));
        s.set_data_tools(tools);
        assert!(s.check_input(at("A3"), "-1").is_some());
        assert!(s.check_input(at("A3"), "1").is_none());
        assert!(s.check_input(at("B3"), "-1").is_none());
        // Clearing the middle splits the rule.
        let tools = s.with_validation(Range::new(at("A4"), at("A5")), None);
        s.set_data_tools(tools);
        assert!(s.validation_at(at("A4")).is_none());
        assert!(s.validation_at(at("A3")).is_some() && s.validation_at(at("A6")).is_some());
        s.set_input(at("A7"), "-5");
        assert_eq!(s.invalid_cells(area), vec![at("A7")]);
    }

    #[test]
    fn json_for_the_dialog() {
        let r = rule(ValidationKind::List, ValidationOp::Between, "\"a,b\"", "");
        assert!(r.to_json().starts_with(r#"{"kind":"list","operator":"between","formula1":"\"a,b\"""#));
        assert!(r.to_json().ends_with(r#""ranges":"A1:A10"}"#));
    }
}
