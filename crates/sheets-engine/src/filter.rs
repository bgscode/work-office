//! AutoFilter: the header-row dropdowns of Data → Filter (Ctrl+Shift+L).
//!
//! An [`AutoFilter`] covers a range whose first row holds the headers. Each
//! column may have one [`ColumnFilter`]; a data row is hidden when any
//! column's filter rejects it. As in Excel, rows are hidden when a filter is
//! applied (`apply`) and stay as they are while cells are edited, until the
//! filter is applied again (Data → Reapply).
//!
//! Matching follows Excel: value lists compare the *displayed* text
//! case-insensitively; custom conditions compare numbers numerically and text
//! case-insensitively with `*` / `?` wildcards ("begins with abc" is stored as
//! `= abc*`, exactly like Excel's file format).

use std::cmp::Ordering;
use std::collections::{BTreeMap, BTreeSet, HashSet};

use crate::cell::{CellRef, Range};
use crate::data_tools::{self, has_wildcards, json_str, wildcard_match};
use crate::format;
use crate::formula::refs::Axis;
use crate::sheet::Sheet;
use crate::value::{compare, Value};

#[derive(Debug, Clone, PartialEq)]
pub struct AutoFilter {
    /// Header row plus data rows.
    pub range: Range,
    /// Absolute column index -> its criteria.
    pub columns: BTreeMap<u32, ColumnFilter>,
    /// Last sort done from a dropdown: (column, ascending), shown on the button.
    pub sort: Option<(u32, bool)>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ColumnFilter {
    /// Show rows whose displayed text is one of `values` (compared
    /// case-insensitively), plus empty cells when `blanks` is set.
    Values { values: BTreeSet<String>, blanks: bool },
    /// One or two conditions joined by AND / OR (Text and Number Filters).
    Custom { and: bool, conditions: Vec<Condition> },
    /// Top/Bottom N items or percent (numbers only).
    Top10 { top: bool, percent: bool, value: f64 },
    /// Above / Below Average (numbers only).
    Average { above: bool },
}

#[derive(Debug, Clone, PartialEq)]
pub struct Condition {
    pub op: CompareOp,
    /// What the user typed; may contain `*` and `?`.
    pub value: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CompareOp {
    Equal,
    NotEqual,
    Greater,
    GreaterEqual,
    Less,
    LessEqual,
}

impl CompareOp {
    /// Names used in `.xlsx` (`customFilter operator=`) and by the UI.
    pub fn name(self) -> &'static str {
        match self {
            CompareOp::Equal => "equal",
            CompareOp::NotEqual => "notEqual",
            CompareOp::Greater => "greaterThan",
            CompareOp::GreaterEqual => "greaterThanOrEqual",
            CompareOp::Less => "lessThan",
            CompareOp::LessEqual => "lessThanOrEqual",
        }
    }

    pub fn parse(s: &str) -> Option<CompareOp> {
        Some(match s {
            "equal" | "" => CompareOp::Equal,
            "notEqual" => CompareOp::NotEqual,
            "greaterThan" => CompareOp::Greater,
            "greaterThanOrEqual" => CompareOp::GreaterEqual,
            "lessThan" => CompareOp::Less,
            "lessThanOrEqual" => CompareOp::LessEqual,
            _ => return None,
        })
    }
}

/// A criterion typed by the user as a number, if it is one (`5`, `12%`, `₹500`, `15-08-2025`).
fn criterion_number(s: &str) -> Option<f64> {
    let t = s.trim();
    t.parse::<f64>().ok().filter(|n| n.is_finite()).or_else(|| format::parse_formatted(t).map(|(n, _)| n))
}

impl Condition {
    pub fn matches(&self, value: &Value, display: &str) -> bool {
        let crit = self.value.trim();
        match self.op {
            CompareOp::Equal => equals(crit, value, display),
            CompareOp::NotEqual => !equals(crit, value, display),
            op => {
                let ord = match (value, criterion_number(crit)) {
                    (Value::Number(n), Some(c)) => n.partial_cmp(&c).unwrap_or(Ordering::Equal),
                    (Value::Text(t), None) => t.to_lowercase().cmp(&crit.to_lowercase()),
                    _ => return false,
                };
                match op {
                    CompareOp::Greater => ord == Ordering::Greater,
                    CompareOp::GreaterEqual => ord != Ordering::Less,
                    CompareOp::Less => ord == Ordering::Less,
                    _ => ord != Ordering::Greater,
                }
            }
        }
    }
}

fn equals(crit: &str, value: &Value, display: &str) -> bool {
    if crit.is_empty() {
        return value.is_empty();
    }
    if has_wildcards(crit) {
        return wildcard_match(crit, display);
    }
    if let (Value::Number(n), Some(c)) = (value, criterion_number(crit)) {
        if *n == c {
            return true;
        }
    }
    display.to_lowercase() == crit.to_lowercase()
}

/// Per-column numbers needed by Top 10 and Average filters, computed once per apply.
enum Prepared<'a> {
    Plain(&'a ColumnFilter),
    Threshold { top: bool, cut: f64 },
    Mean { above: bool, mean: f64 },
    Values { lower: HashSet<String>, blanks: bool },
    RejectAll,
}

impl AutoFilter {
    pub fn new(range: Range) -> AutoFilter {
        AutoFilter { range, columns: BTreeMap::new(), sort: None }
    }

    pub fn header_row(&self) -> u32 {
        self.range.start.row
    }

    pub fn data_rows(&self) -> std::ops::RangeInclusive<u32> {
        self.range.start.row + 1..=self.range.end.row
    }

    pub fn is_filtering(&self) -> bool {
        !self.columns.is_empty()
    }

    /// Grow the range down over data typed right below it, like Excel does
    /// when a filter is reapplied.
    pub fn extend_to_data(&mut self, sheet: &Sheet) {
        let (left, right) = (self.range.start.col, self.range.end.col);
        let (rows, _) = sheet.extent();
        while self.range.end.row + 1 < rows
            && (left..=right).any(|c| !sheet.value(CellRef::new(self.range.end.row + 1, c)).is_empty())
        {
            self.range.end.row += 1;
        }
    }

    pub(crate) fn remapped(&self, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Option<AutoFilter> {
        let range = data_tools::remap_range(self.range, axis, map)?;
        let columns = match axis {
            Axis::Row => self.columns.clone(),
            Axis::Col => self.columns.iter().filter_map(|(c, f)| Some((map(*c)?, f.clone()))).collect(),
        };
        let sort = match (axis, self.sort) {
            (Axis::Col, Some((c, asc))) => map(c).map(|c| (c, asc)),
            (_, s) => s,
        };
        Some(AutoFilter { range, columns, sort })
    }

    fn prepare<'a>(&'a self, sheet: &Sheet, col: u32, f: &'a ColumnFilter) -> Prepared<'a> {
        let numbers = || -> Vec<f64> {
            self.data_rows()
                .filter_map(|r| match sheet.value(CellRef::new(r, col)) {
                    Value::Number(n) => Some(*n),
                    _ => None,
                })
                .collect()
        };
        match f {
            ColumnFilter::Top10 { top, percent, value } => {
                let mut nums = numbers();
                if nums.is_empty() {
                    return Prepared::RejectAll;
                }
                nums.sort_by(|a, b| if *top { b.total_cmp(a) } else { a.total_cmp(b) });
                let n = if *percent { (nums.len() as f64 * value / 100.0).floor() } else { value.floor() };
                let n = (n.max(1.0) as usize).min(nums.len());
                Prepared::Threshold { top: *top, cut: nums[n - 1] }
            }
            ColumnFilter::Average { above } => {
                let nums = numbers();
                if nums.is_empty() {
                    return Prepared::RejectAll;
                }
                Prepared::Mean { above: *above, mean: nums.iter().sum::<f64>() / nums.len() as f64 }
            }
            ColumnFilter::Values { values, blanks } => {
                Prepared::Values { lower: values.iter().map(|v| v.to_lowercase()).collect(), blanks: *blanks }
            }
            other => Prepared::Plain(other),
        }
    }

    /// Rows this filter hides right now.
    pub fn hidden_rows(&self, sheet: &Sheet) -> BTreeSet<u32> {
        self.hidden_rows_except(sheet, None)
    }

    fn hidden_rows_except(&self, sheet: &Sheet, skip_col: Option<u32>) -> BTreeSet<u32> {
        let prepared: Vec<(u32, Prepared)> = self
            .columns
            .iter()
            .filter(|(c, _)| Some(**c) != skip_col && **c >= self.range.start.col && **c <= self.range.end.col)
            .map(|(c, f)| (*c, self.prepare(sheet, *c, f)))
            .collect();
        if prepared.is_empty() {
            return BTreeSet::new();
        }
        self.data_rows().filter(|r| prepared.iter().any(|(c, p)| !row_passes(sheet, CellRef::new(*r, *c), p))).collect()
    }

    /// Distinct values of one column for the dropdown checklist, sorted like
    /// Excel (numbers, then text, then booleans; blanks last), each with
    /// whether it is currently shown. Rows hidden by *other* columns' filters
    /// are left out, as in Excel.
    pub fn column_values(&self, sheet: &Sheet, col: u32) -> Vec<FilterItem> {
        let hidden_by_others = self.hidden_rows_except(sheet, Some(col));
        let own = self.columns.get(&col).map(|f| self.prepare(sheet, col, f));
        let mut seen: HashSet<String> = HashSet::new();
        let mut items: Vec<(Value, FilterItem)> = Vec::new();
        let mut blank: Option<FilterItem> = None;
        for r in self.data_rows().filter(|r| !hidden_by_others.contains(r)) {
            let at = CellRef::new(r, col);
            let shown = own.as_ref().is_none_or(|p| row_passes(sheet, at, p));
            let value = sheet.value(at);
            if value.is_empty() {
                let b = blank.get_or_insert(FilterItem { text: String::new(), shown: false, blank: true });
                b.shown |= shown;
                continue;
            }
            let text = sheet.display(at);
            let key = text.to_lowercase();
            if seen.insert(key.clone()) {
                items.push((value.clone(), FilterItem { text, shown, blank: false }));
            } else if shown {
                if let Some((_, item)) = items.iter_mut().find(|(_, i)| i.text.to_lowercase() == key) {
                    item.shown = true;
                }
            }
        }
        items.sort_by(|(a, _), (b, _)| compare(a, b));
        let mut out: Vec<FilterItem> = items.into_iter().map(|(_, i)| i).collect();
        out.extend(blank);
        out
    }

    /// [rows shown, data rows] for the status bar ("4 of 10 records found").
    pub fn counts(&self, hidden: &BTreeSet<u32>) -> (u32, u32) {
        let total = self.range.end.row - self.range.start.row;
        let hidden = hidden.range(self.data_rows()).count() as u32;
        (total - hidden, total)
    }

    /// JSON for the UI: range, filtered columns and sort state.
    pub fn to_json(&self) -> String {
        let r = self.range;
        let cols: Vec<String> = self
            .columns
            .iter()
            .map(|(c, f)| {
                let kind = match f {
                    ColumnFilter::Values { .. } => "values",
                    ColumnFilter::Custom { .. } => "custom",
                    ColumnFilter::Top10 { .. } => "top10",
                    ColumnFilter::Average { .. } => "average",
                };
                format!("\"{c}\":{}", json_str(kind))
            })
            .collect();
        let sort = self.sort.map_or("null".to_string(), |(c, asc)| format!("[{c},{}]", asc as u8));
        format!(
            "{{\"range\":[{},{},{},{}],\"columns\":{{{}}},\"sort\":{sort}}}",
            r.start.row,
            r.start.col,
            r.end.row,
            r.end.col,
            cols.join(",")
        )
    }

    /// Describe one column's criteria for the UI dialogs, e.g.
    /// `{"kind":"custom","and":true,"conditions":[["greaterThan","5"]]}`.
    pub fn column_json(&self, col: u32) -> String {
        match self.columns.get(&col) {
            None => String::new(),
            Some(ColumnFilter::Values { values, blanks }) => format!(
                "{{\"kind\":\"values\",\"blanks\":{blanks},\"values\":[{}]}}",
                values.iter().map(|v| json_str(v)).collect::<Vec<_>>().join(",")
            ),
            Some(ColumnFilter::Custom { and, conditions }) => format!(
                "{{\"kind\":\"custom\",\"and\":{and},\"conditions\":[{}]}}",
                conditions
                    .iter()
                    .map(|c| format!("[{},{}]", json_str(c.op.name()), json_str(&c.value)))
                    .collect::<Vec<_>>()
                    .join(",")
            ),
            Some(ColumnFilter::Top10 { top, percent, value }) => {
                format!("{{\"kind\":\"top10\",\"top\":{top},\"percent\":{percent},\"value\":{value}}}")
            }
            Some(ColumnFilter::Average { above }) => format!("{{\"kind\":\"average\",\"above\":{above}}}"),
        }
    }
}

fn row_passes(sheet: &Sheet, at: CellRef, p: &Prepared) -> bool {
    let value = sheet.value(at);
    match p {
        Prepared::RejectAll => false,
        Prepared::Threshold { top, cut } => match value {
            Value::Number(n) if *top => *n >= *cut,
            Value::Number(n) => *n <= *cut,
            _ => false,
        },
        Prepared::Mean { above, mean } => match value {
            Value::Number(n) if *above => *n > *mean,
            Value::Number(n) => *n < *mean,
            _ => false,
        },
        Prepared::Values { lower, blanks } => {
            if value.is_empty() {
                *blanks
            } else {
                lower.contains(&sheet.display(at).to_lowercase())
            }
        }
        Prepared::Plain(ColumnFilter::Custom { and, conditions }) => {
            let display = sheet.display(at);
            let mut results = conditions.iter().map(|c| c.matches(value, &display));
            if *and {
                results.all(|ok| ok)
            } else {
                results.any(|ok| ok)
            }
        }
        // Values / Top10 / Average are always prepared.
        Prepared::Plain(_) => true,
    }
}

/// One entry of a filter dropdown's checklist.
#[derive(Debug, Clone, PartialEq)]
pub struct FilterItem {
    /// Displayed text ("" for the "(Blanks)" entry).
    pub text: String,
    /// Currently visible (checked).
    pub shown: bool,
    pub blank: bool,
}

// ------------------------------------------------------------ sheet-level actions

impl Sheet {
    /// Data tools with the AutoFilter (re)applied: range grown over new data
    /// and hidden rows recomputed. Returns the tools to store with `SetDataTools`.
    pub fn with_filter_applied(&self, mut filter: AutoFilter) -> crate::data_tools::DataTools {
        filter.extend_to_data(self);
        let hidden = filter.hidden_rows(self);
        crate::data_tools::DataTools { filter_hidden: hidden, auto_filter: Some(filter), ..self.data_tools().clone() }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    /// Name | Qty | City, 6 data rows (one blank city).
    fn table() -> Sheet {
        let mut s = Sheet::new();
        let rows = [
            ["Name", "Qty", "City"],
            ["Asha", "5", "Delhi"],
            ["Ravi", "12", "Mumbai"],
            ["Meera", "7", "delhi"],
            ["John", "30", ""],
            ["Zara", "1", "Pune"],
            ["Amit", "12", "Mumbai"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                s.set_input(CellRef::new(r as u32, c as u32), v);
            }
        }
        s
    }

    fn filter(cols: Vec<(u32, ColumnFilter)>) -> AutoFilter {
        AutoFilter { range: Range::new(at("A1"), at("C7")), columns: cols.into_iter().collect(), sort: None }
    }

    fn rows(set: &BTreeSet<u32>) -> Vec<u32> {
        set.iter().map(|r| r + 1).collect()
    }

    #[test]
    fn filters_by_values_case_insensitively() {
        let s = table();
        let f = filter(vec![(2, ColumnFilter::Values { values: ["DELHI".to_string()].into(), blanks: false })]);
        assert_eq!(rows(&f.hidden_rows(&s)), vec![3, 5, 6, 7]);
        let f = filter(vec![(2, ColumnFilter::Values { values: BTreeSet::new(), blanks: true })]);
        assert_eq!(rows(&f.hidden_rows(&s)), vec![2, 3, 4, 6, 7]);
        assert_eq!(f.counts(&f.hidden_rows(&s)), (1, 6));
    }

    #[test]
    fn custom_conditions() {
        let s = table();
        let cond = |op, v: &str| Condition { op, value: v.into() };
        let f = filter(vec![(1, ColumnFilter::Custom { and: true, conditions: vec![cond(CompareOp::Greater, "6")] })]);
        assert_eq!(rows(&f.hidden_rows(&s)), vec![2, 6]);
        let between = ColumnFilter::Custom {
            and: true,
            conditions: vec![cond(CompareOp::GreaterEqual, "5"), cond(CompareOp::LessEqual, "12")],
        };
        assert_eq!(rows(&filter(vec![(1, between)]).hidden_rows(&s)), vec![5, 6]);
        let begins = ColumnFilter::Custom { and: true, conditions: vec![cond(CompareOp::Equal, "m*")] };
        assert_eq!(rows(&filter(vec![(0, begins)]).hidden_rows(&s)), vec![2, 3, 5, 6, 7]);
        let not_contains = ColumnFilter::Custom { and: true, conditions: vec![cond(CompareOp::NotEqual, "*a*")] };
        assert_eq!(rows(&filter(vec![(0, not_contains)]).hidden_rows(&s)), vec![2, 3, 4, 6, 7]);
        let either = ColumnFilter::Custom {
            and: false,
            conditions: vec![cond(CompareOp::Equal, "pune"), cond(CompareOp::Equal, "Mumbai")],
        };
        assert_eq!(rows(&filter(vec![(2, either)]).hidden_rows(&s)), vec![2, 4, 5]);
    }

    #[test]
    fn top_ten_and_average() {
        let s = table();
        let top2 = filter(vec![(1, ColumnFilter::Top10 { top: true, percent: false, value: 2.0 })]);
        // 30 and both 12s (ties are kept).
        assert_eq!(rows(&top2.hidden_rows(&s)), vec![2, 4, 6]);
        let bottom = filter(vec![(1, ColumnFilter::Top10 { top: false, percent: true, value: 20.0 })]);
        assert_eq!(rows(&bottom.hidden_rows(&s)), vec![2, 3, 4, 5, 7]);
        // Mean of 5,12,7,30,1,12 is 11.17.
        let above = filter(vec![(1, ColumnFilter::Average { above: true })]);
        assert_eq!(rows(&above.hidden_rows(&s)), vec![2, 4, 6]);
    }

    #[test]
    fn combines_columns_and_lists_values() {
        let s = table();
        let f = filter(vec![
            (
                1,
                ColumnFilter::Custom {
                    and: true,
                    conditions: vec![Condition { op: CompareOp::Greater, value: "6".into() }],
                },
            ),
            (2, ColumnFilter::Values { values: ["mumbai".to_string()].into(), blanks: false }),
        ]);
        assert_eq!(rows(&f.hidden_rows(&s)), vec![2, 4, 5, 6]);
        // City list excludes rows hidden by the Qty filter (Asha/Delhi, Zara/Pune).
        let items = f.column_values(&s, 2);
        let texts: Vec<(&str, bool)> = items.iter().map(|i| (i.text.as_str(), i.shown)).collect();
        assert_eq!(texts, vec![("delhi", false), ("Mumbai", true), ("", false)]);
        assert!(items[2].blank);
        let qty = filter(vec![]).column_values(&s, 1);
        assert_eq!(qty.iter().map(|i| i.text.as_str()).collect::<Vec<_>>(), vec!["1", "5", "7", "12", "30"]);
    }

    #[test]
    fn applies_and_extends_over_new_rows() {
        let mut s = table();
        s.set_input(at("A8"), "Neha");
        s.set_input(at("C8"), "Delhi");
        let f = filter(vec![(2, ColumnFilter::Values { values: ["pune".to_string()].into(), blanks: false })]);
        let tools = s.with_filter_applied(f);
        assert_eq!(tools.auto_filter.as_ref().unwrap().range, Range::new(at("A1"), at("C8")));
        assert_eq!(rows(&tools.filter_hidden), vec![2, 3, 4, 5, 7, 8]);
        s.set_data_tools(tools);
        assert!(s.is_filter_hidden(1));
        assert!(!s.is_filter_hidden(5));
        // Inserting a row above the table moves the hidden rows with it.
        let moved = s.with_inserted(Axis::Row, 0, 1);
        assert!(moved.is_filter_hidden(2));
        assert_eq!(moved.data_tools().auto_filter.as_ref().unwrap().range, Range::new(at("A2"), at("C9")));
    }

    #[test]
    fn describes_itself_as_json() {
        let f = filter(vec![(1, ColumnFilter::Top10 { top: true, percent: false, value: 10.0 })]);
        assert_eq!(f.to_json(), r#"{"range":[0,0,6,2],"columns":{"1":"top10"},"sort":null}"#);
        assert_eq!(f.column_json(1), r#"{"kind":"top10","top":true,"percent":false,"value":10}"#);
        assert_eq!(f.column_json(0), "");
    }
}
