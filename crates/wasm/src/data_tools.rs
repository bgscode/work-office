//! Bindings for the Data tools: AutoFilter, Data Validation, Conditional
//! Formatting, Remove Duplicates and multi-level Sort (`sheets_engine::
//! {filter, validation, conditional_format, dedupe, sort}`). Same conventions
//! as `lib.rs`: rectangles are (top, left, bottom, right), zero-based; JSON
//! comes back as strings, "" when there is nothing.

use sheets_engine::conditional_format::{self, rule_from_spec};
use sheets_engine::data_tools::{current_region, json_str, DataTools, SetDataTools};
use sheets_engine::filter::{AutoFilter, ColumnFilter, CompareOp, Condition};
use sheets_engine::ops::ReplaceSheet;
use sheets_engine::sort::SortKey;
use sheets_engine::validation::{ErrorStyle, Validation, ValidationKind, ValidationOp};
use sheets_engine::{CellRef, Range};
use wasm_bindgen::prelude::*;

use crate::{rect, SheetsDocument};

impl SheetsDocument {
    fn set_tools(&mut self, sheet: usize, tools: DataTools) {
        if &tools != self.sheet(sheet).data_tools() {
            self.apply(SetDataTools { sheet, tools });
        }
    }

    /// Change one column's criteria and reapply the filter.
    fn set_column_filter(&mut self, sheet: usize, col: u32, filter: Option<ColumnFilter>) {
        let s = self.sheet(sheet);
        let Some(mut f) = s.data_tools().auto_filter.clone() else { return };
        match filter {
            Some(c) => f.columns.insert(col, c),
            None => f.columns.remove(&col),
        };
        let tools = s.with_filter_applied(f);
        self.set_tools(sheet, tools);
    }
}

#[wasm_bindgen]
impl SheetsDocument {
    /// Excel's current region around a cell (the table Ctrl+A or Filter picks), as [top, left, bottom, right].
    #[wasm_bindgen(js_name = currentRegion)]
    pub fn current_region(&self, sheet: usize, row: u32, col: u32) -> Vec<u32> {
        let r = current_region(self.sheet(sheet), CellRef::new(row, col));
        vec![r.start.row, r.start.col, r.end.row, r.end.col]
    }

    // ------------------------------------------------------------ AutoFilter

    /// `{"range":[t,l,b,r],"columns":{"2":"values"},"sort":[col,asc]|null}`, or "" without a filter.
    #[wasm_bindgen(js_name = autoFilter)]
    pub fn auto_filter(&self, sheet: usize) -> String {
        self.sheet(sheet).data_tools().auto_filter.as_ref().map_or(String::new(), AutoFilter::to_json)
    }

    /// Data → Filter (Ctrl+Shift+L): turn the AutoFilter off, or on over the
    /// selection (a single cell means the table around it).
    #[wasm_bindgen(js_name = toggleFilter)]
    pub fn toggle_filter(&mut self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32) {
        let s = self.sheet(sheet);
        if s.data_tools().auto_filter.is_some() {
            let tools = DataTools { auto_filter: None, filter_hidden: Default::default(), ..s.data_tools().clone() };
            self.set_tools(sheet, tools);
            return;
        }
        let mut range = rect(top, left, bottom, right);
        if range.rows() == 1 && range.cols() == 1 {
            range = current_region(s, range.start);
        }
        // A header row with no data below still gets dropdowns.
        let tools = s.with_filter_applied(AutoFilter::new(range));
        self.set_tools(sheet, tools);
    }

    /// The checklist of a column's dropdown as JSON: `[{"text":"Delhi","shown":true,"blank":false}, ...]`.
    #[wasm_bindgen(js_name = filterItems)]
    pub fn filter_items(&self, sheet: usize, col: u32) -> String {
        let s = self.sheet(sheet);
        let Some(f) = &s.data_tools().auto_filter else { return "[]".into() };
        let items: Vec<String> = f
            .column_values(s, col)
            .iter()
            .map(|i| format!("{{\"text\":{},\"shown\":{},\"blank\":{}}}", json_str(&i.text), i.shown, i.blank))
            .collect();
        format!("[{}]", items.join(","))
    }

    /// One column's criteria (see `AutoFilter::column_json`), "" when unfiltered.
    #[wasm_bindgen(js_name = filterColumn)]
    pub fn filter_column(&self, sheet: usize, col: u32) -> String {
        self.sheet(sheet).data_tools().auto_filter.as_ref().map_or(String::new(), |f| f.column_json(col))
    }

    /// Show only rows whose displayed text is one of `values` (plus blanks).
    #[wasm_bindgen(js_name = filterByValues)]
    pub fn filter_by_values(&mut self, sheet: usize, col: u32, values: Vec<String>, blanks: bool) {
        let filter = ColumnFilter::Values { values: values.into_iter().collect(), blanks };
        self.set_column_filter(sheet, col, Some(filter));
    }

    /// Text/Number Filters: one or two conditions; `ops` use the `.xlsx` names
    /// (`equal`, `notEqual`, `greaterThan`, ...), values may hold `*` and `?`.
    #[wasm_bindgen(js_name = filterByConditions)]
    pub fn filter_by_conditions(&mut self, sheet: usize, col: u32, and: bool, ops: Vec<String>, values: Vec<String>) {
        let conditions: Vec<Condition> = ops
            .iter()
            .zip(values)
            .filter(|(_, v)| !v.is_empty())
            .filter_map(|(op, value)| Some(Condition { op: CompareOp::parse(op)?, value }))
            .collect();
        let filter = (!conditions.is_empty()).then_some(ColumnFilter::Custom { and, conditions });
        self.set_column_filter(sheet, col, filter);
    }

    #[wasm_bindgen(js_name = filterTop10)]
    pub fn filter_top10(&mut self, sheet: usize, col: u32, top: bool, percent: bool, value: f64) {
        self.set_column_filter(sheet, col, Some(ColumnFilter::Top10 { top, percent, value }));
    }

    #[wasm_bindgen(js_name = filterAverage)]
    pub fn filter_average(&mut self, sheet: usize, col: u32, above: bool) {
        self.set_column_filter(sheet, col, Some(ColumnFilter::Average { above }));
    }

    /// "Clear Filter From" one column.
    #[wasm_bindgen(js_name = clearColumnFilter)]
    pub fn clear_column_filter(&mut self, sheet: usize, col: u32) {
        self.set_column_filter(sheet, col, None);
    }

    /// Data → Clear: keep the dropdowns, show every row.
    #[wasm_bindgen(js_name = clearFilters)]
    pub fn clear_filters(&mut self, sheet: usize) {
        let s = self.sheet(sheet);
        let Some(f) = &s.data_tools().auto_filter else { return };
        let tools = s.with_filter_applied(AutoFilter { columns: Default::default(), sort: None, ..f.clone() });
        self.set_tools(sheet, tools);
    }

    /// Data → Reapply: recompute hidden rows after edits.
    #[wasm_bindgen(js_name = reapplyFilter)]
    pub fn reapply_filter(&mut self, sheet: usize) {
        let s = self.sheet(sheet);
        let Some(f) = s.data_tools().auto_filter.clone() else { return };
        let tools = s.with_filter_applied(f);
        self.set_tools(sheet, tools);
    }

    /// Sort A→Z / Z→A from a filter dropdown: sorts the data rows under the
    /// header and remembers the sort for the button's arrow. One undo step.
    #[wasm_bindgen(js_name = filterSort)]
    pub fn filter_sort(&mut self, sheet: usize, col: u32, ascending: bool) {
        let s = self.sheet(sheet);
        let Some(mut f) = s.data_tools().auto_filter.clone() else { return };
        f.extend_to_data(s);
        if f.range.rows() < 2 {
            return;
        }
        let data = Range::new(CellRef::new(f.range.start.row + 1, f.range.start.col), f.range.end);
        let mut sorted = s.with_sorted_rows_by(data, &[SortKey { col, ascending }]);
        f.sort = Some((col, ascending));
        let tools = sorted.with_filter_applied(f);
        sorted.set_data_tools(tools);
        self.apply(ReplaceSheet::new(sheet, sorted));
    }

    /// [rows shown, data rows] for "4 of 10 records found", empty while nothing is filtered.
    #[wasm_bindgen(js_name = filterCounts)]
    pub fn filter_counts(&self, sheet: usize) -> Vec<u32> {
        let s = self.sheet(sheet);
        match &s.data_tools().auto_filter {
            Some(f) if f.is_filtering() => {
                let (shown, total) = f.counts(s.filter_hidden_rows());
                vec![shown, total]
            }
            _ => Vec::new(),
        }
    }

    // ------------------------------------------------------------ Data Validation

    /// The rule for a cell as JSON (see `Validation::to_json`), "" when none.
    #[wasm_bindgen(js_name = validationAt)]
    pub fn validation_at(&self, sheet: usize, row: u32, col: u32) -> String {
        self.sheet(sheet).validation_at(CellRef::new(row, col)).map_or(String::new(), Validation::to_json)
    }

    /// Set the validation of a rectangle (Data → Data Validation → OK).
    /// `kind` "any" with no input message clears it.
    #[allow(clippy::too_many_arguments)]
    #[wasm_bindgen(js_name = setValidation)]
    pub fn set_validation(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        kind: &str,
        operator: &str,
        formula1: String,
        formula2: String,
        allow_blank: bool,
        dropdown: bool,
        show_input: bool,
        input_title: String,
        input_message: String,
        show_error: bool,
        error_style: &str,
        error_title: String,
        error_message: String,
    ) {
        let rule = Validation {
            ranges: Vec::new(),
            kind: ValidationKind::parse(kind),
            operator: ValidationOp::parse(operator),
            formula1: formula1.trim_start_matches('=').to_string(),
            formula2: formula2.trim_start_matches('=').to_string(),
            allow_blank,
            in_cell_dropdown: dropdown,
            show_input,
            input_title,
            input_message,
            show_error,
            error_style: ErrorStyle::parse(error_style),
            error_title,
            error_message,
        };
        let tools = self.sheet(sheet).with_validation(rect(top, left, bottom, right), Some(rule));
        self.set_tools(sheet, tools);
    }

    /// Data Validation → Clear All for a rectangle.
    #[wasm_bindgen(js_name = clearValidation)]
    pub fn clear_validation(&mut self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32) {
        let tools = self.sheet(sheet).with_validation(rect(top, left, bottom, right), None);
        self.set_tools(sheet, tools);
    }

    /// The rule `input` would break in that cell as JSON, "" when it's allowed.
    /// The UI shows the error alert (Stop / Warning / Information) from it.
    #[wasm_bindgen(js_name = checkInput)]
    pub fn check_input(&self, sheet: usize, row: u32, col: u32, input: &str) -> String {
        self.sheet(sheet).check_input(CellRef::new(row, col), input).map_or(String::new(), Validation::to_json)
    }

    /// Choices of an in-cell dropdown (List rules), empty when the cell has none.
    #[wasm_bindgen(js_name = listItems)]
    pub fn list_items(&self, sheet: usize, row: u32, col: u32) -> Vec<String> {
        let s = self.sheet(sheet);
        let at = CellRef::new(row, col);
        match s.validation_at(at) {
            Some(v) if v.kind == ValidationKind::List && v.in_cell_dropdown => v.list_items(s, at),
            _ => Vec::new(),
        }
    }

    /// Circle Invalid Data: cells breaking their rule, as [row, col, row, col, ...].
    #[wasm_bindgen(js_name = invalidCells)]
    pub fn invalid_cells(&self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32) -> Vec<u32> {
        let s = self.sheet(sheet);
        let (rows, cols) = s.extent();
        if rows == 0 || cols == 0 {
            return Vec::new();
        }
        let area = Range::new(
            CellRef::new(top, left),
            CellRef::new(bottom.min(rows - 1).max(top), right.min(cols - 1).max(left)),
        );
        s.invalid_cells(area).into_iter().flat_map(|at| [at.row, at.col]).collect()
    }

    // ------------------------------------------------------------ Conditional Formatting

    /// Add a rule over a rectangle; see `conditional_format::rule_from_spec`
    /// for `kind` and `args`. `format` is a preset name or `fill=#..;color=#..;b=1`.
    /// Returns false when the rule is incomplete.
    #[allow(clippy::too_many_arguments)]
    #[wasm_bindgen(js_name = addCfRule)]
    pub fn add_cf_rule(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        kind: &str,
        args: Vec<String>,
        format: &str,
    ) -> bool {
        let Some(rule) = rule_from_spec(vec![rect(top, left, bottom, right)], kind, &args, format) else {
            return false;
        };
        let tools = self.sheet(sheet).with_cf_rule(rule);
        self.set_tools(sheet, tools);
        true
    }

    /// Clear Rules from Selected Cells (`whole` false) or from the Entire Sheet.
    #[wasm_bindgen(js_name = clearCfRules)]
    pub fn clear_cf_rules(&mut self, sheet: usize, whole: bool, top: u32, left: u32, bottom: u32, right: u32) {
        let area = (!whole).then(|| rect(top, left, bottom, right));
        let tools = self.sheet(sheet).with_cf_cleared(area);
        self.set_tools(sheet, tools);
    }

    /// The sheet's rules in priority order, each as JSON (see `CfRule::to_json`).
    #[wasm_bindgen(js_name = cfRules)]
    pub fn cf_rules(&self, sheet: usize) -> Vec<String> {
        self.sheet(sheet).data_tools().cond_formats.iter().map(|r| r.to_json()).collect()
    }

    /// Manage Rules → OK: keep the rules at `order` (old indices), in that order.
    #[wasm_bindgen(js_name = setCfOrder)]
    pub fn set_cf_order(&mut self, sheet: usize, order: Vec<u32>) {
        let order: Vec<usize> = order.into_iter().map(|i| i as usize).collect();
        let tools = self.sheet(sheet).with_cf_order(&order);
        self.set_tools(sheet, tools);
    }

    /// Conditional formatting of a block, row-major: JSON with style keys plus
    /// `bar`/`barColor` for data bars, "" where no rule applies. Empty when the
    /// sheet has no rules at all, so painting skips it cheaply.
    #[wasm_bindgen(js_name = cfBlock)]
    pub fn cf_block(&self, sheet: usize, row: u32, col: u32, rows: u32, cols: u32) -> Vec<String> {
        let s = self.sheet(sheet);
        if s.data_tools().cond_formats.is_empty() || rows == 0 || cols == 0 {
            return Vec::new();
        }
        let block = Range::new(CellRef::new(row, col), CellRef::new(row + rows - 1, col + cols - 1));
        // Range statistics are worked out once per change, not on every paint.
        let prepared = {
            let mut cache = self.cf_cache.borrow_mut();
            match &*cache {
                Some((sh, generation, p)) if *sh == sheet && *generation == self.generation => p.clone(),
                _ => {
                    let p = std::rc::Rc::new(conditional_format::prepare(s));
                    *cache = Some((sheet, self.generation, p.clone()));
                    p
                }
            }
        };
        conditional_format::evaluate_prepared(s, &prepared, block)
            .into_iter()
            .map(|r| r.map_or(String::new(), |r| r.to_json()))
            .collect()
    }

    // ------------------------------------------------------------ Remove Duplicates, Sort

    /// Data → Remove Duplicates. `cols` are absolute columns to compare.
    /// Returns [removed, remaining unique].
    #[wasm_bindgen(js_name = removeDuplicates)]
    #[allow(clippy::too_many_arguments)]
    pub fn remove_duplicates(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        cols: Vec<u32>,
        has_header: bool,
    ) -> Vec<u32> {
        let d = self.sheet(sheet).with_duplicates_removed(rect(top, left, bottom, right), &cols, has_header);
        if d.removed > 0 {
            self.apply(ReplaceSheet::new(sheet, d.sheet));
        }
        vec![d.removed, d.remaining]
    }

    /// Data → Sort (custom, several levels). The rectangle is the data rows
    /// only; `cols[i]` is sorted ascending when `ascending[i]` is 1.
    #[wasm_bindgen(js_name = sortBy)]
    #[allow(clippy::too_many_arguments)]
    pub fn sort_by(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        cols: Vec<u32>,
        ascending: Vec<u8>,
    ) {
        let keys: Vec<SortKey> = cols
            .iter()
            .zip(ascending.iter().chain(std::iter::repeat(&1)))
            .map(|(c, a)| SortKey { col: *c, ascending: *a != 0 })
            .collect();
        if keys.is_empty() || top >= bottom {
            return;
        }
        let new = self.sheet(sheet).with_sorted_rows_by(rect(top, left, bottom, right), &keys);
        self.apply(ReplaceSheet::new(sheet, new));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Name | City with 4 data rows, header in row 0.
    fn doc() -> SheetsDocument {
        let mut d = SheetsDocument::new();
        let cells = [["Name", "City"], ["Asha", "Delhi"], ["Ravi", "Pune"], ["asha", "DELHI"], ["Meera", "Delhi"]];
        let (mut rows, mut cols, mut inputs) = (Vec::new(), Vec::new(), Vec::new());
        for (r, row) in cells.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                rows.push(r as u32);
                cols.push(c as u32);
                inputs.push(v.to_string());
            }
        }
        d.set_inputs(0, rows, cols, inputs);
        d
    }

    #[test]
    fn filter_hides_rows_and_undoes() {
        let mut d = doc();
        d.toggle_filter(0, 1, 1, 1, 1);
        assert!(d.auto_filter(0).starts_with(r#"{"range":[0,0,4,1]"#));
        d.filter_by_values(0, 1, vec!["Delhi".into()], false);
        assert!(d.sheet(0).is_row_hidden(2) && !d.sheet(0).is_row_hidden(3));
        assert_eq!(d.filter_counts(0), vec![3, 4]);
        assert!(d.filter_items(0, 1).contains(r#"{"text":"Pune","shown":false"#));
        d.undo();
        assert!(!d.sheet(0).is_row_hidden(2));
        d.filter_sort(0, 0, false);
        assert_eq!(d.input(0, 1, 0), "Ravi");
        d.toggle_filter(0, 0, 0, 0, 0);
        assert_eq!(d.auto_filter(0), "");
    }

    #[test]
    fn validation_cf_and_duplicates() {
        let mut d = doc();
        d.set_validation(
            0,
            1,
            1,
            9,
            1,
            "list",
            "between",
            "\"Delhi,Pune\"".into(),
            String::new(),
            true,
            true,
            false,
            String::new(),
            String::new(),
            true,
            "stop",
            String::new(),
            String::new(),
        );
        assert_eq!(d.list_items(0, 2, 1), vec!["Delhi", "Pune"]);
        assert!(d.check_input(0, 2, 1, "Agra").contains("\"kind\":\"list\""));
        assert_eq!(d.check_input(0, 2, 1, "Pune"), "");

        assert!(d.cf_block(0, 0, 0, 2, 2).is_empty());
        assert!(d.add_cf_rule(0, 1, 1, 4, 1, "text", vec!["containsText".into(), "pune".into()], "lightRedDarkRed"));
        let block = d.cf_block(0, 1, 1, 2, 1);
        assert_eq!(block[0], "");
        assert!(block[1].contains("#FFC7CE"));

        assert_eq!(d.remove_duplicates(0, 0, 0, 4, 1, vec![0, 1], true), vec![1, 3]);
        assert_eq!(d.input(0, 3, 0), "Meera");
    }
}
