//! `.xlsx` parts for the Data tools: `<autoFilter>`, `<conditionalFormatting>`
//! and `<dataValidations>` in each worksheet, and the `<dxfs>` table in
//! `styles.xml` that conditional formats point at with `dxfId`.
//!
//! A child module of `xlsx.rs` so it can share its XML helpers. Worksheet
//! element order matters to Excel: `sheetData`, `autoFilter`, (`mergeCells`),
//! `conditionalFormatting`, `dataValidations`, then page setup.

use sheets_engine::conditional_format::{CellOp, CfFormat, CfKind, CfRule, Cfvo, TextOp, TimePeriod};
use sheets_engine::data_tools::{format_sqref, parse_sqref};
use sheets_engine::filter::{AutoFilter, ColumnFilter, CompareOp, Condition};
use sheets_engine::validation::{ErrorStyle, Validation, ValidationKind, ValidationOp};
use sheets_engine::{CellRef, Range, Sheet};

use super::{escape, StyleTable};
use crate::xml::{attr, walk, XmlEvent};

// ---------------------------------------------------------------- reading

/// The `<dxfs>` of `styles.xml`, in order.
pub(super) fn read_dxfs(xml: &str) -> Vec<CfFormat> {
    let mut out = Vec::new();
    let (mut in_dxfs, mut in_font, mut in_fill) = (false, false, false);
    walk(xml, |ev| match ev {
        XmlEvent::Start { name: "dxfs", .. } => in_dxfs = true,
        XmlEvent::End { name: "dxfs" } => in_dxfs = false,
        XmlEvent::Start { name: "dxf", .. } if in_dxfs => out.push(CfFormat::default()),
        XmlEvent::Start { name: "font", .. } if in_dxfs => in_font = true,
        XmlEvent::End { name: "font" } => in_font = false,
        XmlEvent::Start { name: "fill", .. } if in_dxfs => in_fill = true,
        XmlEvent::End { name: "fill" } => in_fill = false,
        XmlEvent::Start { name, attrs } if in_dxfs => {
            let Some(f) = out.last_mut() else { return };
            let on = || Some(!matches!(&*attr(attrs, "val"), "0" | "false"));
            match *name {
                "b" if in_font => f.bold = on(),
                "i" if in_font => f.italic = on(),
                "strike" if in_font => f.strike = on(),
                "u" if in_font => f.underline = Some(attr(attrs, "val") != "none"),
                "color" if in_font => f.color = rgb(&attr(attrs, "rgb")),
                // Solid dxf fills keep their colour in bgColor; fgColor is a fallback.
                "bgColor" if in_fill => f.fill = rgb(&attr(attrs, "rgb")).or(f.fill.take()),
                "fgColor" if in_fill && f.fill.is_none() => f.fill = rgb(&attr(attrs, "rgb")),
                _ => {}
            }
        }
        _ => {}
    });
    out
}

fn rgb(s: &str) -> Option<String> {
    super::argb(s)
}

/// Everything the Data tools keep in one worksheet part.
#[derive(Default)]
pub(super) struct SheetTools {
    pub filter: Option<AutoFilter>,
    pub validations: Vec<Validation>,
    pub cond_formats: Vec<CfRule>,
}

#[derive(Clone, Copy, PartialEq, Default)]
enum Text {
    #[default]
    None,
    Formula1,
    Formula2,
    CfFormula,
}

/// Reads the AutoFilter, validations and conditional formats while
/// `xlsx::sheet_data` walks the worksheet, so a big sheet is read in one pass.
#[derive(Default)]
pub(super) struct ToolsReader {
    out: SheetTools,
    /// AutoFilter column being read: (column, filter, unsupported).
    column: Option<(u32, Option<ColumnFilter>, bool)>,
    validation: Option<Validation>,
    cf_ranges: Vec<Range>,
    rule: Option<RuleDraft>,
    cf_rules: Vec<(i64, CfRule)>,
    text: Text,
}

impl ToolsReader {
    pub(super) fn event(&mut self, ev: &XmlEvent, dxfs: &[CfFormat]) {
        let out = &mut self.out;
        let column = &mut self.column;
        let validation = &mut self.validation;
        let rule = &mut self.rule;
        let text = &mut self.text;
        match ev {
            // ---- autoFilter
            XmlEvent::Start { name: "autoFilter", attrs } => {
                out.filter = parse_sqref(&attr(attrs, "ref")).first().map(|r| AutoFilter::new(*r));
            }
            XmlEvent::Start { name: "filterColumn", attrs } => {
                if let (Some(f), Ok(id)) = (&out.filter, attr(attrs, "colId").parse::<u32>()) {
                    *column = Some((f.range.start.col + id, None, false));
                }
            }
            XmlEvent::Start { name: "filters", attrs } => {
                if let Some((_, slot, _)) = column {
                    *slot = Some(ColumnFilter::Values { values: Default::default(), blanks: attr(attrs, "blank") == "1" });
                }
            }
            XmlEvent::Start { name: "filter", attrs } => {
                if let Some((_, Some(ColumnFilter::Values { values, .. }), _)) = column {
                    values.insert(attr(attrs, "val").into_owned());
                }
            }
            // Date grouping isn't supported; keep such columns unfiltered rather than hiding everything.
            XmlEvent::Start { name: "dateGroupItem" | "iconFilter" | "colorFilter", .. } => {
                if let Some((_, _, unsupported)) = column {
                    *unsupported = true;
                }
            }
            XmlEvent::Start { name: "customFilters", attrs } => {
                if let Some((_, slot, _)) = column {
                    *slot = Some(ColumnFilter::Custom { and: attr(attrs, "and") == "1", conditions: Vec::new() });
                }
            }
            XmlEvent::Start { name: "customFilter", attrs } => {
                if let Some((_, Some(ColumnFilter::Custom { conditions, .. }), _)) = column {
                    if let Some(op) = CompareOp::parse(&attr(attrs, "operator")) {
                        conditions.push(Condition { op, value: attr(attrs, "val").into_owned() });
                    }
                }
            }
            XmlEvent::Start { name: "top10", attrs } => {
                if let Some((_, slot, _)) = column {
                    *slot = Some(ColumnFilter::Top10 {
                        top: attr(attrs, "top") != "0",
                        percent: attr(attrs, "percent") == "1",
                        value: attr(attrs, "val").parse().unwrap_or(10.0),
                    });
                }
            }
            XmlEvent::Start { name: "dynamicFilter", attrs } => {
                if let Some((_, slot, unsupported)) = column {
                    match &*attr(attrs, "type") {
                        "aboveAverage" => *slot = Some(ColumnFilter::Average { above: true }),
                        "belowAverage" => *slot = Some(ColumnFilter::Average { above: false }),
                        _ => *unsupported = true,
                    }
                }
            }
            XmlEvent::End { name: "filterColumn" } => {
                if let (Some((col, Some(f), false)), Some(filter)) = (column.take(), &mut out.filter) {
                    filter.columns.insert(col, f);
                }
            }

            // ---- dataValidations
            XmlEvent::Start { name: "dataValidation", attrs } => {
                let flag = |k: &str| matches!(&*attr(attrs, k), "1" | "true");
                *validation = Some(Validation {
                    ranges: parse_sqref(&attr(attrs, "sqref")),
                    kind: ValidationKind::parse(&attr(attrs, "type")),
                    operator: ValidationOp::parse(&attr(attrs, "operator")),
                    allow_blank: flag("allowBlank"),
                    // Excel's showDropDown="1" means *hide* the in-cell dropdown.
                    in_cell_dropdown: !flag("showDropDown"),
                    show_input: flag("showInputMessage"),
                    input_title: attr(attrs, "promptTitle").into_owned(),
                    input_message: attr(attrs, "prompt").into_owned(),
                    show_error: flag("showErrorMessage"),
                    error_style: ErrorStyle::parse(&attr(attrs, "errorStyle")),
                    error_title: attr(attrs, "errorTitle").into_owned(),
                    error_message: attr(attrs, "error").into_owned(),
                    ..Validation::default()
                });
            }
            XmlEvent::Start { name: "formula1", .. } if validation.is_some() => *text = Text::Formula1,
            XmlEvent::Start { name: "formula2", .. } if validation.is_some() => *text = Text::Formula2,
            XmlEvent::End { name: "formula1" | "formula2" } => *text = Text::None,
            XmlEvent::End { name: "dataValidation" } => {
                if let Some(v) = validation.take().filter(|v| !v.ranges.is_empty()) {
                    out.validations.push(v);
                }
            }

            // ---- conditionalFormatting
            XmlEvent::Start { name: "conditionalFormatting", attrs } => {
                self.cf_ranges = parse_sqref(&attr(attrs, "sqref"))
            }
            XmlEvent::Start { name: "cfRule", attrs } => {
                let attrs = ["type", "operator", "text", "timePeriod", "bottom", "percent", "rank", "aboveAverage",
                    "equalAverage", "dxfId", "stopIfTrue", "priority"]
                    .iter()
                    .filter_map(|k| attrs.get(k).map(|v| (k.to_string(), v.into_owned())))
                    .collect();
                *rule = Some(RuleDraft { attrs, formulas: Vec::new(), cfvos: Vec::new(), colors: Vec::new() });
            }
            XmlEvent::Start { name: "formula", .. } if rule.is_some() => {
                *text = Text::CfFormula;
                if let Some(r) = rule {
                    r.formulas.push(String::new());
                }
            }
            XmlEvent::End { name: "formula" } => *text = Text::None,
            XmlEvent::Start { name: "cfvo", attrs } => {
                if let Some(r) = rule {
                    r.cfvos.push(Cfvo::parse(&attr(attrs, "type"), &attr(attrs, "val")));
                }
            }
            XmlEvent::Start { name: "color", attrs } if rule.is_some() => {
                if let Some(r) = rule {
                    r.colors.push(rgb(&attr(attrs, "rgb")).unwrap_or_else(|| "#638EC6".into()));
                }
            }
            XmlEvent::End { name: "cfRule" } => {
                if let Some(draft) = rule.take() {
                    if let Some(r) = draft.build(&self.cf_ranges, dxfs) {
                        let priority = draft.attr("priority").parse().unwrap_or(i64::MAX);
                        self.cf_rules.push((priority, r));
                    }
                }
            }

            XmlEvent::Text(t) => match *text {
                Text::Formula1 => validation.as_mut().unwrap().formula1.push_str(t),
                Text::Formula2 => validation.as_mut().unwrap().formula2.push_str(t),
                Text::CfFormula => {
                    if let Some(f) = rule.as_mut().and_then(|r| r.formulas.last_mut()) {
                        f.push_str(t)
                    }
                }
                Text::None => {}
            },
            _ => {}
        }
    }

    pub(super) fn finish(mut self) -> SheetTools {
        self.cf_rules.sort_by_key(|(p, _)| *p);
        self.out.cond_formats = self.cf_rules.into_iter().map(|(_, r)| r).collect();
        self.out
    }
}

struct RuleDraft {
    attrs: Vec<(String, String)>,
    formulas: Vec<String>,
    cfvos: Vec<Cfvo>,
    colors: Vec<String>,
}

impl RuleDraft {
    fn attr(&self, key: &str) -> &str {
        self.attrs.iter().find(|(k, _)| k == key).map_or("", |(_, v)| v)
    }

    fn build(&self, ranges: &[Range], dxfs: &[CfFormat]) -> Option<CfRule> {
        let a = |k: &str| self.attr(k);
        let formula = |i: usize| self.formulas.get(i).cloned().unwrap_or_default();
        let kind = match a("type") {
            "cellIs" => CfKind::CellIs { op: CellOp::parse(a("operator"))?, f1: formula(0), f2: formula(1) },
            t @ ("containsText" | "notContainsText" | "beginsWith" | "endsWith") => {
                CfKind::Text { op: TextOp::parse(t)?, text: a("text").to_string() }
            }
            "timePeriod" => CfKind::TimePeriod(TimePeriod::parse(a("timePeriod"))?),
            "duplicateValues" => CfKind::Duplicates { unique: false },
            "uniqueValues" => CfKind::Duplicates { unique: true },
            "top10" => CfKind::Top10 {
                bottom: a("bottom") == "1",
                percent: a("percent") == "1",
                rank: a("rank").parse().unwrap_or(10),
            },
            "aboveAverage" => CfKind::Average { below: a("aboveAverage") == "0", equal: a("equalAverage") == "1" },
            "expression" => CfKind::Expression { formula: formula(0) },
            "colorScale" if self.cfvos.len() >= 2 && self.colors.len() == self.cfvos.len() => {
                CfKind::ColorScale { points: self.cfvos.iter().copied().zip(self.colors.iter().cloned()).collect() }
            }
            "dataBar" if self.cfvos.len() >= 2 => CfKind::DataBar {
                min: self.cfvos[0],
                max: self.cfvos[1],
                color: self.colors.first().cloned().unwrap_or_else(|| "#638EC6".into()),
            },
            _ => return None,
        };
        let format = a("dxfId").parse::<usize>().ok().and_then(|i| dxfs.get(i)).cloned().unwrap_or_default();
        Some(CfRule { ranges: ranges.to_vec(), kind, format, stop_if_true: a("stopIfTrue") == "1" })
    }
}

/// Install what `read_sheet` found, recomputing the rows the filter hides.
pub(super) fn apply(sheet: &mut Sheet, found: SheetTools) {
    let mut tools = sheet.data_tools().clone();
    tools.validations = found.validations;
    tools.cond_formats = found.cond_formats;
    sheet.set_data_tools(tools);
    if let Some(f) = found.filter {
        let tools = sheet.with_filter_applied(f);
        sheet.set_data_tools(tools);
    }
}

// ---------------------------------------------------------------- writing

impl StyleTable {
    /// The `dxfId` for a conditional format.
    fn dxf_index(&mut self, f: &CfFormat) -> usize {
        super::position_or_push(&mut self.dxfs, f.clone())
    }

    /// `<dxfs>` for `styles.xml` (empty when no conditional formats were written).
    pub(super) fn dxfs_xml(&self) -> String {
        if self.dxfs.is_empty() {
            return String::new();
        }
        let mut xml = format!(r#"<dxfs count="{}">"#, self.dxfs.len());
        for f in &self.dxfs {
            xml.push_str("<dxf>");
            let has_font = f.bold.is_some() || f.italic.is_some() || f.underline.is_some() || f.strike.is_some();
            if has_font || f.color.is_some() {
                xml.push_str("<font>");
                for (tag, v) in [("b", f.bold), ("i", f.italic), ("strike", f.strike)] {
                    match v {
                        Some(true) => xml.push_str(&format!("<{tag}/>")),
                        Some(false) => xml.push_str(&format!(r#"<{tag} val="0"/>"#)),
                        None => {}
                    }
                }
                match f.underline {
                    Some(true) => xml.push_str("<u/>"),
                    Some(false) => xml.push_str(r#"<u val="none"/>"#),
                    None => {}
                }
                if let Some(c) = &f.color {
                    xml.push_str(&format!(r#"<color rgb="FF{}"/>"#, c.trim_start_matches('#')));
                }
                xml.push_str("</font>");
            }
            if let Some(c) = &f.fill {
                xml.push_str(&format!(
                    r#"<fill><patternFill><bgColor rgb="FF{}"/></patternFill></fill>"#,
                    c.trim_start_matches('#')
                ));
            }
            xml.push_str("</dxf>");
        }
        xml.push_str("</dxfs>");
        xml
    }
}

/// `<autoFilter>`; goes right after `</sheetData>`.
pub(super) fn autofilter_xml(sheet: &Sheet) -> String {
    let Some(f) = &sheet.data_tools().auto_filter else { return String::new() };
    if f.columns.is_empty() {
        return format!(r#"<autoFilter ref="{}"/>"#, f.range);
    }
    let mut xml = format!(r#"<autoFilter ref="{}">"#, f.range);
    for (col, filter) in &f.columns {
        let id = col.saturating_sub(f.range.start.col);
        xml.push_str(&format!(r#"<filterColumn colId="{id}">"#));
        match filter {
            ColumnFilter::Values { values, blanks } => {
                xml.push_str(if *blanks { r#"<filters blank="1">"# } else { "<filters>" });
                for v in values {
                    xml.push_str(&format!(r#"<filter val="{}"/>"#, escape(v)));
                }
                xml.push_str("</filters>");
            }
            ColumnFilter::Custom { and, conditions } => {
                xml.push_str(if *and { r#"<customFilters and="1">"# } else { "<customFilters>" });
                for c in conditions {
                    let op = if c.op == CompareOp::Equal {
                        String::new()
                    } else {
                        format!(r#" operator="{}""#, c.op.name())
                    };
                    xml.push_str(&format!(r#"<customFilter{op} val="{}"/>"#, escape(&c.value)));
                }
                xml.push_str("</customFilters>");
            }
            ColumnFilter::Top10 { top, percent, value } => {
                let top = if *top { "" } else { r#" top="0""# };
                let percent = if *percent { r#" percent="1""# } else { "" };
                xml.push_str(&format!(r#"<top10{top}{percent} val="{value}"/>"#));
            }
            ColumnFilter::Average { above } => {
                let t = if *above { "aboveAverage" } else { "belowAverage" };
                xml.push_str(&format!(r#"<dynamicFilter type="{t}"/>"#));
            }
        }
        xml.push_str("</filterColumn>");
    }
    xml.push_str("</autoFilter>");
    xml
}

/// `<conditionalFormatting>` elements then `<dataValidations>`; they follow
/// `<mergeCells>` (if any) and come before page setup.
pub(super) fn cf_and_validation_xml(sheet: &Sheet, styles: &mut StyleTable) -> String {
    let tools = sheet.data_tools();
    let mut xml = String::new();
    for (i, rule) in tools.cond_formats.iter().enumerate() {
        xml.push_str(&format!(r#"<conditionalFormatting sqref="{}">"#, format_sqref(&rule.ranges)));
        xml.push_str(&cf_rule_xml(rule, i + 1, styles));
        xml.push_str("</conditionalFormatting>");
    }
    if !tools.validations.is_empty() {
        xml.push_str(&format!(r#"<dataValidations count="{}">"#, tools.validations.len()));
        for v in &tools.validations {
            xml.push_str(&validation_xml(v));
        }
        xml.push_str("</dataValidations>");
    }
    xml
}

fn cf_rule_xml(rule: &CfRule, priority: usize, styles: &mut StyleTable) -> String {
    let stop = if rule.stop_if_true { r#" stopIfTrue="1""# } else { "" };
    let mut dxf = || format!(r#" dxfId="{}""#, styles.dxf_index(&rule.format));
    let origin = rule.ranges.first().map_or(CellRef::new(0, 0), |r| r.start).to_string();
    let formula = |f: &str| format!("<formula>{}</formula>", escape(f));
    match &rule.kind {
        CfKind::CellIs { op, f1, f2 } => {
            let mut body = formula(f1);
            if matches!(op, CellOp::Between | CellOp::NotBetween) {
                body.push_str(&formula(f2));
            }
            format!(
                r#"<cfRule type="cellIs"{} priority="{priority}"{stop} operator="{}">{body}</cfRule>"#,
                dxf(),
                op.name()
            )
        }
        CfKind::Text { op, text } => {
            let q = format!("\"{}\"", text.replace('"', "\"\""));
            let (operator, f) = match op {
                TextOp::Contains => ("containsText", format!("NOT(ISERROR(SEARCH({q},{origin})))")),
                TextOp::NotContains => ("notContains", format!("ISERROR(SEARCH({q},{origin}))")),
                TextOp::BeginsWith => ("beginsWith", format!("LEFT({origin},LEN({q}))={q}")),
                TextOp::EndsWith => ("endsWith", format!("RIGHT({origin},LEN({q}))={q}")),
            };
            format!(
                r#"<cfRule type="{}"{} priority="{priority}"{stop} operator="{operator}" text="{}">{}</cfRule>"#,
                op.rule_type(),
                dxf(),
                escape(text),
                formula(&f)
            )
        }
        CfKind::TimePeriod(p) => format!(
            r#"<cfRule type="timePeriod"{} priority="{priority}"{stop} timePeriod="{}">{}</cfRule>"#,
            dxf(),
            p.name(),
            formula(&time_period_formula(*p, &origin))
        ),
        CfKind::Duplicates { unique } => {
            let t = if *unique { "uniqueValues" } else { "duplicateValues" };
            format!(r#"<cfRule type="{t}"{} priority="{priority}"{stop}/>"#, dxf())
        }
        CfKind::Top10 { bottom, percent, rank } => {
            let b = if *bottom { r#" bottom="1""# } else { "" };
            let p = if *percent { r#" percent="1""# } else { "" };
            format!(r#"<cfRule type="top10"{} priority="{priority}"{stop} rank="{rank}"{p}{b}/>"#, dxf())
        }
        CfKind::Average { below, equal } => {
            let b = if *below { r#" aboveAverage="0""# } else { "" };
            let e = if *equal { r#" equalAverage="1""# } else { "" };
            format!(r#"<cfRule type="aboveAverage"{} priority="{priority}"{stop}{b}{e}/>"#, dxf())
        }
        CfKind::Expression { formula: f } => {
            format!(r#"<cfRule type="expression"{} priority="{priority}"{stop}>{}</cfRule>"#, dxf(), formula(f))
        }
        CfKind::ColorScale { points } => {
            let mut body = String::from("<colorScale>");
            for (p, _) in points {
                body.push_str(&cfvo_xml(p));
            }
            for (_, c) in points {
                body.push_str(&format!(r#"<color rgb="FF{}"/>"#, c.trim_start_matches('#')));
            }
            body.push_str("</colorScale>");
            format!(r#"<cfRule type="colorScale" priority="{priority}"{stop}>{body}</cfRule>"#)
        }
        CfKind::DataBar { min, max, color } => format!(
            r#"<cfRule type="dataBar" priority="{priority}"{stop}><dataBar>{}{}<color rgb="FF{}"/></dataBar></cfRule>"#,
            cfvo_xml(min),
            cfvo_xml(max),
            color.trim_start_matches('#')
        ),
    }
}

fn cfvo_xml(p: &Cfvo) -> String {
    match p.value() {
        Some(v) => format!(r#"<cfvo type="{}" val="{v}"/>"#, p.kind()),
        None => format!(r#"<cfvo type="{}"/>"#, p.kind()),
    }
}

/// The formulas Excel itself writes for "A Date Occurring" rules.
fn time_period_formula(p: TimePeriod, a: &str) -> String {
    match p {
        TimePeriod::Today => format!("FLOOR({a},1)=TODAY()"),
        TimePeriod::Yesterday => format!("FLOOR({a},1)=TODAY()-1"),
        TimePeriod::Tomorrow => format!("FLOOR({a},1)=TODAY()+1"),
        TimePeriod::Last7Days => format!("AND(TODAY()-FLOOR({a},1)<=6,FLOOR({a},1)<=TODAY())"),
        TimePeriod::ThisWeek => {
            format!("AND(TODAY()-ROUNDDOWN({a},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({a},0)-TODAY()<=7-WEEKDAY(TODAY()))")
        }
        TimePeriod::LastWeek => {
            format!("AND(TODAY()-ROUNDDOWN({a},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({a},0)<(WEEKDAY(TODAY())+7))")
        }
        TimePeriod::NextWeek => {
            format!("AND(ROUNDDOWN({a},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({a},0)-TODAY()<(15-WEEKDAY(TODAY())))")
        }
        TimePeriod::ThisMonth => format!("AND(MONTH({a})=MONTH(TODAY()),YEAR({a})=YEAR(TODAY()))"),
        TimePeriod::LastMonth => {
            format!("AND(MONTH({a})=MONTH(EDATE(TODAY(),0-1)),YEAR({a})=YEAR(EDATE(TODAY(),0-1)))")
        }
        TimePeriod::NextMonth => {
            format!("AND(MONTH({a})=MONTH(EDATE(TODAY(),0+1)),YEAR({a})=YEAR(EDATE(TODAY(),0+1)))")
        }
    }
}

fn validation_xml(v: &Validation) -> String {
    let mut a = String::new();
    if v.kind != ValidationKind::Any {
        a.push_str(&format!(r#" type="{}""#, v.kind.name()));
    }
    if v.error_style != ErrorStyle::Stop {
        a.push_str(&format!(r#" errorStyle="{}""#, v.error_style.name()));
    }
    let uses_operator = !matches!(v.kind, ValidationKind::Any | ValidationKind::List | ValidationKind::Custom);
    if uses_operator && v.operator != ValidationOp::Between {
        a.push_str(&format!(r#" operator="{}""#, v.operator.name()));
    }
    for (key, on) in [
        ("allowBlank", v.allow_blank),
        ("showDropDown", v.kind == ValidationKind::List && !v.in_cell_dropdown),
        ("showInputMessage", v.show_input),
        ("showErrorMessage", v.show_error),
    ] {
        if on {
            a.push_str(&format!(r#" {key}="1""#));
        }
    }
    for (key, text) in [
        ("errorTitle", &v.error_title),
        ("error", &v.error_message),
        ("promptTitle", &v.input_title),
        ("prompt", &v.input_message),
    ] {
        if !text.is_empty() {
            a.push_str(&format!(r#" {key}="{}""#, escape(text)));
        }
    }
    let mut body = String::new();
    if !v.formula1.is_empty() {
        body.push_str(&format!("<formula1>{}</formula1>", escape(&v.formula1)));
    }
    if !v.formula2.is_empty() && matches!(v.operator, ValidationOp::Between | ValidationOp::NotBetween) {
        body.push_str(&format!("<formula2>{}</formula2>", escape(&v.formula2)));
    }
    format!(r#"<dataValidation{a} sqref="{}">{body}</dataValidation>"#, format_sqref(&v.ranges))
}

/// Column filters keyed by absolute column, for tests.
#[cfg(test)]
pub(super) fn filter_columns(sheet: &Sheet) -> std::collections::BTreeMap<u32, ColumnFilter> {
    sheet.data_tools().auto_filter.as_ref().map(|f| f.columns.clone()).unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::super::{read, write};
    use super::*;
    use sheets_engine::conditional_format::rule_from_spec;
    use sheets_engine::Workbook;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn range(s: &str) -> Vec<Range> {
        parse_sqref(s)
    }

    #[test]
    fn round_trips_data_tools() {
        let mut book = Workbook::new();
        let s = &mut book.sheets[0].sheet;
        for (r, (name, qty)) in
            [("Name", "Qty"), ("Asha", "5"), ("Ravi", "12"), ("Meera", "7"), ("Asha", "30")].iter().enumerate()
        {
            s.set_input(CellRef::new(r as u32, 0), name);
            s.set_input(CellRef::new(r as u32, 1), qty);
        }
        let mut filter = AutoFilter::new(Range::new(at("A1"), at("B5")));
        filter.columns.insert(0, ColumnFilter::Values { values: ["Asha".to_string()].into(), blanks: true });
        filter.columns.insert(
            1,
            ColumnFilter::Custom {
                and: true,
                conditions: vec![
                    Condition { op: CompareOp::GreaterEqual, value: "5".into() },
                    Condition { op: CompareOp::NotEqual, value: "*0".into() },
                ],
            },
        );
        let tools = s.with_filter_applied(filter);
        s.set_data_tools(tools);
        let list = Validation {
            ranges: range("C2:C5 E1"),
            kind: ValidationKind::List,
            formula1: "\"Yes,No\"".into(),
            input_title: "Pick".into(),
            input_message: "Yes or <No>".into(),
            error_style: ErrorStyle::Warning,
            error_message: "Bad & wrong".into(),
            in_cell_dropdown: false,
            ..Validation::default()
        };
        let whole = Validation {
            ranges: range("D2:D5"),
            kind: ValidationKind::Whole,
            operator: ValidationOp::GreaterEqual,
            formula1: "0".into(),
            show_input: false,
            ..Validation::default()
        };
        let mut tools = s.data_tools().clone();
        tools.validations = vec![list.clone(), whole.clone()];
        let rules = vec![
            rule_from_spec(range("B2:B5"), "cellIs", &["greaterThan".into(), "10".into()], "lightRedDarkRed").unwrap(),
            rule_from_spec(range("A2:A5"), "duplicates", &[], "b=1;color=#0000FF").unwrap(),
            rule_from_spec(range("A2:A5"), "text", &["beginsWith".into(), "M".into()], "greenDarkGreen").unwrap(),
            rule_from_spec(range("B2:B5"), "top10", &["bottom".into(), "percent".into(), "25".into()], "redText")
                .unwrap(),
            rule_from_spec(range("B2:B5"), "average", &["belowOrEqual".into()], "lightRed").unwrap(),
            rule_from_spec(range("B2:B5"), "timePeriod", &["lastMonth".into()], "lightRed").unwrap(),
            rule_from_spec(range("B2:B5"), "colorScale", &["#F8696B".into(), "#FFEB84".into(), "#63BE7B".into()], "")
                .unwrap(),
            rule_from_spec(range("B2:B5"), "dataBar", &["#638EC6".into()], "").unwrap(),
            CfRule {
                stop_if_true: true,
                ..rule_from_spec(range("A1"), "expression", &["=$B1>1".into()], "i=1").unwrap()
            },
        ];
        tools.cond_formats = rules.clone();
        s.set_data_tools(tools);
        let hidden: Vec<u32> = s.filter_hidden_rows().iter().copied().collect();
        assert_eq!(hidden, vec![2, 3, 4]);

        let bytes = write(&book).unwrap();
        let back = read(&bytes).unwrap();
        let t = back.sheets[0].sheet.data_tools();
        assert_eq!(filter_columns(&back.sheets[0].sheet), filter_columns(&book.sheets[0].sheet));
        assert_eq!(t.auto_filter.as_ref().unwrap().range, Range::new(at("A1"), at("B5")));
        assert_eq!(t.filter_hidden.iter().copied().collect::<Vec<_>>(), hidden);
        assert_eq!(t.validations, vec![list, whole]);
        assert_eq!(t.cond_formats, rules);
    }

    #[test]
    fn writes_excel_structure() {
        let mut book = Workbook::new();
        let s = &mut book.sheets[0].sheet;
        s.set_input(at("A1"), "x");
        let mut tools = s.data_tools().clone();
        tools.auto_filter = Some(AutoFilter::new(Range::new(at("A1"), at("A3"))));
        tools.cond_formats =
            vec![rule_from_spec(range("A1:A3"), "text", &["containsText".into(), "a\"b".into()], "lightRed").unwrap()];
        s.set_data_tools(tools);
        let xml = super::super::worksheet_xml(s, &mut StyleTable::default(), &mut Default::default());
        let sheet_data_end = xml.find("</sheetData>").unwrap();
        let af = xml.find(r#"<autoFilter ref="A1:A3"/>"#).unwrap();
        let cf = xml.find(r#"<conditionalFormatting sqref="A1:A3">"#).unwrap();
        assert!(sheet_data_end < af && af < cf);
        assert!(xml.contains(r#"operator="containsText" text="a&quot;b"><formula>NOT(ISERROR(SEARCH("a""b",A1)))</formula>"#)
            || xml.contains(r#"operator="containsText" text="a&quot;b"><formula>NOT(ISERROR(SEARCH(&quot;a&quot;&quot;b&quot;,A1)))</formula>"#));
        let mut styles = StyleTable::default();
        styles.dxf_index(&CfFormat::preset("lightRedDarkRed").unwrap());
        let dxfs = styles.dxfs_xml();
        assert_eq!(
            dxfs,
            r#"<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>"#
        );
        assert_eq!(
            read_dxfs(&format!("<styleSheet>{dxfs}</styleSheet>")),
            vec![CfFormat::preset("lightRedDarkRed").unwrap()]
        );
    }
}
