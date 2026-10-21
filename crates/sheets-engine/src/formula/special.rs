//! Functions that need more than argument values: the formula's own cell
//! (`ROW()`, `COLUMN()`), references themselves rather than their values
//! (`ROWS`, `COLUMNS`, `OFFSET`, `INDIRECT`), or facts about the cells they
//! read (`SUBTOTAL` and `AGGREGATE` skip nested subtotals and hidden rows).
//!
//! They are still registered in `functions::FUNCTIONS` (for arity checks,
//! autocomplete and docs); the evaluator calls [`call`] first and only falls
//! back to the registry implementation when it returns `None`.

use super::ast::Expr;
use super::eval::{eval_arg, eval_in, eval_ref, ref_to_value, Ctx, Reference};
use super::parser::parse;
use crate::cell::{CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::functions::{self, Arg, RangeValues};
use crate::value::{ErrorKind, Value};

/// Functions that return a reference (usable inside `SUM(...)` like a range).
pub(crate) fn is_reference_function(name: &str) -> bool {
    matches!(name, "OFFSET" | "INDIRECT")
}

/// Evaluate a special function, or `None` if `name` is an ordinary one.
pub(crate) fn call(name: &str, args: &[Expr], ctx: &Ctx) -> Option<Result<Value, ErrorKind>> {
    Some(match name {
        "ROW" | "COLUMN" => row_or_column(name == "ROW", args, ctx),
        "ROWS" | "COLUMNS" => rows_or_columns(name == "ROWS", args, ctx),
        "OFFSET" | "INDIRECT" => match reference_call(name, args, ctx)? {
            Ok(r) => Ok(ref_to_value(&r, ctx)),
            Err(e) => Err(e),
        },
        "SUBTOTAL" => subtotal(args, ctx),
        "AGGREGATE" => aggregate(args, ctx),
        _ => return None,
    })
}

/// `OFFSET`/`INDIRECT` as a reference. Every reference produced is recorded
/// in `ctx.dynamic` so the cell is recalculated when that range changes.
pub(crate) fn reference_call(name: &str, args: &[Expr], ctx: &Ctx) -> Option<Result<Reference, ErrorKind>> {
    let result = match name {
        "OFFSET" => offset(args, ctx),
        "INDIRECT" => indirect(args, ctx),
        _ => return None,
    };
    if let Ok(r) = &result {
        ctx.dynamic.borrow_mut().push(r.clone());
    }
    Some(result)
}

fn number(e: &Expr, ctx: &Ctx) -> Result<f64, ErrorKind> {
    eval_in(e, ctx).as_number()
}

fn row_or_column(row: bool, args: &[Expr], ctx: &Ctx) -> Result<Value, ErrorKind> {
    let at = match args.first() {
        None => ctx.at.ok_or(ErrorKind::Value)?,
        Some(e) => match eval_ref(e, ctx) {
            Some(r) => r?.range.start,
            None => return Err(ErrorKind::Value),
        },
    };
    Ok(Value::Number(if row { at.row } else { at.col } as f64 + 1.0))
}

fn rows_or_columns(rows: bool, args: &[Expr], ctx: &Ctx) -> Result<Value, ErrorKind> {
    let (r, c) = match eval_ref(&args[0], ctx) {
        Some(r) => {
            let range = r?.range;
            (range.rows(), range.cols())
        }
        None => match eval_arg(&args[0], ctx) {
            Arg::Range(v) => (v.rows, v.cols),
            Arg::Value(Value::Error(e)) => return Err(e),
            Arg::Value(_) => (1, 1),
        },
    };
    Ok(Value::Number(if rows { r } else { c } as f64))
}

/// `OFFSET(reference, rows, cols, [height], [width])`.
fn offset(args: &[Expr], ctx: &Ctx) -> Result<Reference, ErrorKind> {
    if !(3..=5).contains(&args.len()) {
        return Err(ErrorKind::Value);
    }
    let base = match eval_ref(&args[0], ctx) {
        Some(r) => r?,
        None => return Err(ErrorKind::Value),
    };
    let d_rows = number(&args[1], ctx)?.trunc() as i64;
    let d_cols = number(&args[2], ctx)?.trunc() as i64;
    let size = |i: usize, default: u32| -> Result<i64, ErrorKind> {
        match args.get(i) {
            None => Ok(default as i64),
            Some(e) => Ok(number(e, ctx)?.trunc() as i64),
        }
    };
    let height = size(3, base.range.rows())?;
    let width = size(4, base.range.cols())?;
    if height <= 0 || width <= 0 {
        return Err(ErrorKind::Ref);
    }
    let top = base.range.start.row as i64 + d_rows;
    let left = base.range.start.col as i64 + d_cols;
    let (bottom, right) = (top + height - 1, left + width - 1);
    if top < 0 || left < 0 || bottom >= MAX_ROWS as i64 || right >= MAX_COLS as i64 {
        return Err(ErrorKind::Ref);
    }
    let range = Range::new(CellRef::new(top as u32, left as u32), CellRef::new(bottom as u32, right as u32));
    Ok(Reference { sheet: base.sheet, range })
}

/// `INDIRECT(ref_text, [a1])`: "B3", "Sheet2!A1:B5", "'My Sheet'!C:C", or R1C1 ("R2C3", "R[-1]C").
fn indirect(args: &[Expr], ctx: &Ctx) -> Result<Reference, ErrorKind> {
    if args.is_empty() || args.len() > 2 {
        return Err(ErrorKind::Value);
    }
    let text = eval_in(&args[0], ctx).as_text()?;
    let a1 = match args.get(1) {
        Some(e) => eval_in(e, ctx).as_bool()?,
        None => true,
    };
    let text = text.trim();
    if a1 {
        return match parse(text) {
            Ok(Expr::Ref(c)) => Ok(Reference { sheet: None, range: Range::single(c) }),
            Ok(Expr::Range(r)) => Ok(Reference { sheet: None, range: r }),
            Ok(Expr::SheetRef(s, c)) => Ok(Reference { sheet: Some(s), range: Range::single(c) }),
            Ok(Expr::SheetRange(s, r)) => Ok(Reference { sheet: Some(s), range: r }),
            _ => Err(ErrorKind::Ref),
        };
    }
    // R1C1: optional sheet prefix, then one or two R..C.. parts.
    let (sheet, body) = match text.rfind('!') {
        Some(i) => {
            let raw = &text[..i];
            let name = raw
                .strip_prefix('\'')
                .and_then(|s| s.strip_suffix('\''))
                .map_or_else(|| raw.to_string(), |s| s.replace("''", "'"));
            (Some(name), &text[i + 1..])
        }
        None => (None, text),
    };
    let origin = ctx.at.unwrap_or(CellRef::new(0, 0));
    let mut parts = body.split(':');
    let a = parse_r1c1(parts.next().unwrap_or(""), origin).ok_or(ErrorKind::Ref)?;
    let b = match parts.next() {
        Some(p) => parse_r1c1(p, origin).ok_or(ErrorKind::Ref)?,
        None => a,
    };
    if parts.next().is_some() {
        return Err(ErrorKind::Ref);
    }
    Ok(Reference { sheet, range: Range::new(a, b) })
}

/// One R1C1 address relative to `origin`: `R2C3`, `R[-1]C[2]`, `RC`.
fn parse_r1c1(s: &str, origin: CellRef) -> Option<CellRef> {
    let upper = s.trim().to_ascii_uppercase();
    let rest = upper.strip_prefix('R')?;
    let (row, rest) = r1c1_part(rest, origin.row, MAX_ROWS)?;
    let rest = rest.strip_prefix('C')?;
    let (col, rest) = r1c1_part(rest, origin.col, MAX_COLS)?;
    rest.is_empty().then_some(CellRef::new(row, col))
}

/// `5` (absolute, 1-based), `[-2]` (relative) or nothing (same as origin).
fn r1c1_part(s: &str, origin: u32, limit: u32) -> Option<(u32, &str)> {
    if let Some(rest) = s.strip_prefix('[') {
        let end = rest.find(']')?;
        let d: i64 = rest[..end].parse().ok()?;
        let n = origin as i64 + d;
        return (0..limit as i64).contains(&n).then_some((n as u32, &rest[end + 1..]));
    }
    let digits = s.len() - s.trim_start_matches(|c: char| c.is_ascii_digit()).len();
    if digits == 0 {
        return Some((origin, s));
    }
    let n: u32 = s[..digits].parse().ok()?;
    (1..=limit).contains(&n).then_some((n - 1, &s[digits..]))
}

/// Values of reference arguments for `SUBTOTAL`/`AGGREGATE`, skipping cells
/// holding other subtotals and, if asked, hidden rows.
fn collect(r: &Reference, ctx: &Ctx, skip_hidden: bool, skip_subtotals: bool, out: &mut Vec<Value>) {
    let range = ctx.clip(r.range);
    let sheet = r.sheet.as_deref();
    if let Some(s) = sheet {
        if !ctx.src.has_sheet(s) {
            out.push(Value::Error(ErrorKind::Ref));
            return;
        }
    }
    for row in range.start.row..=range.end.row {
        if skip_hidden && ctx.src.is_row_hidden(sheet, row) {
            continue;
        }
        for col in range.start.col..=range.end.col {
            let at = CellRef::new(row, col);
            if skip_subtotals && ctx.src.is_subtotal_cell(sheet, at) {
                continue;
            }
            out.push(ctx.value(sheet, at));
        }
    }
}

/// Run a registered function on a flat list of values, as if they came from one range.
fn run(name: &str, values: Vec<Value>, extra: Option<Value>) -> Result<Value, ErrorKind> {
    let f = functions::lookup(name).expect("registered");
    let n = values.len() as u32;
    let mut args = vec![Arg::Range(RangeValues { rows: 1, cols: n.max(1), values })];
    if n == 0 {
        args[0] = Arg::Range(RangeValues { rows: 0, cols: 0, values: Vec::new() });
    }
    if let Some(v) = extra {
        args.push(Arg::Value(v));
    }
    (f.call)(&args)
}

const SUBTOTAL_FUNCTIONS: [&str; 11] =
    ["AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "PRODUCT", "STDEV.S", "STDEV.P", "SUM", "VAR.S", "VAR.P"];

/// `SUBTOTAL(function_num, ref1, [ref2], ...)`: 1-11 include hidden rows, 101-111 skip them.
fn subtotal(args: &[Expr], ctx: &Ctx) -> Result<Value, ErrorKind> {
    let code = number(&args[0], ctx)?.trunc() as i64;
    let (index, skip_hidden) = match code {
        1..=11 => (code - 1, false),
        101..=111 => (code - 101, true),
        _ => return Err(ErrorKind::Value),
    };
    let mut values = Vec::new();
    for e in &args[1..] {
        match eval_ref(e, ctx) {
            Some(r) => collect(&r?, ctx, skip_hidden, true, &mut values),
            None => return Err(ErrorKind::Value),
        }
    }
    run(SUBTOTAL_FUNCTIONS[index as usize], values, None)
}

const AGGREGATE_FUNCTIONS: [&str; 19] = [
    "AVERAGE",
    "COUNT",
    "COUNTA",
    "MAX",
    "MIN",
    "PRODUCT",
    "STDEV.S",
    "STDEV.P",
    "SUM",
    "VAR.S",
    "VAR.P",
    "MEDIAN",
    "MODE.SNGL",
    "LARGE",
    "SMALL",
    "PERCENTILE.INC",
    "QUARTILE.INC",
    "PERCENTILE.EXC",
    "QUARTILE.EXC",
];

/// `AGGREGATE(function_num, options, ref1, [ref2], ...)` or, for 14-19,
/// `AGGREGATE(function_num, options, array, k)`.
fn aggregate(args: &[Expr], ctx: &Ctx) -> Result<Value, ErrorKind> {
    let code = number(&args[0], ctx)?.trunc() as i64;
    let options = number(&args[1], ctx)?.trunc() as i64;
    if !(1..=19).contains(&code) || !(0..=7).contains(&options) {
        return Err(ErrorKind::Value);
    }
    let skip_subtotals = options <= 3;
    let skip_hidden = matches!(options, 1 | 3 | 5 | 7);
    let skip_errors = matches!(options, 2 | 3 | 6 | 7);
    let with_k = code >= 14;
    let data_args = if with_k {
        if args.len() != 4 {
            return Err(ErrorKind::Value);
        }
        &args[2..3]
    } else {
        &args[2..]
    };
    if data_args.is_empty() {
        return Err(ErrorKind::Value);
    }
    let mut values = Vec::new();
    for e in data_args {
        match eval_ref(e, ctx) {
            Some(r) => collect(&r?, ctx, skip_hidden, skip_subtotals, &mut values),
            None => match eval_arg(e, ctx) {
                Arg::Range(r) => values.extend(r.values),
                Arg::Value(v) => values.push(v),
            },
        }
    }
    if skip_errors {
        values.retain(|v| !matches!(v, Value::Error(_)));
    }
    let k = if with_k { Some(eval_in(&args[3], ctx)) } else { None };
    run(AGGREGATE_FUNCTIONS[code as usize - 1], values, k)
}

#[cfg(test)]
mod tests {
    use std::collections::{HashMap, HashSet};

    use crate::cell::CellRef;
    use crate::formula::{eval_at, parse, CellSource};
    use crate::value::{ErrorKind, Value};

    /// A1:A5 = 10,20,30,40,50 (A3 is a SUBTOTAL), B1 = "B2", row 4 hidden;
    /// sheet "Other" has A1 = 7.
    struct Src {
        cells: HashMap<CellRef, Value>,
        hidden: HashSet<u32>,
    }

    impl CellSource for Src {
        fn value(&self, at: CellRef) -> Value {
            self.cells.get(&at).cloned().unwrap_or_default()
        }

        fn has_sheet(&self, sheet: &str) -> bool {
            sheet.eq_ignore_ascii_case("other")
        }

        fn sheet_value(&self, _sheet: &str, at: CellRef) -> Value {
            if at == CellRef::new(0, 0) {
                Value::Number(7.0)
            } else {
                Value::Empty
            }
        }

        fn used_extent(&self) -> (u32, u32) {
            (6, 3)
        }

        fn is_subtotal_cell(&self, sheet: Option<&str>, at: CellRef) -> bool {
            sheet.is_none() && at == CellRef::new(2, 0)
        }

        fn is_row_hidden(&self, sheet: Option<&str>, row: u32) -> bool {
            sheet.is_none() && self.hidden.contains(&row)
        }
    }

    fn run(formula: &str) -> Value {
        let mut cells: HashMap<CellRef, Value> =
            (0..5).map(|r| (CellRef::new(r, 0), Value::Number((r as f64 + 1.0) * 10.0))).collect();
        cells.insert(CellRef::new(0, 1), Value::Text("B2".into()));
        cells.insert(CellRef::new(1, 1), Value::Number(99.0));
        cells.insert(CellRef::new(4, 2), Value::Error(ErrorKind::NA));
        let src = Src { cells, hidden: HashSet::from([3]) };
        eval_at(&parse(formula).unwrap(), &src, CellRef::new(6, 2)).0
    }

    fn n(x: f64) -> Value {
        Value::Number(x)
    }

    #[test]
    fn row_and_column() {
        assert_eq!(run("ROW()"), n(7.0));
        assert_eq!(run("COLUMN()"), n(3.0));
        assert_eq!(run("ROW(B5)"), n(5.0));
        assert_eq!(run("COLUMN(D2:F9)"), n(4.0));
        assert_eq!(run("ROW(3:9)"), n(3.0));
        assert_eq!(run("COLUMN(Other!C1)"), n(3.0));
        assert_eq!(run("ROW(5)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("ROWS(A1:C10)"), n(10.0));
        assert_eq!(run("COLUMNS(A1:C10)"), n(3.0));
        assert_eq!(run("ROWS(A:A)"), n(1_048_576.0));
        assert_eq!(run("COLUMNS(1:1)"), n(16_384.0));
        assert_eq!(run("ROWS(5)"), n(1.0));
    }

    #[test]
    fn offset_and_indirect() {
        assert_eq!(run("OFFSET(A1,2,0)"), n(30.0));
        assert_eq!(run("SUM(OFFSET(A1,1,0,3,1))"), n(90.0));
        assert_eq!(run("SUM(OFFSET(A1:A2,3,0))"), n(90.0));
        assert_eq!(run("OFFSET(A1,-1,0)"), Value::Error(ErrorKind::Ref));
        assert_eq!(run("OFFSET(A1,0,0,0,1)"), Value::Error(ErrorKind::Ref));
        assert_eq!(run("OFFSET(5,0,0)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("INDIRECT(\"A2\")"), n(20.0));
        assert_eq!(run("INDIRECT(B1)"), n(99.0));
        assert_eq!(run("SUM(INDIRECT(\"A1:A3\"))"), n(60.0));
        assert_eq!(run("INDIRECT(\"Other!A1\")"), n(7.0));
        assert_eq!(run("INDIRECT(\"'other'!A1\")+1"), n(8.0));
        assert_eq!(run("INDIRECT(\"R2C1\",FALSE)"), n(20.0));
        assert_eq!(run("INDIRECT(\"R[-6]C[-2]\",FALSE)"), n(10.0));
        assert_eq!(run("SUM(INDIRECT(\"R1C1:R2C1\",FALSE))"), n(30.0));
        assert_eq!(run("INDIRECT(\"nope\")"), Value::Error(ErrorKind::Ref));
        assert_eq!(run("INDIRECT(\"Missing!A1\")"), Value::Error(ErrorKind::Ref));
    }

    #[test]
    fn records_dynamic_references() {
        let src = Src { cells: HashMap::new(), hidden: HashSet::new() };
        let (_, refs) = eval_at(&parse("INDIRECT(\"Other!B2\")").unwrap(), &src, CellRef::new(0, 0));
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].sheet.as_deref(), Some("Other"));
    }

    #[test]
    fn subtotal_skips_subtotals_and_hidden_rows() {
        // A3 (30) is a nested subtotal and always skipped; row 4 (40) is hidden.
        assert_eq!(run("SUBTOTAL(9,A1:A5)"), n(120.0));
        assert_eq!(run("SUBTOTAL(109,A1:A5)"), n(80.0));
        assert_eq!(run("SUBTOTAL(1,A1:A5)"), n(30.0));
        assert_eq!(run("SUBTOTAL(2,A:A)"), n(4.0));
        assert_eq!(run("SUBTOTAL(3,A1:B2)"), n(4.0));
        assert_eq!(run("SUBTOTAL(4,A1:A5)"), n(50.0));
        assert_eq!(run("SUBTOTAL(105,A1:A5)"), n(10.0));
        assert_eq!(run("SUBTOTAL(9,A1:A2,A5)"), n(80.0));
        assert_eq!(run("SUBTOTAL(12,A1:A5)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("SUBTOTAL(9,5)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("SUBTOTAL(9,C1:C5)"), Value::Error(ErrorKind::NA));
    }

    #[test]
    fn aggregate_options() {
        assert_eq!(run("AGGREGATE(9,4,A1:A5)"), n(150.0));
        assert_eq!(run("AGGREGATE(9,0,A1:A5)"), n(120.0));
        assert_eq!(run("AGGREGATE(9,5,A1:A5)"), n(110.0));
        assert_eq!(run("AGGREGATE(9,3,A1:A5)"), n(80.0));
        assert_eq!(run("AGGREGATE(9,4,A1:C5)"), Value::Error(ErrorKind::NA));
        assert_eq!(run("AGGREGATE(9,6,A1:C5)"), n(249.0));
        assert_eq!(run("AGGREGATE(14,4,A1:A5,2)"), n(40.0));
        assert_eq!(run("AGGREGATE(15,6,A1:C5,1)"), n(10.0));
        assert_eq!(run("AGGREGATE(12,4,A1:A5)"), n(30.0));
        assert_eq!(run("AGGREGATE(16,4,A1:A5,0.5)"), n(30.0));
        assert_eq!(run("AGGREGATE(17,4,A1:A5,1)"), n(20.0));
        assert_eq!(run("AGGREGATE(14,4,A1:A5)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("AGGREGATE(20,4,A1:A5)"), Value::Error(ErrorKind::Value));
    }
}
