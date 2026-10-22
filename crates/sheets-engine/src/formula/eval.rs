use std::cell::RefCell;
use std::cmp::Ordering;

use super::arrays;
use super::ast::{BinOp, Expr, UnaryOp};
use super::special;
use crate::cell::{CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::functions::{self, Arg, RangeValues};
use crate::value::{compare, ErrorKind, Value};

/// Where the evaluator reads cell values from. `Sheet` implements this for a
/// lone sheet, the workbook calc layer for a sheet inside a workbook (so
/// `Sheet2!A1` works); tests can use a plain `HashMap`.
///
/// Only `value` is required. The other methods have defaults that describe a
/// source with a single sheet and no extra information.
pub trait CellSource {
    /// A cell on the formula's own sheet.
    fn value(&self, at: CellRef) -> Value;

    fn range(&self, r: Range) -> RangeValues {
        RangeValues { rows: r.rows(), cols: r.cols(), values: r.cells().map(|c| self.value(c)).collect() }
    }

    /// True if a sheet with this name (any case) exists.
    fn has_sheet(&self, _sheet: &str) -> bool {
        false
    }

    /// A cell on the named sheet. Only called when `has_sheet` is true.
    fn sheet_value(&self, _sheet: &str, _at: CellRef) -> Value {
        Value::Error(ErrorKind::Ref)
    }

    /// (rows, cols) beyond which every sheet is empty. Ranges reaching the
    /// last row or column (`A:A`, `A5:A1048576`, `1:1`) are cut off there, so a
    /// whole-column reference costs as much as the data, not a million cells.
    fn used_extent(&self) -> (u32, u32) {
        (MAX_ROWS, MAX_COLS)
    }

    /// True if the cell holds a `SUBTOTAL` or `AGGREGATE` formula (which
    /// `SUBTOTAL` skips to avoid double counting). `sheet` is `None` for the own sheet.
    fn is_subtotal_cell(&self, _sheet: Option<&str>, _at: CellRef) -> bool {
        false
    }

    /// True if the row is hidden (`SUBTOTAL(101..111)` skips hidden rows).
    fn is_row_hidden(&self, _sheet: Option<&str>, _row: u32) -> bool {
        false
    }
}

/// A resolved reference: a range on the formula's own sheet (`sheet: None`) or on a named sheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    pub sheet: Option<String>,
    pub range: Range,
}

/// Evaluation state for one formula.
pub(crate) struct Ctx<'a> {
    pub src: &'a dyn CellSource,
    /// The cell holding the formula, for `ROW()`, `COLUMN()` and relative `INDIRECT`.
    pub at: Option<CellRef>,
    /// References produced by `OFFSET`/`INDIRECT`, so the caller can track them as dependencies.
    pub dynamic: RefCell<Vec<Reference>>,
    /// Names bound by `LET`, innermost last (upper case).
    pub lets: RefCell<Vec<(String, Arg)>>,
}

impl Ctx<'_> {
    /// Values of a reference; ranges reaching the sheet edge are clipped to the used area.
    pub fn values(&self, r: &Reference) -> RangeValues {
        let range = self.clip(r.range);
        match &r.sheet {
            None => self.src.range(range),
            Some(s) if self.src.has_sheet(s) => RangeValues {
                rows: range.rows(),
                cols: range.cols(),
                values: range.cells().map(|c| self.src.sheet_value(s, c)).collect(),
            },
            Some(_) => RangeValues { rows: 1, cols: 1, values: vec![Value::Error(ErrorKind::Ref)] },
        }
    }

    /// The value of a name: a `LET` variable.
    pub fn name(&self, name: &str) -> Option<Arg> {
        let upper = name.to_ascii_uppercase();
        self.lets.borrow().iter().rev().find(|(n, _)| *n == upper).map(|(_, v)| v.clone())
    }

    pub fn value(&self, sheet: Option<&str>, at: CellRef) -> Value {
        match sheet {
            None => self.src.value(at),
            Some(s) if self.src.has_sheet(s) => self.src.sheet_value(s, at),
            Some(_) => Value::Error(ErrorKind::Ref),
        }
    }

    /// Cut a range that reaches the last row/column down to the used area.
    pub fn clip(&self, r: Range) -> Range {
        let (rows, cols) = self.src.used_extent();
        let mut end = r.end;
        if end.row == MAX_ROWS - 1 {
            end.row = rows.saturating_sub(1).clamp(r.start.row, MAX_ROWS - 1);
        }
        if end.col == MAX_COLS - 1 {
            end.col = cols.saturating_sub(1).clamp(r.start.col, MAX_COLS - 1);
        }
        Range { start: r.start, end }
    }
}

/// Evaluate to a single value, with no current cell (`ROW()` without arguments is `#VALUE!`).
pub fn eval(expr: &Expr, src: &dyn CellSource) -> Value {
    let ctx = Ctx { src, at: None, dynamic: RefCell::new(Vec::new()), lets: RefCell::new(Vec::new()) };
    eval_in(expr, &ctx)
}

/// Evaluate the formula in cell `at`. Also returns the references that
/// `OFFSET`/`INDIRECT` resolved to, which the caller should treat as extra
/// precedents of the cell.
pub fn eval_at(expr: &Expr, src: &dyn CellSource, at: CellRef) -> (Value, Vec<Reference>) {
    let ctx = Ctx { src, at: Some(at), dynamic: RefCell::new(Vec::new()), lets: RefCell::new(Vec::new()) };
    let value = eval_in(expr, &ctx);
    (value, ctx.dynamic.into_inner())
}

/// Evaluate the formula in cell `at` as Excel 365 does: the result may be an
/// array, which the sheet spills into the cells around `at`.
pub fn eval_at_any(expr: &Expr, src: &dyn CellSource, at: CellRef) -> (Arg, Vec<Reference>) {
    let ctx = Ctx { src, at: Some(at), dynamic: RefCell::new(Vec::new()), lets: RefCell::new(Vec::new()) };
    let value = arrays::eval_any(expr, &ctx);
    (value, ctx.dynamic.into_inner())
}

pub(crate) fn eval_in(expr: &Expr, ctx: &Ctx) -> Value {
    match expr {
        Expr::Number(n) => Value::Number(*n),
        Expr::Text(s) => Value::Text(s.clone()),
        Expr::Bool(b) => Value::Bool(*b),
        Expr::Ref(c) => ctx.src.value(*c),
        Expr::SheetRef(s, c) => ctx.value(Some(s), *c),
        Expr::Range(_) | Expr::SheetRange(..) => match eval_ref(expr, ctx) {
            Some(Ok(r)) => ref_to_value(&r, ctx),
            Some(Err(e)) => Value::Error(e),
            None => Value::Error(ErrorKind::Value),
        },
        Expr::Name(n) => match ctx.name(n) {
            Some(Arg::Value(v)) => v,
            Some(Arg::Range(r)) => r.values.into_iter().next().unwrap_or(Value::Empty),
            None => Value::Error(ErrorKind::Name),
        },
        Expr::Error(e) => Value::Error(*e),
        Expr::Missing => Value::Empty,
        Expr::Unary(op, e) => unary_value(*op, eval_in(e, ctx)),
        Expr::Binary(op, a, b) => binary(*op, eval_in(a, ctx), eval_in(b, ctx)),
        // An array function where one value is expected gives its first value.
        Expr::Call(name, args) if arrays::is_array_function(name) => match arrays::call(name, args, ctx) {
            Ok(Arg::Value(v)) => v,
            Ok(Arg::Range(r)) => r.values.into_iter().next().unwrap_or(Value::Empty),
            Err(e) => Value::Error(e),
        },
        Expr::Call(name, args) => call(name, args, ctx),
    }
}

/// A reference used where one value is expected: a single cell gives its
/// value; a one-column (or one-row) range gives the cell in the formula's row
/// (column), as Excel's implicit intersection does; anything else is `#VALUE!`.
pub(crate) fn ref_to_value(r: &Reference, ctx: &Ctx) -> Value {
    let range = r.range;
    let sheet = r.sheet.as_deref();
    if range.start == range.end {
        return ctx.value(sheet, range.start);
    }
    if let Some(at) = ctx.at {
        if range.start.col == range.end.col && (range.start.row..=range.end.row).contains(&at.row) {
            return ctx.value(sheet, CellRef::new(at.row, range.start.col));
        }
        if range.start.row == range.end.row && (range.start.col..=range.end.col).contains(&at.col) {
            return ctx.value(sheet, CellRef::new(range.start.row, at.col));
        }
    }
    Value::Error(ErrorKind::Value)
}

/// The reference an expression stands for, if it is reference-valued:
/// `A1`, `A1:B2`, `Sheet2!A1`, `OFFSET(...)`, `INDIRECT(...)`.
pub(crate) fn eval_ref(expr: &Expr, ctx: &Ctx) -> Option<Result<Reference, ErrorKind>> {
    let local = |range| Some(Ok(Reference { sheet: None, range }));
    match expr {
        Expr::Ref(c) => local(Range::single(*c)),
        Expr::Range(r) => local(*r),
        Expr::SheetRef(s, c) => Some(Ok(Reference { sheet: Some(s.clone()), range: Range::single(*c) })),
        Expr::SheetRange(s, r) => Some(Ok(Reference { sheet: Some(s.clone()), range: *r })),
        Expr::Call(name, args) => special::reference_call(name, args, ctx),
        _ => None,
    }
}

/// Evaluate a function argument: ranges stay ranges so `SUM(A1:A9)` sees every cell.
pub(crate) fn eval_arg(expr: &Expr, ctx: &Ctx) -> Arg {
    match expr {
        Expr::Range(_) | Expr::SheetRange(..) => match eval_ref(expr, ctx) {
            Some(Ok(r)) => Arg::Range(ctx.values(&r)),
            Some(Err(e)) => Arg::Value(Value::Error(e)),
            None => Arg::Value(Value::Error(ErrorKind::Value)),
        },
        Expr::Call(name, args) if special::is_reference_function(name) => match eval_ref(expr, ctx) {
            Some(Ok(r)) => Arg::Range(ctx.values(&r)),
            Some(Err(e)) => Arg::Value(Value::Error(e)),
            None => Arg::Value(call(name, args, ctx)),
        },
        // Operators, names and array functions may give arrays: `SUM(A1:A9*B1:B9)`, `SUM(FILTER(...))`.
        Expr::Binary(..) | Expr::Unary(..) | Expr::Name(_) => arrays::eval_any(expr, ctx),
        Expr::Call(name, _) if arrays::is_array_function(name) => arrays::eval_any(expr, ctx),
        other => Arg::Value(eval_in(other, ctx)),
    }
}

pub(crate) fn unary_value(op: UnaryOp, v: Value) -> Value {
    match (op, v.as_number()) {
        (_, Err(e)) => Value::Error(e),
        (UnaryOp::Neg, Ok(n)) => Value::Number(-n),
        (UnaryOp::Percent, Ok(n)) => Value::Number(n / 100.0),
    }
}

pub(crate) fn binary_values(op: BinOp, a: Value, b: Value) -> Value {
    binary(op, a, b)
}

fn call(name: &str, args: &[Expr], ctx: &Ctx) -> Value {
    let Some(f) = functions::lookup(name) else {
        return Value::Error(ErrorKind::Name);
    };
    if args.len() < f.min_args || f.max_args.is_some_and(|max| args.len() > max) {
        return Value::Error(ErrorKind::Value);
    }
    let result = match special::call(name, args, ctx) {
        Some(result) => result,
        None => {
            let args: Vec<Arg> = args.iter().map(|a| eval_arg(a, ctx)).collect();
            (f.call)(&args)
        }
    };
    match result {
        Ok(Value::Number(n)) if !n.is_finite() => Value::Error(ErrorKind::Num),
        Ok(v) => v,
        Err(e) => Value::Error(e),
    }
}

fn binary(op: BinOp, a: Value, b: Value) -> Value {
    if let Value::Error(e) = a {
        return Value::Error(e);
    }
    if let Value::Error(e) = b {
        return Value::Error(e);
    }
    let result = match op {
        BinOp::Concat => a.as_text().and_then(|x| Ok(Value::Text(x + &b.as_text()?))),
        BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge => {
            let ord = compare(&a, &b);
            Ok(Value::Bool(match op {
                BinOp::Eq => ord == Ordering::Equal,
                BinOp::Ne => ord != Ordering::Equal,
                BinOp::Lt => ord == Ordering::Less,
                BinOp::Le => ord != Ordering::Greater,
                BinOp::Gt => ord == Ordering::Greater,
                _ => ord != Ordering::Less,
            }))
        }
        _ => arithmetic(op, &a, &b),
    };
    match result {
        Ok(Value::Number(n)) if !n.is_finite() => Value::Error(ErrorKind::Num),
        Ok(v) => v,
        Err(e) => Value::Error(e),
    }
}

fn arithmetic(op: BinOp, a: &Value, b: &Value) -> Result<Value, ErrorKind> {
    let (x, y) = (a.as_number()?, b.as_number()?);
    Ok(Value::Number(match op {
        BinOp::Add => x + y,
        BinOp::Sub => x - y,
        BinOp::Mul => x * y,
        BinOp::Div if y == 0.0 => return Err(ErrorKind::Div0),
        BinOp::Div => x / y,
        BinOp::Pow => x.powf(y),
        _ => unreachable!("not an arithmetic operator"),
    }))
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use super::*;
    use crate::formula::parse;

    struct Cells(HashMap<CellRef, Value>);

    impl CellSource for Cells {
        fn value(&self, at: CellRef) -> Value {
            self.0.get(&at).cloned().unwrap_or_default()
        }

        fn used_extent(&self) -> (u32, u32) {
            (3, 3)
        }
    }

    fn cells() -> Cells {
        Cells(HashMap::from([
            (CellRef::new(0, 0), Value::Number(2.0)),
            (CellRef::new(1, 0), Value::Text("x".into())),
            (CellRef::new(2, 0), Value::Number(5.0)),
        ]))
    }

    fn run(src: &str) -> Value {
        eval(&parse(src).unwrap(), &cells())
    }

    fn run_at(src: &str, at: &str) -> Value {
        eval_at(&parse(src).unwrap(), &cells(), CellRef::parse(at).unwrap()).0
    }

    #[test]
    fn evaluates_operators() {
        assert_eq!(run("1+2*3"), Value::Number(7.0));
        assert_eq!(run("-2^2"), Value::Number(4.0));
        assert_eq!(run("2^3^2"), Value::Number(64.0));
        assert_eq!(run("50%"), Value::Number(0.5));
        assert_eq!(run("A1*10"), Value::Number(20.0));
        assert_eq!(run("\"n=\"&A1"), Value::Text("n=2".into()));
        assert_eq!(run("A1>1"), Value::Bool(true));
        assert_eq!(run("\"abc\"=\"ABC\""), Value::Bool(true));
    }

    #[test]
    fn produces_errors() {
        assert_eq!(run("1/0"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("A2+1"), Value::Error(ErrorKind::Value));
        assert_eq!(run("NOPE(1)"), Value::Error(ErrorKind::Name));
        assert_eq!(run("foo"), Value::Error(ErrorKind::Name));
        assert_eq!(run("(1/0)&\"x\""), Value::Error(ErrorKind::Div0));
        assert_eq!(run("Nowhere!A1"), Value::Error(ErrorKind::Ref));
        assert_eq!(run("SUM(Nowhere!A1:A3)"), Value::Error(ErrorKind::Ref));
    }

    #[test]
    fn whole_columns_are_clipped_and_intersected() {
        assert_eq!(run("SUM(A:A)"), Value::Number(7.0));
        assert_eq!(run("SUM(1:1)"), Value::Number(2.0));
        assert_eq!(run("COUNTA(A:C)"), Value::Number(3.0));
        // Implicit intersection: a column used as one value gives the cell in the formula's row.
        assert_eq!(run_at("A:A*10", "C3"), Value::Number(50.0));
        assert_eq!(run_at("A1:A2", "D9"), Value::Error(ErrorKind::Value));
    }
}
