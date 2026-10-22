//! Dynamic arrays, as in Excel 365: formulas that return several values
//! "spill" into the cells below and to the right (see `Sheet::store_result`).
//!
//! `eval_any` evaluates an expression to a value or an array: ranges stay
//! arrays, operators work element by element (`A2:A10*B2:B10`,
//! `B2:B10>100`), many one-value functions apply to each element
//! (`UPPER(A2:A5)`), and the array functions here build new arrays: FILTER,
//! SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, TRANSPOSE, TAKE, DROP,
//! CHOOSEROWS, CHOOSECOLS, VSTACK, HSTACK, TOCOL, TOROW, EXPAND, TEXTSPLIT,
//! plus XMATCH and LET.

use std::cmp::Ordering;
use std::collections::HashSet;

use super::ast::Expr;
use super::eval::{binary_values, eval_ref, unary_value, Ctx};
use super::special;
use crate::functions::{self, Arg, RangeValues};
use crate::value::{compare, ErrorKind, Value};

/// Largest array a formula may produce (cells), to keep a typo like
/// `SEQUENCE(1e9)` from exhausting memory.
const MAX_CELLS: u64 = 5_000_000;

/// Functions that are applied to each element when given an array where they
/// expect one value (Excel's "lifting").
const LIFTED: &[&str] = &[
    "ABS",
    "CHAR",
    "CODE",
    "DATE",
    "DAY",
    "EXACT",
    "EXP",
    "HOUR",
    "IF",
    "IFERROR",
    "IFNA",
    "INT",
    "ISBLANK",
    "ISERR",
    "ISERROR",
    "ISEVEN",
    "ISLOGICAL",
    "ISNA",
    "ISNONTEXT",
    "ISNUMBER",
    "ISODD",
    "ISTEXT",
    "LEFT",
    "LEN",
    "LN",
    "LOG",
    "LOG10",
    "LOWER",
    "MID",
    "MINUTE",
    "MOD",
    "MONTH",
    "N",
    "NOT",
    "POWER",
    "PROPER",
    "REPT",
    "RIGHT",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "SECOND",
    "SIGN",
    "SQRT",
    "T",
    "TEXT",
    "TRIM",
    "TRUNC",
    "UPPER",
    "VALUE",
    "WEEKDAY",
    "YEAR",
];

/// The array functions (they need whole arrays as arguments and may return arrays).
pub const ARRAY_FUNCTIONS: &[&str] = &[
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "EXPAND",
    "FILTER",
    "HSTACK",
    "LET",
    "RANDARRAY",
    "SEQUENCE",
    "SORT",
    "SORTBY",
    "TAKE",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "TRANSPOSE",
    "UNIQUE",
    "VSTACK",
    "XMATCH",
];

pub fn is_array_function(name: &str) -> bool {
    ARRAY_FUNCTIONS.contains(&name)
}

fn arr(rows: u32, cols: u32, values: Vec<Value>) -> Arg {
    if rows == 1 && cols == 1 {
        return Arg::Value(values.into_iter().next().unwrap_or(Value::Empty));
    }
    Arg::Range(RangeValues { rows, cols, values })
}

/// Any argument as an array (a single value is 1x1).
pub(crate) fn as_array(a: Arg) -> RangeValues {
    match a {
        Arg::Range(r) => r,
        Arg::Value(v) => RangeValues { rows: 1, cols: 1, values: vec![v] },
    }
}

fn too_big(rows: u64, cols: u64) -> bool {
    rows * cols > MAX_CELLS
}

/// Evaluate to a value or an array.
pub(crate) fn eval_any(expr: &Expr, ctx: &Ctx) -> Arg {
    match expr {
        Expr::Range(_) | Expr::SheetRange(..) => match eval_ref(expr, ctx) {
            Some(Ok(r)) => normalize(Arg::Range(ctx.values(&r))),
            Some(Err(e)) => Arg::Value(Value::Error(e)),
            None => Arg::Value(Value::Error(ErrorKind::Value)),
        },
        Expr::Name(n) => match ctx.name(n) {
            Some(a) => a,
            // A defined name: its cells as an array, or its formula's result.
            None => match ctx.defined(n) {
                Some(e) => ctx.deeper(|| eval_any(&e, ctx)),
                None => Arg::Value(Value::Error(ErrorKind::Name)),
            },
        },
        Expr::Unary(op, e) => match eval_any(e, ctx) {
            Arg::Value(v) => Arg::Value(unary_value(*op, v)),
            Arg::Range(r) => map(r, |v| unary_value(*op, v)),
        },
        Expr::Binary(op, a, b) => {
            let (x, y) = (eval_any(a, ctx), eval_any(b, ctx));
            match (x, y) {
                (Arg::Value(x), Arg::Value(y)) => Arg::Value(binary_values(*op, x, y)),
                (x, y) => broadcast2(as_array(x), as_array(y), |p, q| binary_values(*op, p, q)),
            }
        }
        Expr::Call(name, args) if is_array_function(name) => match call(name, args, ctx) {
            Ok(a) => normalize(a),
            Err(e) => Arg::Value(Value::Error(e)),
        },
        Expr::Call(name, args) if special::is_reference_function(name) => match eval_ref(expr, ctx) {
            Some(Ok(r)) => normalize(Arg::Range(ctx.values(&r))),
            Some(Err(e)) => Arg::Value(Value::Error(e)),
            None => Arg::Value(super::eval::eval_in(expr, ctx)),
        },
        Expr::Call(name, args) if LIFTED.contains(&name.as_str()) => lifted(name, args, ctx),
        other => Arg::Value(super::eval::eval_in(other, ctx)),
    }
}

fn normalize(a: Arg) -> Arg {
    match a {
        Arg::Range(r) if r.rows == 1 && r.cols == 1 => Arg::Value(r.values.into_iter().next().unwrap_or(Value::Empty)),
        Arg::Range(r) if r.rows == 0 || r.cols == 0 => Arg::Value(Value::Error(ErrorKind::Calc)),
        other => other,
    }
}

fn map(r: RangeValues, f: impl Fn(Value) -> Value) -> Arg {
    let values = r.values.into_iter().map(f).collect();
    arr(r.rows, r.cols, values)
}

/// Element by element with Excel's broadcasting: a single row or column (or
/// value) is repeated to match the other; cells beyond a smaller array are `#N/A`.
fn broadcast2(a: RangeValues, b: RangeValues, f: impl Fn(Value, Value) -> Value) -> Arg {
    let rows = a.rows.max(b.rows);
    let cols = a.cols.max(b.cols);
    if too_big(rows as u64, cols as u64) {
        return Arg::Value(Value::Error(ErrorKind::Num));
    }
    let pick = |m: &RangeValues, r: u32, c: u32| -> Value {
        let rr = if m.rows == 1 { 0 } else { r };
        let cc = if m.cols == 1 { 0 } else { c };
        m.get(rr, cc).cloned().unwrap_or(Value::Error(ErrorKind::NA))
    };
    let mut values = Vec::with_capacity((rows * cols) as usize);
    for r in 0..rows {
        for c in 0..cols {
            values.push(f(pick(&a, r, c), pick(&b, r, c)));
        }
    }
    arr(rows, cols, values)
}

/// A one-value function applied to each element of its array arguments.
fn lifted(name: &str, args: &[Expr], ctx: &Ctx) -> Arg {
    let evaluated: Vec<Arg> = args.iter().map(|a| eval_any(a, ctx)).collect();
    if evaluated.iter().all(|a| matches!(a, Arg::Value(_))) {
        // No arrays: the ordinary function (IF and friends keep their lazy evaluation there).
        return Arg::Value(super::eval::eval_in(&Expr::Call(name.to_string(), args.to_vec()), ctx));
    }
    let Some(f) = functions::lookup(name) else { return Arg::Value(Value::Error(ErrorKind::Name)) };
    let arrays: Vec<RangeValues> = evaluated.into_iter().map(as_array).collect();
    let rows = arrays.iter().map(|a| a.rows).max().unwrap_or(1);
    let cols = arrays.iter().map(|a| a.cols).max().unwrap_or(1);
    if too_big(rows as u64, cols as u64) {
        return Arg::Value(Value::Error(ErrorKind::Num));
    }
    let mut values = Vec::with_capacity((rows * cols) as usize);
    for r in 0..rows {
        for c in 0..cols {
            let call_args: Vec<Arg> = arrays
                .iter()
                .map(|a| {
                    let rr = if a.rows == 1 { 0 } else { r };
                    let cc = if a.cols == 1 { 0 } else { c };
                    Arg::Value(a.get(rr, cc).cloned().unwrap_or(Value::Error(ErrorKind::NA)))
                })
                .collect();
            values.push(match (f.call)(&call_args) {
                Ok(Value::Number(n)) if !n.is_finite() => Value::Error(ErrorKind::Num),
                Ok(v) => v,
                Err(e) => Value::Error(e),
            });
        }
    }
    arr(rows, cols, values)
}

// ---------------------------------------------------------------- argument helpers

fn arg_array(args: &[Expr], i: usize, ctx: &Ctx) -> Result<RangeValues, ErrorKind> {
    let a = as_array(eval_any(args.get(i).ok_or(ErrorKind::Value)?, ctx));
    if a.rows == 1 && a.cols == 1 {
        if let Value::Error(e) = &a.values[0] {
            return Err(*e);
        }
    }
    Ok(a)
}

fn arg_value(args: &[Expr], i: usize, ctx: &Ctx) -> Option<Value> {
    let e = args.get(i)?;
    // An omitted argument (`SORT(A1:A5,,-1)`) is parsed as an empty text or missing.
    let v = match eval_any(e, ctx) {
        Arg::Value(v) => v,
        Arg::Range(r) => r.values.into_iter().next().unwrap_or(Value::Empty),
    };
    Some(v)
}

fn num_or(args: &[Expr], i: usize, ctx: &Ctx, default: f64) -> Result<f64, ErrorKind> {
    match arg_value(args, i, ctx) {
        None | Some(Value::Empty) => Ok(default),
        Some(v) => v.as_number(),
    }
}

fn bool_or(args: &[Expr], i: usize, ctx: &Ctx, default: bool) -> Result<bool, ErrorKind> {
    match arg_value(args, i, ctx) {
        None | Some(Value::Empty) => Ok(default),
        Some(v) => v.as_bool(),
    }
}

fn truthy(v: &Value) -> Result<bool, ErrorKind> {
    match v {
        Value::Error(e) => Err(*e),
        Value::Empty => Ok(false),
        Value::Text(_) => Err(ErrorKind::Value),
        other => other.as_bool(),
    }
}

fn row(a: &RangeValues, r: u32) -> Vec<Value> {
    (0..a.cols).map(|c| a.get(r, c).cloned().unwrap_or(Value::Empty)).collect()
}

fn from_rows(rows: Vec<Vec<Value>>) -> Result<Arg, ErrorKind> {
    let n = rows.len() as u32;
    if n == 0 {
        return Err(ErrorKind::Calc);
    }
    let cols = rows.iter().map(Vec::len).max().unwrap_or(0) as u32;
    if cols == 0 {
        return Err(ErrorKind::Calc);
    }
    let mut values = Vec::with_capacity((n * cols) as usize);
    for mut r in rows {
        r.resize(cols as usize, Value::Error(ErrorKind::NA));
        values.extend(r);
    }
    Ok(arr(n, cols, values))
}

fn transpose(a: &RangeValues) -> RangeValues {
    let mut values = Vec::with_capacity(a.values.len());
    for c in 0..a.cols {
        for r in 0..a.rows {
            values.push(a.get(r, c).cloned().unwrap_or(Value::Empty));
        }
    }
    RangeValues { rows: a.cols, cols: a.rows, values }
}

/// Excel's sort order: numbers < text < booleans; errors and blanks last.
fn sort_cmp(a: &Value, b: &Value) -> Ordering {
    let rank = |v: &Value| match v {
        Value::Empty => 4,
        Value::Error(_) => 3,
        _ => 0,
    };
    match (rank(a), rank(b)) {
        (0, 0) => compare(a, b),
        (x, y) => x.cmp(&y),
    }
}

/// Text compared without case for UNIQUE, like Excel.
fn unique_key(v: &Value) -> String {
    match v {
        Value::Text(t) => format!("t{}", t.to_lowercase()),
        Value::Number(n) => format!("n{n}"),
        Value::Bool(b) => format!("b{b}"),
        Value::Error(e) => format!("e{}", e.as_str()),
        Value::Empty => "_".into(),
    }
}

// ---------------------------------------------------------------- the functions

pub(crate) fn call(name: &str, args: &[Expr], ctx: &Ctx) -> Result<Arg, ErrorKind> {
    match name {
        "LET" => let_(args, ctx),
        "SEQUENCE" => {
            let rows = num_or(args, 0, ctx, 1.0)?.trunc();
            let cols = num_or(args, 1, ctx, 1.0)?.trunc();
            let start = num_or(args, 2, ctx, 1.0)?;
            let step = num_or(args, 3, ctx, 1.0)?;
            if rows < 1.0 || cols < 1.0 {
                return Err(ErrorKind::Calc);
            }
            if too_big(rows as u64, cols as u64) {
                return Err(ErrorKind::Num);
            }
            let n = (rows * cols) as usize;
            Ok(arr(rows as u32, cols as u32, (0..n).map(|i| Value::Number(start + step * i as f64)).collect()))
        }
        "RANDARRAY" => {
            let rows = num_or(args, 0, ctx, 1.0)?.trunc();
            let cols = num_or(args, 1, ctx, 1.0)?.trunc();
            let min = num_or(args, 2, ctx, 0.0)?;
            let max = num_or(args, 3, ctx, 1.0)?;
            let whole = bool_or(args, 4, ctx, false)?;
            if rows < 1.0 || cols < 1.0 || min > max {
                return Err(if min > max { ErrorKind::Value } else { ErrorKind::Calc });
            }
            if too_big(rows as u64, cols as u64) {
                return Err(ErrorKind::Num);
            }
            let n = (rows * cols) as usize;
            let values = (0..n)
                .map(|_| {
                    let r = crate::functions::random();
                    Value::Number(if whole {
                        (min + (r * (max - min + 1.0)).floor()).min(max)
                    } else {
                        min + r * (max - min)
                    })
                })
                .collect();
            Ok(arr(rows as u32, cols as u32, values))
        }
        "TRANSPOSE" => {
            let a = arg_array(args, 0, ctx)?;
            let t = transpose(&a);
            Ok(arr(t.rows, t.cols, t.values))
        }
        "FILTER" => {
            let a = arg_array(args, 0, ctx)?;
            let include = arg_array(args, 1, ctx)?;
            let keep: Vec<bool> = include.values.iter().map(truthy).collect::<Result<_, _>>()?;
            let out = if include.cols == 1 && include.rows == a.rows {
                (0..a.rows).filter(|r| keep[*r as usize]).map(|r| row(&a, r)).collect::<Vec<_>>()
            } else if include.rows == 1 && include.cols == a.cols {
                let cols: Vec<u32> = (0..a.cols).filter(|c| keep[*c as usize]).collect();
                (0..a.rows)
                    .map(|r| cols.iter().map(|c| a.get(r, *c).cloned().unwrap_or(Value::Empty)).collect())
                    .collect()
            } else {
                return Err(ErrorKind::Value);
            };
            if out.is_empty() || out[0].is_empty() {
                return match args.get(2) {
                    Some(e) => Ok(eval_any(e, ctx)),
                    None => Err(ErrorKind::Calc),
                };
            }
            from_rows(out)
        }
        "SORT" => {
            let a = arg_array(args, 0, ctx)?;
            let index = num_or(args, 1, ctx, 1.0)?.trunc() as i64;
            let order = num_or(args, 2, ctx, 1.0)?;
            let by_col = bool_or(args, 3, ctx, false)?;
            let a = if by_col { transpose(&a) } else { a };
            if index < 1 || index as u32 > a.cols || (order != 1.0 && order != -1.0) {
                return Err(ErrorKind::Value);
            }
            let mut rows: Vec<Vec<Value>> = (0..a.rows).map(|r| row(&a, r)).collect();
            let k = (index - 1) as usize;
            rows.sort_by(|x, y| {
                let o = sort_cmp(&x[k], &y[k]);
                if order < 0.0 && !matches!((&x[k], &y[k]), (Value::Empty, _) | (_, Value::Empty)) {
                    o.reverse()
                } else {
                    o
                }
            });
            let out = from_rows(rows)?;
            Ok(if by_col { Arg::Range(transpose(&as_array(out))) } else { out })
        }
        "SORTBY" => {
            let a = arg_array(args, 0, ctx)?;
            let mut keys: Vec<(RangeValues, f64)> = Vec::new();
            let mut i = 1;
            while i < args.len() {
                let by = arg_array(args, i, ctx)?;
                let order = num_or(args, i + 1, ctx, 1.0)?;
                if by.values.len() != a.rows as usize && by.values.len() != a.cols as usize {
                    return Err(ErrorKind::Value);
                }
                keys.push((by, order));
                i += 2;
            }
            let mut idx: Vec<u32> = (0..a.rows).collect();
            idx.sort_by(|x, y| {
                for (by, order) in &keys {
                    let (p, q) = (&by.values[*x as usize], &by.values[*y as usize]);
                    let o = sort_cmp(p, q);
                    let o = if *order < 0.0 { o.reverse() } else { o };
                    if o != Ordering::Equal {
                        return o;
                    }
                }
                Ordering::Equal
            });
            from_rows(idx.into_iter().map(|r| row(&a, r)).collect())
        }
        "UNIQUE" => {
            let a = arg_array(args, 0, ctx)?;
            let by_col = bool_or(args, 1, ctx, false)?;
            let exactly_once = bool_or(args, 2, ctx, false)?;
            let a = if by_col { transpose(&a) } else { a };
            let rows: Vec<Vec<Value>> = (0..a.rows).map(|r| row(&a, r)).collect();
            let key = |r: &Vec<Value>| r.iter().map(unique_key).collect::<Vec<_>>().join("\u{1f}");
            let mut counts = std::collections::HashMap::new();
            for r in &rows {
                *counts.entry(key(r)).or_insert(0u32) += 1;
            }
            let mut seen = HashSet::new();
            let out: Vec<Vec<Value>> = rows
                .into_iter()
                .filter(|r| {
                    let k = key(r);
                    (!exactly_once || counts[&k] == 1) && seen.insert(k)
                })
                .collect();
            let out = from_rows(out)?;
            Ok(if by_col { Arg::Range(transpose(&as_array(out))) } else { out })
        }
        "TAKE" | "DROP" => {
            let a = arg_array(args, 0, ctx)?;
            let take = name == "TAKE";
            let rows = match arg_value(args, 1, ctx) {
                None | Some(Value::Empty) => None,
                Some(v) => Some(v.as_number()?.trunc() as i64),
            };
            let cols = match arg_value(args, 2, ctx) {
                None | Some(Value::Empty) => None,
                Some(v) => Some(v.as_number()?.trunc() as i64),
            };
            let span = |n: Option<i64>, len: u32| -> (u32, u32) {
                let len = len as i64;
                match n {
                    None => (0, len as u32),
                    Some(n) if take && n >= 0 => (0, n.min(len) as u32),
                    Some(n) if take => ((len + n).max(0) as u32, len as u32),
                    Some(n) if n >= 0 => (n.min(len) as u32, len as u32),
                    Some(n) => (0, (len + n).max(0) as u32),
                }
            };
            let (r0, r1) = span(rows, a.rows);
            let (c0, c1) = span(cols, a.cols);
            if r1 <= r0 || c1 <= c0 {
                return Err(ErrorKind::Calc);
            }
            from_rows(
                (r0..r1).map(|r| (c0..c1).map(|c| a.get(r, c).cloned().unwrap_or(Value::Empty)).collect()).collect(),
            )
        }
        "CHOOSECOLS" | "CHOOSEROWS" => {
            let a = arg_array(args, 0, ctx)?;
            let a = if name == "CHOOSECOLS" { transpose(&a) } else { a };
            let mut picked = Vec::new();
            for arg in &args[1..] {
                for v in as_array(eval_any(arg, ctx)).values {
                    let n = v.as_number()?.trunc() as i64;
                    let idx = if n < 0 { a.rows as i64 + n } else { n - 1 };
                    if n == 0 || idx < 0 || idx >= a.rows as i64 {
                        return Err(ErrorKind::Value);
                    }
                    picked.push(row(&a, idx as u32));
                }
            }
            let out = from_rows(picked)?;
            Ok(if name == "CHOOSECOLS" { Arg::Range(transpose(&as_array(out))) } else { out })
        }
        "VSTACK" | "HSTACK" => {
            let mut rows: Vec<Vec<Value>> = Vec::new();
            for i in 0..args.len() {
                let a = arg_array(args, i, ctx)?;
                let a = if name == "HSTACK" { transpose(&a) } else { a };
                rows.extend((0..a.rows).map(|r| row(&a, r)));
            }
            let out = from_rows(rows)?;
            Ok(if name == "HSTACK" { Arg::Range(transpose(&as_array(out))) } else { out })
        }
        "TOCOL" | "TOROW" => {
            let a = arg_array(args, 0, ctx)?;
            let ignore = num_or(args, 1, ctx, 0.0)? as i64;
            let by_col = bool_or(args, 2, ctx, false)?;
            let a = if by_col { transpose(&a) } else { a };
            let values: Vec<Value> = a
                .values
                .into_iter()
                .filter(|v| match ignore {
                    1 => !v.is_empty(),
                    2 => !matches!(v, Value::Error(_)),
                    3 => !v.is_empty() && !matches!(v, Value::Error(_)),
                    _ => true,
                })
                .collect();
            if values.is_empty() {
                return Err(ErrorKind::Calc);
            }
            let n = values.len() as u32;
            Ok(if name == "TOCOL" { arr(n, 1, values) } else { arr(1, n, values) })
        }
        "EXPAND" => {
            let a = arg_array(args, 0, ctx)?;
            let rows = num_or(args, 1, ctx, a.rows as f64)?.trunc() as u32;
            let cols = num_or(args, 2, ctx, a.cols as f64)?.trunc() as u32;
            let pad = arg_value(args, 3, ctx).unwrap_or(Value::Error(ErrorKind::NA));
            if rows < a.rows || cols < a.cols || too_big(rows as u64, cols as u64) {
                return Err(ErrorKind::Value);
            }
            from_rows(
                (0..rows)
                    .map(|r| (0..cols).map(|c| a.get(r, c).cloned().unwrap_or_else(|| pad.clone())).collect())
                    .collect(),
            )
        }
        "TEXTSPLIT" => {
            let text = arg_value(args, 0, ctx).ok_or(ErrorKind::Value)?.as_text()?;
            let col_delims: Vec<String> = match args.get(1) {
                Some(e) => as_array(eval_any(e, ctx)).values.iter().map(Value::as_text).collect::<Result<_, _>>()?,
                None => Vec::new(),
            };
            let row_delims: Vec<String> = match args.get(2) {
                Some(e) => as_array(eval_any(e, ctx))
                    .values
                    .iter()
                    .filter(|v| !v.is_empty())
                    .map(Value::as_text)
                    .collect::<Result<_, _>>()?,
                None => Vec::new(),
            };
            let ignore_empty = bool_or(args, 3, ctx, false)?;
            let split = |s: &str, delims: &[String]| -> Vec<String> {
                let delims: Vec<&String> = delims.iter().filter(|d| !d.is_empty()).collect();
                if delims.is_empty() {
                    return vec![s.to_string()];
                }
                let mut out = vec![String::new()];
                let mut rest = s;
                'outer: while !rest.is_empty() {
                    for d in &delims {
                        if let Some(r) = rest.strip_prefix(d.as_str()) {
                            out.push(String::new());
                            rest = r;
                            continue 'outer;
                        }
                    }
                    let ch = rest.chars().next().unwrap();
                    out.last_mut().unwrap().push(ch);
                    rest = &rest[ch.len_utf8()..];
                }
                if ignore_empty {
                    out.retain(|p| !p.is_empty());
                }
                out
            };
            let rows: Vec<Vec<Value>> = split(&text, &row_delims)
                .iter()
                .map(|line| split(line, &col_delims).into_iter().map(Value::Text).collect())
                .collect();
            from_rows(rows)
        }
        "XMATCH" => {
            let target = arg_value(args, 0, ctx).ok_or(ErrorKind::Value)?;
            let a = arg_array(args, 1, ctx)?;
            let mode = num_or(args, 2, ctx, 0.0)? as i64;
            let search = num_or(args, 3, ctx, 1.0)? as i64;
            if a.rows > 1 && a.cols > 1 {
                return Err(ErrorKind::Value);
            }
            let list = &a.values;
            let order: Box<dyn Iterator<Item = usize>> =
                if search < 0 { Box::new((0..list.len()).rev()) } else { Box::new(0..list.len()) };
            let eq = |v: &Value| match (v, &target) {
                (Value::Text(x), Value::Text(t)) if mode == 2 => crate::data_tools::wildcard_match(t, x),
                _ => compare(v, &target) == Ordering::Equal,
            };
            let mut best: Option<usize> = None;
            for i in order {
                let v = &list[i];
                if eq(v) {
                    return Ok(Arg::Value(Value::Number((i + 1) as f64)));
                }
                let better = |b: usize| match mode {
                    -1 => compare(v, &list[b]) == Ordering::Greater,
                    1 => compare(v, &list[b]) == Ordering::Less,
                    _ => false,
                };
                let ok = match mode {
                    -1 => compare(v, &target) == Ordering::Less && !v.is_empty(),
                    1 => compare(v, &target) == Ordering::Greater,
                    _ => false,
                };
                if ok && best.is_none_or(better) {
                    best = Some(i);
                }
            }
            best.map(|i| Arg::Value(Value::Number((i + 1) as f64))).ok_or(ErrorKind::NA)
        }
        _ => Err(ErrorKind::Name),
    }
}

/// LET(name1, value1, [name2, value2, ...], calculation)
fn let_(args: &[Expr], ctx: &Ctx) -> Result<Arg, ErrorKind> {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Err(ErrorKind::Value);
    }
    let depth = ctx.lets.borrow().len();
    let mut i = 0;
    while i + 1 < args.len() {
        let Expr::Name(name) = &args[i] else {
            ctx.lets.borrow_mut().truncate(depth);
            return Err(ErrorKind::Value);
        };
        let value = eval_any(&args[i + 1], ctx);
        ctx.lets.borrow_mut().push((name.to_ascii_uppercase(), value));
        i += 2;
    }
    let result = eval_any(&args[args.len() - 1], ctx);
    ctx.lets.borrow_mut().truncate(depth);
    Ok(result)
}

#[cfg(test)]
mod tests {
    use crate::{CellRef, Sheet, Value};

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn sheet(cells: &[(&str, &str)]) -> Sheet {
        let mut s = Sheet::new();
        s.set_inputs(cells.iter().map(|(a, v)| (at(a), v.to_string())));
        s
    }

    fn texts(s: &Sheet, cells: &[&str]) -> Vec<String> {
        cells.iter().map(|c| s.display(at(c))).collect()
    }

    #[test]
    fn sequence_spills_and_feeds_other_formulas() {
        let s = sheet(&[("A1", "=SEQUENCE(3,2,10,5)"), ("D1", "=SUM(A1:B3)"), ("E1", "=B3")]);
        assert_eq!(texts(&s, &["A1", "B1", "A2", "B3"]), ["10", "15", "20", "35"]);
        assert_eq!(s.value(at("D1")), &Value::Number(135.0));
        assert_eq!(s.value(at("E1")), &Value::Number(35.0));
    }

    #[test]
    fn spill_is_blocked_and_unblocked() {
        let mut s = sheet(&[("A1", "=SEQUENCE(3)"), ("A3", "x")]);
        assert_eq!(s.display(at("A1")), "#SPILL!");
        s.set_input(at("A3"), "");
        assert_eq!(texts(&s, &["A1", "A2", "A3"]), ["1", "2", "3"]);
        // Typing into the spill area blocks it again.
        s.set_input(at("A2"), "5");
        assert_eq!(s.display(at("A1")), "#SPILL!");
    }

    #[test]
    fn filter_sort_unique() {
        let s = sheet(&[
            ("A1", "Asha"),
            ("A2", "Ravi"),
            ("A3", "asha"),
            ("A4", "Meera"),
            ("B1", "120"),
            ("B2", "80"),
            ("B3", "200"),
            ("B4", "150"),
            ("D1", "=FILTER(A1:B4, B1:B4>100)"),
            ("G1", "=SORT(B1:B4,,-1)"),
            ("H1", "=UNIQUE(A1:A4)"),
            ("I1", "=SORTBY(A1:A4, B1:B4, -1)"),
        ]);
        assert_eq!(texts(&s, &["D1", "E1", "D2", "E2", "D3", "E3"]), ["Asha", "120", "asha", "200", "Meera", "150"]);
        assert_eq!(texts(&s, &["G1", "G2", "G3", "G4"]), ["200", "150", "120", "80"]);
        assert_eq!(texts(&s, &["H1", "H2", "H3"]), ["Asha", "Ravi", "Meera"]);
        assert_eq!(texts(&s, &["I1", "I2"]), ["asha", "Meera"]);
    }

    #[test]
    fn element_wise_operations_and_lifting() {
        let s = sheet(&[
            ("A1", "1"),
            ("A2", "2"),
            ("A3", "3"),
            ("B1", "=A1:A3*10"),
            ("C1", "=SUM(A1:A3*A1:A3)"),
            ("D1", "=IF(A1:A3>1,\"big\",\"small\")"),
            ("F1", "=LET(x, A3, y, 2, x*y)"),
            ("G1", "=XMATCH(3, A1:A3)"),
        ]);
        assert_eq!(texts(&s, &["B1", "B2", "B3"]), ["10", "20", "30"]);
        assert_eq!(s.value(at("C1")), &Value::Number(14.0));
        assert_eq!(texts(&s, &["D1", "D2", "D3"]), ["small", "big", "big"]);
        assert_eq!(s.value(at("F1")), &Value::Number(6.0));
        assert_eq!(s.value(at("G1")), &Value::Number(3.0));
    }

    #[test]
    fn stacking_and_shaping() {
        let s = sheet(&[
            ("A1", "1"),
            ("A2", "2"),
            ("B1", "3"),
            ("B2", "4"),
            ("D1", "=TRANSPOSE(A1:B2)"),
            ("G1", "=VSTACK(A1:B1, A2:B2, 9)"),
            ("J1", "=TAKE(A1:B2, -1)"),
            ("M1", "=TEXTSPLIT(\"a,b;c,d\", \",\", \";\")"),
            ("P1", "=TOCOL(A1:B2)"),
        ]);
        assert_eq!(texts(&s, &["D1", "E1", "D2", "E2"]), ["1", "2", "3", "4"]);
        assert_eq!(texts(&s, &["G3", "H3"]), ["9", "#N/A"]);
        assert_eq!(texts(&s, &["J1", "K1"]), ["2", "4"]);
        assert_eq!(texts(&s, &["M1", "N1", "M2", "N2"]), ["a", "b", "c", "d"]);
        assert_eq!(texts(&s, &["P1", "P2", "P3", "P4"]), ["1", "3", "2", "4"]);
    }
}
