use std::cmp::Ordering;

use super::ast::{BinOp, Expr, UnaryOp};
use crate::cell::{CellRef, Range};
use crate::functions::{self, Arg, RangeValues};
use crate::value::{compare, ErrorKind, Value};

/// Where the evaluator reads cell values from. `Sheet` implements this;
/// tests can use a plain `HashMap`.
pub trait CellSource {
    fn value(&self, at: CellRef) -> Value;

    fn range(&self, r: Range) -> RangeValues {
        RangeValues { rows: r.rows(), cols: r.cols(), values: r.cells().map(|c| self.value(c)).collect() }
    }
}

/// Evaluate to a single value.
pub fn eval(expr: &Expr, src: &dyn CellSource) -> Value {
    match expr {
        Expr::Number(n) => Value::Number(*n),
        Expr::Text(s) => Value::Text(s.clone()),
        Expr::Bool(b) => Value::Bool(*b),
        Expr::Ref(c) => src.value(*c),
        // A multi-cell range where one value is expected.
        Expr::Range(r) if r.start == r.end => src.value(r.start),
        Expr::Range(_) => Value::Error(ErrorKind::Value),
        Expr::Name(_) => Value::Error(ErrorKind::Name),
        Expr::Error(e) => Value::Error(*e),
        Expr::Unary(op, e) => match (op, eval(e, src).as_number()) {
            (_, Err(e)) => Value::Error(e),
            (UnaryOp::Neg, Ok(n)) => Value::Number(-n),
            (UnaryOp::Percent, Ok(n)) => Value::Number(n / 100.0),
        },
        Expr::Binary(op, a, b) => binary(*op, eval(a, src), eval(b, src)),
        Expr::Call(name, args) => call(name, args, src),
    }
}

/// Evaluate a function argument: ranges stay ranges so `SUM(A1:A9)` sees every cell.
fn eval_arg(expr: &Expr, src: &dyn CellSource) -> Arg {
    match expr {
        Expr::Range(r) => Arg::Range(src.range(*r)),
        other => Arg::Value(eval(other, src)),
    }
}

fn call(name: &str, args: &[Expr], src: &dyn CellSource) -> Value {
    let Some(f) = functions::lookup(name) else {
        return Value::Error(ErrorKind::Name);
    };
    if args.len() < f.min_args || f.max_args.is_some_and(|max| args.len() > max) {
        return Value::Error(ErrorKind::Value);
    }
    let args: Vec<Arg> = args.iter().map(|a| eval_arg(a, src)).collect();
    match (f.call)(&args) {
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
    }

    fn run(src: &str) -> Value {
        let cells = Cells(HashMap::from([
            (CellRef::new(0, 0), Value::Number(2.0)),
            (CellRef::new(1, 0), Value::Text("x".into())),
        ]));
        eval(&parse(src).unwrap(), &cells)
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
    }
}
