use std::cmp::Ordering;

use super::Arg;
use crate::value::{compare, ErrorKind, Value};

pub fn if_(args: &[Arg]) -> Result<Value, ErrorKind> {
    if args[0].boolean()? {
        Ok(args[1].scalar())
    } else {
        Ok(args.get(2).map_or(Value::Bool(false), Arg::scalar))
    }
}

pub fn iferror(args: &[Arg]) -> Result<Value, ErrorKind> {
    match args[0].scalar() {
        Value::Error(_) => Ok(args[1].scalar()),
        v => Ok(v),
    }
}

pub fn ifna(args: &[Arg]) -> Result<Value, ErrorKind> {
    match args[0].scalar() {
        Value::Error(ErrorKind::NA) => Ok(args[1].scalar()),
        v => Ok(v),
    }
}

/// IFS(condition1, value1, ...): the value for the first TRUE condition, `#N/A` if none.
pub fn ifs(args: &[Arg]) -> Result<Value, ErrorKind> {
    if !args.len().is_multiple_of(2) {
        return Err(ErrorKind::Value);
    }
    for pair in args.chunks(2) {
        if pair[0].boolean()? {
            return Ok(pair[1].scalar());
        }
    }
    Err(ErrorKind::NA)
}

/// SWITCH(expression, value1, result1, ..., [default]). Values must have the
/// same type as the expression to match (text ignores case); `#N/A` if none.
pub fn switch(args: &[Arg]) -> Result<Value, ErrorKind> {
    let target = args[0].scalar();
    if let Value::Error(e) = target {
        return Err(e);
    }
    for pair in args[1..].chunks(2) {
        let [case, result] = pair else {
            return Ok(pair[0].scalar());
        };
        let case = case.scalar();
        if let Value::Error(e) = case {
            return Err(e);
        }
        let same_type = matches!(
            (&target, &case),
            (Value::Number(_), Value::Number(_)) | (Value::Text(_), Value::Text(_)) | (Value::Bool(_), Value::Bool(_))
        );
        if (same_type || target.is_empty()) && compare(&target, &case) == Ordering::Equal {
            return Ok(result.scalar());
        }
    }
    Err(ErrorKind::NA)
}

pub fn and(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut result = true;
    each_bool(args, |b| result &= b)?;
    Ok(Value::Bool(result))
}

pub fn or(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut result = false;
    each_bool(args, |b| result |= b)?;
    Ok(Value::Bool(result))
}

/// TRUE when an odd number of the arguments are TRUE; `#VALUE!` if there are no logical values.
pub fn xor(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (mut result, mut seen) = (false, false);
    each_bool(args, |b| {
        result ^= b;
        seen = true;
    })?;
    if !seen {
        return Err(ErrorKind::Value);
    }
    Ok(Value::Bool(result))
}

pub fn not(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(!args[0].boolean()?))
}

pub fn true_(_: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(true))
}

pub fn false_(_: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(false))
}

/// Ranges contribute only their numbers and booleans; scalars are coerced.
fn each_bool(args: &[Arg], mut f: impl FnMut(bool)) -> Result<(), ErrorKind> {
    for arg in args {
        match arg {
            Arg::Value(v) => f(v.as_bool()?),
            Arg::Range(r) => {
                for v in &r.values {
                    match v {
                        Value::Bool(_) | Value::Number(_) => f(v.as_bool()?),
                        Value::Error(e) => return Err(*e),
                        _ => {}
                    }
                }
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use crate::functions::test_util::{n, run, t};
    use crate::{ErrorKind, Value};

    #[test]
    fn logic() {
        assert_eq!(run("IF(A1>0, \"yes\", \"no\")"), t("yes"));
        assert_eq!(run("IF(FALSE, 1)"), Value::Bool(false));
        assert_eq!(run("IFERROR(1/0, 42)"), n(42.0));
        assert_eq!(run("IFERROR(7, 42)"), n(7.0));
        assert_eq!(run("AND(TRUE, A1:A5)"), Value::Bool(true));
        assert_eq!(run("OR(FALSE, 0)"), Value::Bool(false));
        assert_eq!(run("NOT(C1)"), Value::Bool(false));
        assert_eq!(run("TRUE()"), Value::Bool(true));
    }

    #[test]
    fn multi_branch_logic() {
        assert_eq!(run("IFS(A1>5, \"big\", A1>0, \"small\")"), t("small"));
        assert_eq!(run("IFS(FALSE, 1)"), Value::Error(ErrorKind::NA));
        assert_eq!(run("IFS(\"x\", 1)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("IFS(TRUE, 1, FALSE)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("SWITCH(A2, 1, \"one\", 2, \"two\")"), t("two"));
        assert_eq!(run("SWITCH(B1, \"A\", \"first\", \"other\")"), t("first"));
        assert_eq!(run("SWITCH(9, 1, \"one\", \"other\")"), t("other"));
        assert_eq!(run("SWITCH(1, \"1\", \"text\")"), Value::Error(ErrorKind::NA));
        assert_eq!(run("SWITCH(1/0, 1, 2)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("XOR(TRUE, FALSE)"), Value::Bool(true));
        assert_eq!(run("XOR(TRUE, C1)"), Value::Bool(false));
        assert_eq!(run("XOR(A1:A3)"), Value::Bool(true));
        assert_eq!(run("XOR(B1:B5)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("IFNA(NA(), \"none\")"), t("none"));
        assert_eq!(run("IFNA(1/0, \"none\")"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("IFNA(5, 0)"), n(5.0));
    }
}
