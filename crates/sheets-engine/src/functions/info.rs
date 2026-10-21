use super::Arg;
use crate::value::{ErrorKind, Value};

pub fn isblank(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(matches!(args[0].scalar(), Value::Empty)))
}

pub fn isnumber(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(matches!(args[0].scalar(), Value::Number(_))))
}

pub fn istext(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(matches!(args[0].scalar(), Value::Text(_))))
}

pub fn iserror(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(matches!(args[0].scalar(), Value::Error(_))))
}

pub fn isna(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(matches!(args[0].scalar(), Value::Error(ErrorKind::NA))))
}

/// Any error except `#N/A`.
pub fn iserr(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(matches!(args[0].scalar(), Value::Error(e) if e != ErrorKind::NA)))
}

pub fn islogical(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(matches!(args[0].scalar(), Value::Bool(_))))
}

/// TRUE for anything that is not text, including empty cells and errors.
pub fn isnontext(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(!matches!(args[0].scalar(), Value::Text(_))))
}

pub fn iseven(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(integer_part(&args[0])? % 2.0 == 0.0))
}

pub fn isodd(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(integer_part(&args[0])? % 2.0 != 0.0))
}

/// ISEVEN/ISODD argument: truncated number; booleans are `#VALUE!` as in Excel.
fn integer_part(arg: &Arg) -> Result<f64, ErrorKind> {
    match arg.scalar() {
        Value::Bool(_) => Err(ErrorKind::Value),
        v => Ok(v.as_number()?.trunc()),
    }
}

pub fn na(_: &[Arg]) -> Result<Value, ErrorKind> {
    Err(ErrorKind::NA)
}

#[cfg(test)]
mod tests {
    use crate::functions::test_util::run;
    use crate::{ErrorKind, Value};

    #[test]
    fn info_functions() {
        assert_eq!(run("ISBLANK(Z9)"), Value::Bool(true));
        assert_eq!(run("ISBLANK(D1)"), Value::Bool(false));
        assert_eq!(run("ISNUMBER(A1)"), Value::Bool(true));
        assert_eq!(run("ISTEXT(B1)"), Value::Bool(true));
        assert_eq!(run("ISERROR(1/0)"), Value::Bool(true));
        assert_eq!(run("NA()"), Value::Error(ErrorKind::NA));
    }

    #[test]
    fn more_info_functions() {
        assert_eq!(run("ISNA(NA())"), Value::Bool(true));
        assert_eq!(run("ISNA(1/0)"), Value::Bool(false));
        assert_eq!(run("ISERR(1/0)"), Value::Bool(true));
        assert_eq!(run("ISERR(NA())"), Value::Bool(false));
        assert_eq!(run("ISERR(1)"), Value::Bool(false));
        assert_eq!(run("ISLOGICAL(C1)"), Value::Bool(true));
        assert_eq!(run("ISLOGICAL(1)"), Value::Bool(false));
        assert_eq!(run("ISNONTEXT(A1)"), Value::Bool(true));
        assert_eq!(run("ISNONTEXT(Z9)"), Value::Bool(true));
        assert_eq!(run("ISNONTEXT(D1)"), Value::Bool(false));
        assert_eq!(run("ISEVEN(A2)"), Value::Bool(true));
        assert_eq!(run("ISEVEN(-3.9)"), Value::Bool(false));
        assert_eq!(run("ISEVEN(Z9)"), Value::Bool(true));
        assert_eq!(run("ISODD(A3)"), Value::Bool(true));
        assert_eq!(run("ISODD(\"7\")"), Value::Bool(true));
        assert_eq!(run("ISODD(-3)"), Value::Bool(true));
        assert_eq!(run("ISODD(\"x\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("ISEVEN(TRUE)"), Value::Error(ErrorKind::Value));
    }
}
