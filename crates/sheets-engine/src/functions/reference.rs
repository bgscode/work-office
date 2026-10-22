//! Reference functions. `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `OFFSET`,
//! `INDIRECT`, `SUBTOTAL` and `AGGREGATE` need the formula's cell or the
//! references themselves, so the evaluator runs them in
//! `formula/special.rs`; their registry entries point at [`special`], which
//! is never reached. `ADDRESS` only needs values and lives here.

use super::Arg;
use crate::cell::{col_name, MAX_COLS, MAX_ROWS};
use crate::formula::refs::quote_sheet_name;
use crate::value::{ErrorKind, Value};

/// Placeholder for functions evaluated by `formula/special.rs`.
pub fn special(_: &[Arg]) -> Result<Value, ErrorKind> {
    Err(ErrorKind::Value)
}

/// ADDRESS(row, column, [abs_num], [a1], [sheet_text]): a reference as text.
/// abs_num 1 `$A$1`, 2 `A$1`, 3 `$A1`, 4 `A1`; a1 FALSE gives R1C1 style.
pub fn address(args: &[Arg]) -> Result<Value, ErrorKind> {
    let row = args[0].number()?.trunc();
    let col = args[1].number()?.trunc();
    let abs = match args.get(2) {
        Some(a) => a.number()?.trunc() as i64,
        None => 1,
    };
    let a1 = match args.get(3) {
        Some(a) => a.boolean()?,
        None => true,
    };
    if row < 1.0 || col < 1.0 || row > MAX_ROWS as f64 || col > MAX_COLS as f64 || !(1..=4).contains(&abs) {
        return Err(ErrorKind::Value);
    }
    let (abs_row, abs_col) = (abs == 1 || abs == 2, abs == 1 || abs == 3);
    let (row, col) = (row as u32, col as u32);
    let text = if a1 {
        let d = |on: bool| if on { "$" } else { "" };
        format!("{}{}{}{}", d(abs_col), col_name(col - 1), d(abs_row), row)
    } else {
        let part = |on: bool, n: u32| if on { n.to_string() } else { format!("[{n}]") };
        format!("R{}C{}", part(abs_row, row), part(abs_col, col))
    };
    let sheet = match args.get(4) {
        Some(a) => a.text()?,
        None => String::new(),
    };
    Ok(Value::Text(if sheet.is_empty() { text } else { format!("{}!{text}", quote_sheet_name(&sheet)) }))
}

#[cfg(test)]
mod tests {
    use crate::functions::test_util::{run, t};
    use crate::{ErrorKind, Value};

    #[test]
    fn address() {
        assert_eq!(run("ADDRESS(2, 3)"), t("$C$2"));
        assert_eq!(run("ADDRESS(2, 3, 2)"), t("C$2"));
        assert_eq!(run("ADDRESS(2, 3, 3)"), t("$C2"));
        assert_eq!(run("ADDRESS(2, 3, 4)"), t("C2"));
        assert_eq!(run("ADDRESS(2, 3, 1, FALSE)"), t("R2C3"));
        assert_eq!(run("ADDRESS(2, 3, 4, FALSE)"), t("R[2]C[3]"));
        assert_eq!(run("ADDRESS(1, 1, 1, TRUE, \"Book1\")"), t("Book1!$A$1"));
        assert_eq!(run("ADDRESS(1, 28, 1, TRUE, \"My Sheet\")"), t("'My Sheet'!$AB$1"));
        assert_eq!(run("ADDRESS(0, 1)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("ADDRESS(1, 1, 5)"), Value::Error(ErrorKind::Value));
    }
}
