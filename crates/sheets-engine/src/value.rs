//! Cell values and Excel-compatible coercion rules.

use std::cmp::Ordering;

/// Errors a formula can produce. They are values, not Rust errors: they flow
/// through formulas until a function like `IFERROR` handles them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorKind {
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NA,
    /// Circular reference (Excel shows a warning instead; we show `#CIRC!`).
    Circ,
    /// Formula could not be parsed (Google Sheets shows `#ERROR!` too).
    Parse,
}

impl ErrorKind {
    pub fn as_str(&self) -> &'static str {
        match self {
            ErrorKind::Div0 => "#DIV/0!",
            ErrorKind::Value => "#VALUE!",
            ErrorKind::Ref => "#REF!",
            ErrorKind::Name => "#NAME?",
            ErrorKind::Num => "#NUM!",
            ErrorKind::NA => "#N/A",
            ErrorKind::Circ => "#CIRC!",
            ErrorKind::Parse => "#ERROR!",
        }
    }

    pub fn parse(s: &str) -> Option<Self> {
        Some(match s.to_ascii_uppercase().as_str() {
            "#DIV/0!" => ErrorKind::Div0,
            "#VALUE!" => ErrorKind::Value,
            "#REF!" => ErrorKind::Ref,
            "#NAME?" => ErrorKind::Name,
            "#NUM!" => ErrorKind::Num,
            "#N/A" => ErrorKind::NA,
            "#CIRC!" => ErrorKind::Circ,
            "#ERROR!" => ErrorKind::Parse,
            _ => return None,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub enum Value {
    #[default]
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(ErrorKind),
}

impl Value {
    /// Interpret a non-formula input the way Excel does when you type it.
    /// A leading `'` forces text: `'00123` stays the text "00123".
    pub fn from_input(s: &str) -> Value {
        if s.is_empty() {
            return Value::Empty;
        }
        if let Some(text) = s.strip_prefix('\'') {
            return Value::Text(text.to_string());
        }
        let trimmed = s.trim();
        if let Ok(n) = trimmed.parse::<f64>() {
            if n.is_finite() {
                return Value::Number(n);
            }
        }
        if let Some((n, _)) = crate::format::parse_formatted(trimmed) {
            return Value::Number(n);
        }
        if trimmed.eq_ignore_ascii_case("TRUE") {
            return Value::Bool(true);
        }
        if trimmed.eq_ignore_ascii_case("FALSE") {
            return Value::Bool(false);
        }
        if let Some(e) = ErrorKind::parse(trimmed) {
            return Value::Error(e);
        }
        Value::Text(s.to_string())
    }

    /// Text shown in the grid.
    pub fn display(&self) -> String {
        match self {
            Value::Empty => String::new(),
            Value::Number(n) => format_number(*n),
            Value::Text(s) => s.clone(),
            Value::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            Value::Error(e) => e.as_str().to_string(),
        }
    }

    pub fn as_number(&self) -> Result<f64, ErrorKind> {
        match self {
            Value::Empty => Ok(0.0),
            Value::Number(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Text(s) => s.trim().parse::<f64>().ok().filter(|n| n.is_finite()).ok_or(ErrorKind::Value),
            Value::Error(e) => Err(*e),
        }
    }

    pub fn as_text(&self) -> Result<String, ErrorKind> {
        match self {
            Value::Error(e) => Err(*e),
            other => Ok(other.display()),
        }
    }

    pub fn as_bool(&self) -> Result<bool, ErrorKind> {
        match self {
            Value::Empty => Ok(false),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Bool(b) => Ok(*b),
            Value::Text(s) if s.eq_ignore_ascii_case("TRUE") => Ok(true),
            Value::Text(s) if s.eq_ignore_ascii_case("FALSE") => Ok(false),
            Value::Text(_) => Err(ErrorKind::Value),
            Value::Error(e) => Err(*e),
        }
    }

    /// An input that reproduces this value exactly when typed (used by
    /// "paste values" and file import). Text that would be read as a number,
    /// boolean or formula gets a leading `'`.
    pub fn to_input(&self) -> String {
        match self {
            Value::Empty => String::new(),
            Value::Number(n) => format!("{n}"),
            Value::Text(s) if Value::text_round_trips(s) => s.clone(),
            Value::Text(s) => format!("'{s}"),
            Value::Bool(_) | Value::Error(_) => self.display(),
        }
    }

    /// True when typing `s` back gives exactly the text `s` (no leading `'`
    /// needed). Cheap for ordinary text.
    pub fn text_round_trips(s: &str) -> bool {
        is_plain_text(s) || (!s.starts_with('=') && Value::from_input(s) == Value::Text(s.to_string()))
    }

    pub fn is_empty(&self) -> bool {
        matches!(self, Value::Empty)
    }
}

/// Fast check for text that typing back can't turn into anything else:
/// it starts with a letter and isn't TRUE/FALSE or a "Rs." amount. Saves the
/// full parse for most text when loading big files.
fn is_plain_text(s: &str) -> bool {
    let Some(first) = s.chars().next() else { return false };
    first.is_alphabetic()
        && !s.eq_ignore_ascii_case("TRUE")
        && !s.eq_ignore_ascii_case("FALSE")
        && !s.starts_with("Rs.")
        && !s.starts_with("inf")
        && !s.starts_with("Inf")
        && !s.starts_with("nan")
        && !s.starts_with("NaN")
}

/// Excel ordering for comparisons: numbers < text < booleans; text is
/// case-insensitive; an empty cell acts as 0, "" or FALSE depending on the
/// other side.
pub fn compare(a: &Value, b: &Value) -> Ordering {
    fn rank(v: &Value) -> u8 {
        match v {
            Value::Number(_) | Value::Empty | Value::Error(_) => 0,
            Value::Text(_) => 1,
            Value::Bool(_) => 2,
        }
    }
    let blank_like = |other: &Value| match other {
        Value::Text(_) => Value::Text(String::new()),
        Value::Bool(_) => Value::Bool(false),
        _ => Value::Number(0.0),
    };
    let (a, b) = match (a, b) {
        (Value::Empty, Value::Empty) => return Ordering::Equal,
        (Value::Empty, other) => (blank_like(other), other.clone()),
        (other, Value::Empty) => (other.clone(), blank_like(other)),
        _ => (a.clone(), b.clone()),
    };
    match (&a, &b) {
        (Value::Number(x), Value::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Value::Text(x), Value::Text(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
        (Value::Bool(x), Value::Bool(y)) => x.cmp(y),
        _ => rank(&a).cmp(&rank(&b)),
    }
}

/// "General" number format: up to 15 significant digits, no trailing zeros.
pub fn format_number(n: f64) -> String {
    if !n.is_finite() {
        return ErrorKind::Num.as_str().to_string();
    }
    if n == 0.0 {
        return "0".to_string();
    }
    // Round to 15 significant digits to hide binary noise (0.1 + 0.2).
    let rounded: f64 = format!("{:.14e}", n).parse().unwrap_or(n);
    let abs = rounded.abs();
    if abs >= 1e15 || abs < 1e-9 {
        let s = format!("{:e}", rounded);
        return s.replace('e', "E");
    }
    format!("{}", rounded)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_inputs() {
        assert_eq!(Value::from_input("42"), Value::Number(42.0));
        assert_eq!(Value::from_input(" -1.5e3 "), Value::Number(-1500.0));
        assert_eq!(Value::from_input("true"), Value::Bool(true));
        assert_eq!(Value::from_input("'00123"), Value::Text("00123".into()));
        assert_eq!(Value::from_input("inf"), Value::Text("inf".into()));
        assert_eq!(Value::from_input("NaN"), Value::Text("NaN".into()));
        assert_eq!(Value::from_input("#N/A"), Value::Error(ErrorKind::NA));
        assert_eq!(Value::from_input("hello"), Value::Text("hello".into()));
    }

    #[test]
    fn values_round_trip_through_input() {
        for v in [
            Value::Number(0.1 + 0.2),
            Value::Number(-1e20),
            Value::Text("00123".into()),
            Value::Text("=not a formula".into()),
            Value::Text("hello".into()),
            Value::Text("12%".into()),
            Value::Bool(false),
            Value::Error(ErrorKind::NA),
        ] {
            assert_eq!(Value::from_input(&v.to_input()), v);
        }
    }

    #[test]
    fn formats_numbers() {
        assert_eq!(format_number(0.1 + 0.2), "0.3");
        assert_eq!(format_number(50.0), "50");
        assert_eq!(format_number(-2.5), "-2.5");
        assert_eq!(format_number(1e20), "1E20");
    }

    #[test]
    fn compares_like_excel() {
        assert_eq!(compare(&Value::Text("abc".into()), &Value::Text("ABC".into())), Ordering::Equal);
        assert_eq!(compare(&Value::Number(1e9), &Value::Text("a".into())), Ordering::Less);
        assert_eq!(compare(&Value::Empty, &Value::Number(0.0)), Ordering::Equal);
    }
}
