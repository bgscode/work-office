use super::conditional::wildcard_match;
use super::format::format_value;
use super::{all_values, datetime, Arg};
use crate::value::{ErrorKind, Value};

pub fn len(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(args[0].text()?.chars().count() as f64))
}

pub fn upper(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Text(args[0].text()?.to_uppercase()))
}

pub fn lower(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Text(args[0].text()?.to_lowercase()))
}

/// Removes leading/trailing spaces and collapses inner runs to one space.
pub fn trim(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Text(args[0].text()?.split(' ').filter(|s| !s.is_empty()).collect::<Vec<_>>().join(" ")))
}

pub fn left(args: &[Arg]) -> Result<Value, ErrorKind> {
    let s = args[0].text()?;
    let n = count_arg(args.get(1))?;
    Ok(Value::Text(s.chars().take(n).collect()))
}

pub fn right(args: &[Arg]) -> Result<Value, ErrorKind> {
    let s = args[0].text()?;
    let n = count_arg(args.get(1))?;
    let len = s.chars().count();
    Ok(Value::Text(s.chars().skip(len.saturating_sub(n)).collect()))
}

pub fn mid(args: &[Arg]) -> Result<Value, ErrorKind> {
    let s = args[0].text()?;
    let start = args[1].number()?.trunc();
    if start < 1.0 {
        return Err(ErrorKind::Value);
    }
    let n = count_arg(Some(&args[2]))?;
    Ok(Value::Text(s.chars().skip(start as usize - 1).take(n).collect()))
}

pub fn concat(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut out = String::new();
    for v in all_values(args) {
        out.push_str(&v.as_text()?);
    }
    Ok(Value::Text(out))
}

/// TEXTJOIN(delimiter, ignore_empty, text1, ...)
pub fn textjoin(args: &[Arg]) -> Result<Value, ErrorKind> {
    let delim = args[0].text()?;
    let ignore_empty = args[1].boolean()?;
    let mut parts = Vec::new();
    for v in all_values(&args[2..]) {
        let s = v.as_text()?;
        if !(ignore_empty && s.is_empty()) {
            parts.push(s);
        }
    }
    Ok(Value::Text(parts.join(&delim)))
}

/// SUBSTITUTE(text, old, new, [instance]): case-sensitive; without an
/// instance every occurrence is replaced.
pub fn substitute(args: &[Arg]) -> Result<Value, ErrorKind> {
    let text = args[0].text()?;
    let old = args[1].text()?;
    let new = args[2].text()?;
    let instance = match args.get(3) {
        Some(a) => {
            let n = a.number()?.trunc();
            if n < 1.0 {
                return Err(ErrorKind::Value);
            }
            Some(n as usize)
        }
        None => None,
    };
    if old.is_empty() {
        return Ok(Value::Text(text));
    }
    Ok(Value::Text(match instance {
        None => text.replace(&old, &new),
        Some(n) => match text.match_indices(&old).nth(n - 1) {
            Some((at, _)) => format!("{}{}{}", &text[..at], new, &text[at + old.len()..]),
            None => text,
        },
    }))
}

/// REPLACE(old_text, start, count, new_text), positions in characters.
pub fn replace(args: &[Arg]) -> Result<Value, ErrorKind> {
    let text: Vec<char> = args[0].text()?.chars().collect();
    let start = args[1].number()?.trunc();
    if start < 1.0 {
        return Err(ErrorKind::Value);
    }
    let count = count_arg(Some(&args[2]))?;
    let new = args[3].text()?;
    let from = (start as usize - 1).min(text.len());
    let to = from.saturating_add(count).min(text.len());
    let mut out: String = text[..from].iter().collect();
    out.push_str(&new);
    out.extend(&text[to..]);
    Ok(Value::Text(out))
}

/// FIND(find_text, within_text, [start]): case-sensitive, no wildcards.
pub fn find(args: &[Arg]) -> Result<Value, ErrorKind> {
    position_of(args, |needle, hay| hay.starts_with(needle))
}

/// SEARCH(find_text, within_text, [start]): ignores case; `*`, `?` and `~` work as in criteria.
pub fn search(args: &[Arg]) -> Result<Value, ErrorKind> {
    position_of(args, |needle, hay| {
        let lower = |s: &[char]| -> String { s.iter().map(|c| c.to_lowercase().next().unwrap_or(*c)).collect() };
        wildcard_match(&(lower(needle) + "*"), &lower(hay))
    })
}

/// 1-based character position of the first match at or after `start`, else `#VALUE!`.
fn position_of(args: &[Arg], matches_at: impl Fn(&[char], &[char]) -> bool) -> Result<Value, ErrorKind> {
    let needle: Vec<char> = args[0].text()?.chars().collect();
    let hay: Vec<char> = args[1].text()?.chars().collect();
    let start = args.get(2).map_or(Ok(1.0), Arg::number)?.trunc();
    if start < 1.0 || start > hay.len() as f64 + 1.0 {
        return Err(ErrorKind::Value);
    }
    (start as usize - 1..=hay.len())
        .find(|&i| matches_at(&needle, &hay[i..]))
        .map(|i| Value::Number(i as f64 + 1.0))
        .ok_or(ErrorKind::Value)
}

/// Upper-cases every letter that follows a non-letter, lower-cases the rest.
pub fn proper(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut out = String::new();
    let mut prev_letter = false;
    for c in args[0].text()?.chars() {
        if prev_letter {
            out.extend(c.to_lowercase());
        } else {
            out.extend(c.to_uppercase());
        }
        prev_letter = c.is_alphabetic();
    }
    Ok(Value::Text(out))
}

/// Longest text a cell can hold in Excel.
const MAX_TEXT: usize = 32_767;

pub fn rept(args: &[Arg]) -> Result<Value, ErrorKind> {
    let text = args[0].text()?;
    let times = count_arg(Some(&args[1]))?;
    if text.chars().count().saturating_mul(times) > MAX_TEXT {
        return Err(ErrorKind::Value);
    }
    Ok(Value::Text(text.repeat(times)))
}

/// VALUE(text): numbers like "1,234.5", "$12", "50%", "(3)", dates and times.
pub fn value(args: &[Arg]) -> Result<Value, ErrorKind> {
    match args[0].scalar() {
        Value::Number(n) => Ok(Value::Number(n)),
        Value::Empty => Ok(Value::Number(0.0)),
        Value::Text(s) => parse_number(&s).map(Value::Number).ok_or(ErrorKind::Value),
        Value::Bool(_) => Err(ErrorKind::Value),
        Value::Error(e) => Err(e),
    }
}

fn parse_number(s: &str) -> Option<f64> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }
    if let Some(n) = datetime::parse_datetime(s) {
        return Some(n);
    }
    let (s, percent) = match s.strip_suffix('%') {
        Some(rest) => (rest.trim_end(), true),
        None => (s, false),
    };
    let (s, negative) = match s.strip_prefix('(').and_then(|r| r.strip_suffix(')')) {
        Some(inner) => (inner, true),
        None => (s, false),
    };
    let (sign, s) = match s.strip_prefix('-') {
        Some(rest) => (-1.0, rest),
        None => (1.0, s.strip_prefix('+').unwrap_or(s)),
    };
    let s = s.strip_prefix('$').unwrap_or(s);
    if s.starts_with(['-', '+']) || s.contains(char::is_whitespace) {
        return None;
    }
    // Thousands separators must sit between digits.
    if s.starts_with(',')
        || s.contains(",,")
        || s.contains(",.")
        || s.split('.').next().is_some_and(|i| i.ends_with(','))
    {
        return None;
    }
    let n: f64 = s.replace(',', "").parse().ok().filter(|n: &f64| n.is_finite())?;
    let n = sign * n * if negative { -1.0 } else { 1.0 };
    Some(if percent { n / 100.0 } else { n })
}

pub fn exact(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Bool(args[0].text()? == args[1].text()?))
}

/// TEXT(value, format). Text that is not a number is returned unchanged.
pub fn text(args: &[Arg]) -> Result<Value, ErrorKind> {
    let code = args[1].text()?;
    let x = match args[0].scalar() {
        Value::Text(s) => match parse_number(&s) {
            Some(n) => n,
            None => return Ok(Value::Text(s)),
        },
        v => v.as_number()?,
    };
    format_value(x, &code).map(Value::Text)
}

/// Windows-1252 characters for codes 128-159 (0 where the code page has a gap).
const CP1252_HIGH: [u16; 32] = [
    0x20AC, 0x81, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8D, 0x017D,
    0x8F, 0x90, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x9D,
    0x017E, 0x0178,
];

/// CHAR(number) using the Windows-1252 character set, as Excel on Windows does.
pub fn char(args: &[Arg]) -> Result<Value, ErrorKind> {
    let code = args[0].number()?.trunc();
    if !(1.0..=255.0).contains(&code) {
        return Err(ErrorKind::Value);
    }
    let code = code as u32;
    let unicode = if (128..160).contains(&code) { CP1252_HIGH[code as usize - 128] as u32 } else { code };
    Ok(Value::Text(char::from_u32(unicode).ok_or(ErrorKind::Value)?.to_string()))
}

/// CODE(text): Windows-1252 code of the first character; 63 ("?") if it has none.
pub fn code(args: &[Arg]) -> Result<Value, ErrorKind> {
    let c = args[0].text()?.chars().next().ok_or(ErrorKind::Value)? as u32;
    let code = match c {
        0..=127 | 160..=255 => c,
        _ => CP1252_HIGH.iter().position(|&u| u as u32 == c).map_or(63, |i| i as u32 + 128),
    };
    Ok(Value::Number(code as f64))
}

/// Removes the non-printable ASCII characters 0-31.
pub fn clean(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Text(args[0].text()?.chars().filter(|&c| c as u32 >= 32).collect()))
}

pub fn t(args: &[Arg]) -> Result<Value, ErrorKind> {
    match args[0].scalar() {
        Value::Text(s) => Ok(Value::Text(s)),
        Value::Error(e) => Err(e),
        _ => Ok(Value::Text(String::new())),
    }
}

pub fn n(args: &[Arg]) -> Result<Value, ErrorKind> {
    match args[0].scalar() {
        Value::Number(n) => Ok(Value::Number(n)),
        Value::Bool(b) => Ok(Value::Number(if b { 1.0 } else { 0.0 })),
        Value::Error(e) => Err(e),
        _ => Ok(Value::Number(0.0)),
    }
}

/// Optional character count argument, default 1, must not be negative.
fn count_arg(arg: Option<&Arg>) -> Result<usize, ErrorKind> {
    let n = arg.map_or(Ok(1.0), Arg::number)?.trunc();
    if n < 0.0 {
        return Err(ErrorKind::Value);
    }
    Ok(n as usize)
}

#[cfg(test)]
mod tests {
    use crate::functions::test_util::{n, run, t};
    use crate::{ErrorKind, Value};

    #[test]
    fn text_functions() {
        assert_eq!(run("LEN(\"नमस्ते\")"), n(6.0));
        assert_eq!(run("UPPER(\"abc\")"), t("ABC"));
        assert_eq!(run("LOWER(\"ABC\")"), t("abc"));
        assert_eq!(run("TRIM(\"  a   b \")"), t("a b"));
        assert_eq!(run("LEFT(\"hello\", 2)"), t("he"));
        assert_eq!(run("LEFT(\"hello\")"), t("h"));
        assert_eq!(run("RIGHT(\"hello\", 3)"), t("llo"));
        assert_eq!(run("MID(\"hello\", 2, 3)"), t("ell"));
        assert_eq!(run("MID(\"hello\", 0, 3)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("CONCAT(B1:B3, A1)"), t("abc1"));
        assert_eq!(run("CONCATENATE(\"a\", 1, TRUE)"), t("a1TRUE"));
        assert_eq!(run("TEXTJOIN(\", \", TRUE, B1:B2, D1, \"z\")"), t("a, b, z"));
    }

    #[test]
    fn substitute_and_replace() {
        assert_eq!(run("SUBSTITUTE(\"a-b-c\", \"-\", \"+\")"), t("a+b+c"));
        assert_eq!(run("SUBSTITUTE(\"a-b-c\", \"-\", \"+\", 2)"), t("a-b+c"));
        assert_eq!(run("SUBSTITUTE(\"a-b-c\", \"-\", \"+\", 3)"), t("a-b-c"));
        assert_eq!(run("SUBSTITUTE(\"Aa\", \"a\", \"x\")"), t("Ax"));
        assert_eq!(run("SUBSTITUTE(\"abc\", \"\", \"x\")"), t("abc"));
        assert_eq!(run("SUBSTITUTE(\"abc\", \"b\", \"x\", 0)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("REPLACE(\"abcdef\", 2, 3, \"XY\")"), t("aXYef"));
        assert_eq!(run("REPLACE(\"abc\", 10, 1, \"!\")"), t("abc!"));
        assert_eq!(run("REPLACE(\"abc\", 0, 1, \"!\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("REPT(\"ab\", 3)"), t("ababab"));
        assert_eq!(run("REPT(\"ab\", 0)"), t(""));
        assert_eq!(run("REPT(\"ab\", -1)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("REPT(\"ab\", 20000)"), Value::Error(ErrorKind::Value));
    }

    #[test]
    fn find_and_search() {
        assert_eq!(run("FIND(\"b\", \"abcb\")"), n(2.0));
        assert_eq!(run("FIND(\"b\", \"abcb\", 3)"), n(4.0));
        assert_eq!(run("FIND(\"B\", \"abc\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("FIND(\"\", \"abc\", 2)"), n(2.0));
        assert_eq!(run("FIND(\"a\", \"abc\", 5)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("FIND(\"स्\", \"नमस्ते\")"), n(3.0));
        assert_eq!(run("SEARCH(\"B\", \"abc\")"), n(2.0));
        assert_eq!(run("SEARCH(\"c?e\", \"abcdef\")"), n(3.0));
        assert_eq!(run("SEARCH(\"b*e\", \"abcdef\")"), n(2.0));
        assert_eq!(run("SEARCH(\"~*\", \"a*b\")"), n(2.0));
        assert_eq!(run("SEARCH(\"z\", \"abc\")"), Value::Error(ErrorKind::Value));
    }

    #[test]
    fn case_and_comparison() {
        assert_eq!(run("PROPER(\"hello wORLD 2-way it's\")"), t("Hello World 2-Way It'S"));
        assert_eq!(run("EXACT(\"abc\", \"abc\")"), Value::Bool(true));
        assert_eq!(run("EXACT(\"abc\", \"ABC\")"), Value::Bool(false));
        assert_eq!(run("EXACT(1, \"1\")"), Value::Bool(true));
        assert_eq!(run("EXACT(1/0, 1)"), Value::Error(ErrorKind::Div0));
    }

    #[test]
    fn conversions() {
        assert_eq!(run("VALUE(\"1,234.5\")"), n(1234.5));
        assert_eq!(run("VALUE(\" 42 \")"), n(42.0));
        assert_eq!(run("VALUE(\"50%\")"), n(0.5));
        assert_eq!(run("VALUE(\"$12\")"), n(12.0));
        assert_eq!(run("VALUE(\"-$1,000\")"), n(-1000.0));
        assert_eq!(run("VALUE(\"(3)\")"), n(-3.0));
        assert_eq!(run("VALUE(\"2025-01-01\")"), n(45658.0));
        assert_eq!(run("VALUE(\"18:00\")"), n(0.75));
        assert_eq!(run("VALUE(A3)"), n(3.0));
        assert_eq!(run("VALUE(\"abc\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("VALUE(\"1,,2\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("VALUE(\"\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("CHAR(65)"), t("A"));
        assert_eq!(run("CHAR(128)"), t("€"));
        assert_eq!(run("CHAR(0)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("CHAR(256)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("CODE(\"A\")"), n(65.0));
        assert_eq!(run("CODE(\"€uro\")"), n(128.0));
        assert_eq!(run("CODE(\"\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("CODE(CHAR(200))"), n(200.0));
        assert_eq!(run("CLEAN(\"a\" & CHAR(9) & \"b\" & CHAR(10))"), t("ab"));
        assert_eq!(run("T(B1)"), t("a"));
        assert_eq!(run("T(A1)"), t(""));
        assert_eq!(run("T(1/0)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("N(A2)"), n(2.0));
        assert_eq!(run("N(C1)"), n(1.0));
        assert_eq!(run("N(B1)"), n(0.0));
        assert_eq!(run("N(NA())"), Value::Error(ErrorKind::NA));
    }

    #[test]
    fn text_format() {
        assert_eq!(run("TEXT(1234.567, \"0\")"), t("1235"));
        assert_eq!(run("TEXT(1234.567, \"0.00\")"), t("1234.57"));
        assert_eq!(run("TEXT(1234.567, \"#,##0\")"), t("1,235"));
        assert_eq!(run("TEXT(1234.567, \"#,##0.00\")"), t("1,234.57"));
        assert_eq!(run("TEXT(0.256, \"0%\")"), t("26%"));
        assert_eq!(run("TEXT(0.256, \"0.00%\")"), t("25.60%"));
        assert_eq!(run("TEXT(45658, \"dd/mm/yyyy\")"), t("01/01/2025"));
        assert_eq!(run("TEXT(DATE(2025, 3, 9), \"mm/dd/yyyy\")"), t("03/09/2025"));
        assert_eq!(run("TEXT(DATE(2025, 3, 9), \"yyyy-mm-dd\")"), t("2025-03-09"));
        assert_eq!(run("TEXT(DATE(2025, 3, 9), \"dd-mmm-yyyy\")"), t("09-Mar-2025"));
        assert_eq!(run("TEXT(DATE(2025, 3, 9), \"mmmm\")"), t("March"));
        assert_eq!(run("TEXT(DATE(2025, 3, 9), \"dddd\")"), t("Sunday"));
        assert_eq!(run("TEXT(\"12.5\", \"0.00\")"), t("12.50"));
        assert_eq!(run("TEXT(\"abc\", \"0.00\")"), t("abc"));
        assert_eq!(run("TEXT(-1, \"yyyy\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("TEXT(1/0, \"0\")"), Value::Error(ErrorKind::Div0));
    }
}
