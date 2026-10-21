//! A small Excel number-format formatter for `TEXT(value, format)`.
//!
//! Supported: `General`, `@`, digit placeholders `0` and `#`, decimal point,
//! thousands separator `,`, percent `%`, up to three `;` sections
//! (positive;negative;zero), quoted literals, `\x` escapes, `[color]` tags
//! (ignored), and date/time codes `d dd ddd dddd m mm mmm mmmm yy yyyy h hh
//! m mm s ss AM/PM A/P`. Not supported: scientific notation, fractions,
//! elapsed time (`[h]`), fractional seconds and conditional sections.

use crate::value::{format_number, ErrorKind};

const MONTHS: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];
const DAYS: [&str; 7] = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/// Format a number with an Excel format code.
pub fn format_value(x: f64, code: &str) -> Result<String, ErrorKind> {
    if code.is_empty() {
        return Ok(String::new());
    }
    if code.eq_ignore_ascii_case("general") || code == "@" {
        return Ok(format_number(x));
    }
    let sections = split_sections(code);
    let (section, value, explicit_sign) = match sections.len() {
        1 => (sections[0].as_str(), x, false),
        2 => {
            if x < 0.0 {
                (sections[1].as_str(), -x, true)
            } else {
                (sections[0].as_str(), x, false)
            }
        }
        _ => {
            if x < 0.0 {
                (sections[1].as_str(), -x, true)
            } else if x == 0.0 {
                (sections[2].as_str(), x, false)
            } else {
                (sections[0].as_str(), x, false)
            }
        }
    };
    let tokens = tokenize(section);
    if tokens.iter().any(|t| matches!(t, Token::Date(_))) {
        if value < 0.0 {
            return Err(ErrorKind::Value);
        }
        return Ok(format_date(value, &tokens));
    }
    Ok(format_numeric(value, &tokens, explicit_sign))
}

/// Split on `;` outside quotes and escapes.
fn split_sections(code: &str) -> Vec<String> {
    let mut sections = vec![String::new()];
    let mut chars = code.chars();
    let mut quoted = false;
    while let Some(c) = chars.next() {
        let current = sections.last_mut().unwrap();
        match c {
            '"' => {
                quoted = !quoted;
                current.push(c);
            }
            '\\' if !quoted => {
                current.push(c);
                if let Some(next) = chars.next() {
                    current.push(next);
                }
            }
            ';' if !quoted => sections.push(String::new()),
            _ => current.push(c),
        }
    }
    sections
}

#[derive(Debug, Clone, PartialEq)]
enum Token {
    Literal(String),
    /// A run of `0 # , .` characters.
    Digits(String),
    Percent,
    /// A date/time code such as `yyyy`, `mm`, `AM/PM`.
    Date(String),
}

fn tokenize(section: &str) -> Vec<Token> {
    let chars: Vec<char> = section.chars().collect();
    let mut tokens = Vec::new();
    let mut i = 0;
    let push_literal = |tokens: &mut Vec<Token>, s: &str| match tokens.last_mut() {
        Some(Token::Literal(l)) => l.push_str(s),
        _ => tokens.push(Token::Literal(s.to_string())),
    };
    while i < chars.len() {
        let c = chars[i];
        let rest: String = chars[i..].iter().collect();
        let upper_rest = rest.to_ascii_uppercase();
        match c {
            '"' => {
                let end = chars[i + 1..].iter().position(|&ch| ch == '"').map_or(chars.len(), |p| i + 1 + p);
                push_literal(&mut tokens, &chars[i + 1..end].iter().collect::<String>());
                i = end + 1;
            }
            '\\' => {
                if let Some(next) = chars.get(i + 1) {
                    push_literal(&mut tokens, &next.to_string());
                }
                i += 2;
            }
            '[' => {
                // Colour or locale tag: skipped.
                i = chars[i..].iter().position(|&ch| ch == ']').map_or(chars.len(), |p| i + p + 1);
            }
            '_' => {
                push_literal(&mut tokens, " ");
                i += 2;
            }
            '*' => i += 2,
            '%' => {
                tokens.push(Token::Percent);
                i += 1;
            }
            '0' | '#' | '?' | '.' | ',' => {
                // A `.` or `,` alone between literals stays literal, e.g. "dd.mm.yyyy".
                let start = i;
                while i < chars.len() && matches!(chars[i], '0' | '#' | '?' | '.' | ',') {
                    i += 1;
                }
                let run: String = chars[start..i].iter().collect();
                if run.contains(['0', '#', '?']) {
                    tokens.push(Token::Digits(run));
                } else {
                    push_literal(&mut tokens, &run);
                }
            }
            _ if upper_rest.starts_with("AM/PM") => {
                tokens.push(Token::Date("AM/PM".into()));
                i += 5;
            }
            _ if upper_rest.starts_with("A/P") => {
                tokens.push(Token::Date("A/P".into()));
                i += 3;
            }
            'y' | 'Y' | 'm' | 'M' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => {
                let lower = c.to_ascii_lowercase();
                let start = i;
                while i < chars.len() && chars[i].to_ascii_lowercase() == lower {
                    i += 1;
                }
                tokens.push(Token::Date(std::iter::repeat_n(lower, i - start).collect()));
            }
            _ => {
                push_literal(&mut tokens, &c.to_string());
                i += 1;
            }
        }
    }
    tokens
}

fn format_date(serial: f64, tokens: &[Token]) -> String {
    let mut day = serial.floor();
    let mut secs = ((serial - day) * 86_400.0).round() as i64;
    if secs >= 86_400 {
        day += 1.0;
        secs -= 86_400;
    }
    let (y, m, d) = if day < 1.0 { (1900, 1, 0) } else { crate::date::ymd_from_serial(day) };
    let dow = (day as i64 + 6).rem_euclid(7) as usize;
    let (hour, minute, second) = (secs / 3600, secs / 60 % 60, secs % 60);
    let twelve_hour = tokens.iter().any(|t| matches!(t, Token::Date(s) if s.contains('/')));

    // `m`/`mm` means minutes right after an hour code or right before a seconds code.
    let codes: Vec<(usize, &str)> = tokens
        .iter()
        .enumerate()
        .filter_map(|(i, t)| if let Token::Date(s) = t { Some((i, s.as_str())) } else { None })
        .collect();
    let is_minute = |pos: usize| {
        let k = codes.iter().position(|&(i, _)| i == pos).unwrap();
        let after_hour = k > 0 && codes[k - 1].1.starts_with('h');
        let before_second = codes.get(k + 1).is_some_and(|c| c.1.starts_with('s'));
        after_hour || before_second
    };

    let mut out = String::new();
    for (i, token) in tokens.iter().enumerate() {
        let Token::Date(code) = token else {
            match token {
                Token::Literal(s) | Token::Digits(s) => out.push_str(s),
                _ => out.push('%'),
            }
            continue;
        };
        let text = match code.as_str() {
            "AM/PM" => (if hour < 12 { "AM" } else { "PM" }).to_string(),
            "A/P" => (if hour < 12 { "A" } else { "P" }).to_string(),
            "yy" => format!("{:02}", y.rem_euclid(100)),
            c if c.starts_with('y') => format!("{y:04}"),
            "m" | "mm" if is_minute(i) => pad(minute, code.len()),
            "m" | "mm" => pad(m as i64, code.len()),
            "mmm" => MONTHS[m as usize - 1][..3].to_string(),
            "mmmmm" => MONTHS[m as usize - 1][..1].to_string(),
            c if c.starts_with('m') => MONTHS[m as usize - 1].to_string(),
            "d" | "dd" => pad(d as i64, code.len()),
            "ddd" => DAYS[dow][..3].to_string(),
            c if c.starts_with('d') => DAYS[dow].to_string(),
            c if c.starts_with('h') => {
                let h = if twelve_hour { (hour + 11) % 12 + 1 } else { hour };
                pad(h, code.len().min(2))
            }
            c => pad(second, c.len().min(2)),
        };
        out.push_str(&text);
    }
    out
}

fn pad(n: i64, width: usize) -> String {
    format!("{n:0width$}")
}

fn format_numeric(x: f64, tokens: &[Token], explicit_sign: bool) -> String {
    let percents = tokens.iter().filter(|t| **t == Token::Percent).count() as i32;
    let pattern: String =
        tokens.iter().filter_map(|t| if let Token::Digits(s) = t { Some(s.as_str()) } else { None }).collect();
    let (int_pat, frac_pat) = pattern.split_once('.').unwrap_or((&pattern, ""));
    let decimals = frac_pat.chars().filter(|c| matches!(c, '0' | '#' | '?')).count();
    let min_decimals = frac_pat.chars().filter(|&c| c == '0').count();
    let min_int = int_pat.chars().filter(|&c| c == '0').count();
    let grouping = int_pat.trim_end_matches(',').contains(',');
    // Commas right before the decimal point (or at the end) divide by 1000 each.
    let scale_commas = int_pat.len() - int_pat.trim_end_matches(',').len();

    let scaled = x * 100f64.powi(percents) / 1000f64.powi(scale_commas as i32);
    let negative = scaled < 0.0 && !explicit_sign;
    let digits = round_digits(scaled.abs(), decimals);
    let (int_digits, frac_digits) = digits.split_at(digits.len() - decimals);
    let mut int_part = int_digits.trim_start_matches('0').to_string();
    if int_part.len() < min_int {
        int_part = "0".repeat(min_int - int_part.len()) + &int_part;
    }
    if grouping {
        int_part = group_thousands(&int_part);
    }
    let mut frac_part = frac_digits.to_string();
    while frac_part.len() > min_decimals && frac_part.ends_with('0') {
        frac_part.pop();
    }
    let is_zero = digits.bytes().all(|b| b == b'0');

    let mut out = String::new();
    if negative && !is_zero {
        out.push('-');
    }
    let mut number_written = false;
    for token in tokens {
        match token {
            Token::Literal(s) => out.push_str(s),
            Token::Percent => out.push('%'),
            Token::Digits(_) if !number_written => {
                number_written = true;
                out.push_str(&int_part);
                if decimals > 0 && (min_decimals > 0 || !frac_part.is_empty() || pattern.ends_with('.')) {
                    out.push('.');
                    out.push_str(&frac_part);
                }
            }
            _ => {}
        }
    }
    out
}

/// Digits of |x| rounded half away from zero to `decimals` places, with the
/// decimal point removed and at least `decimals + 1` digits.
fn round_digits(x: f64, decimals: usize) -> String {
    let scaled = x * 10f64.powi(decimals as i32);
    // Trim binary noise first so 2.675 rounds to 2.68 like Excel.
    let tidy: f64 = format!("{:.14e}", scaled).parse().unwrap_or(scaled);
    let digits = format!("{:.0}", tidy.round());
    if digits.len() <= decimals {
        "0".repeat(decimals + 1 - digits.len()) + &digits
    } else {
        digits
    }
}

fn group_thousands(digits: &str) -> String {
    let mut out = String::new();
    for (i, c) in digits.chars().enumerate() {
        if i > 0 && (digits.len() - i).is_multiple_of(3) {
            out.push(',');
        }
        out.push(c);
    }
    out
}

#[cfg(test)]
mod tests {
    use super::format_value;

    fn fmt(x: f64, code: &str) -> String {
        format_value(x, code).unwrap()
    }

    #[test]
    fn numbers() {
        assert_eq!(fmt(3.25, "0"), "3");
        assert_eq!(fmt(2.5, "0"), "3");
        assert_eq!(fmt(-2.5, "0"), "-3");
        assert_eq!(fmt(2.675, "0.00"), "2.68");
        assert_eq!(fmt(0.5, "0.00"), "0.50");
        assert_eq!(fmt(1234567.891, "#,##0"), "1,234,568");
        assert_eq!(fmt(1234567.891, "#,##0.00"), "1,234,567.89");
        assert_eq!(fmt(-1234.5, "#,##0.00"), "-1,234.50");
        assert_eq!(fmt(999.999, "#,##0.00"), "1,000.00");
        assert_eq!(fmt(0.256, "0%"), "26%");
        assert_eq!(fmt(0.25678, "0.00%"), "25.68%");
        assert_eq!(fmt(0.0, "#"), "");
        assert_eq!(fmt(0.0, "0.0#"), "0.0");
        assert_eq!(fmt(1.234, "0.0#"), "1.23");
        assert_eq!(fmt(42.0, "000"), "042");
        assert_eq!(fmt(12345.0, "#,##0,\"K\""), "12K");
        assert_eq!(fmt(5.0, "\"$\"#,##0.00"), "$5.00");
        assert_eq!(fmt(-5.0, "0;(0)"), "(5)");
        assert_eq!(fmt(0.0, "0;-0;\"zero\""), "zero");
        assert_eq!(fmt(-0.001, "0.00"), "0.00");
        assert_eq!(fmt(1.5, "General"), "1.5");
        assert_eq!(fmt(7.0, "@"), "7");
        assert_eq!(fmt(7.0, "\"n/a\""), "n/a");
    }

    #[test]
    fn dates() {
        let serial = 45658.75; // 2025-01-01 18:00, a Wednesday
        assert_eq!(fmt(serial, "dd/mm/yyyy"), "01/01/2025");
        assert_eq!(fmt(serial, "mm/dd/yyyy"), "01/01/2025");
        assert_eq!(fmt(45689.0, "yyyy-mm-dd"), "2025-02-01");
        assert_eq!(fmt(45689.0, "dd-mmm-yyyy"), "01-Feb-2025");
        assert_eq!(fmt(45689.0, "mmmm"), "February");
        assert_eq!(fmt(serial, "dddd"), "Wednesday");
        assert_eq!(fmt(serial, "ddd d mmm yy"), "Wed 1 Jan 25");
        assert_eq!(fmt(serial, "hh:mm:ss"), "18:00:00");
        assert_eq!(fmt(serial, "h:mm AM/PM"), "6:00 PM");
        assert_eq!(fmt(0.5 + 1.0 / 1440.0, "h:mm"), "12:01");
        assert_eq!(fmt(45658.0, "d.m.yyyy"), "1.1.2025");
        assert!(format_value(-1.0, "yyyy").is_err());
    }
}
