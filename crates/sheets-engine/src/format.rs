//! Turning values into display text (and typed text into numbers) according
//! to a cell's number format.

use crate::date;
use crate::style::NumberFormat;
use crate::value::{format_number, Value};

/// Display text for a value under a number format.
pub fn format_value(v: &Value, nf: &NumberFormat) -> String {
    let Value::Number(n) = v else { return v.display() };
    let n = *n;
    match nf {
        NumberFormat::General | NumberFormat::Text => format_number(n),
        NumberFormat::Number { decimals, thousands } => {
            fixed(n, *decimals, if *thousands { Group::Thousands } else { Group::None })
        }
        NumberFormat::Currency { symbol, decimals } => {
            let group = if symbol == "₹" { Group::Indian } else { Group::Thousands };
            let body = fixed(n.abs(), *decimals, group);
            if n < 0.0 && body.chars().any(|c| c.is_ascii_digit() && c != '0') {
                format!("-{symbol}{body}")
            } else {
                format!("{symbol}{body}")
            }
        }
        NumberFormat::Percent { decimals } => format!("{}%", fixed(n * 100.0, *decimals, Group::None)),
        NumberFormat::Date => format_date(n),
        NumberFormat::Time => format_time(n),
        NumberFormat::DateTime => format!("{} {}", format_date(n), format_time(n)),
    }
}

#[derive(Clone, Copy)]
enum Group {
    None,
    Thousands,
    /// 12,34,567: last three digits, then groups of two.
    Indian,
}

fn fixed(n: f64, decimals: u8, group: Group) -> String {
    let s = format!("{:.*}", decimals as usize, n);
    let (sign, s) = match s.strip_prefix('-') {
        Some(rest) if rest.chars().any(|c| c.is_ascii_digit() && c != '0') => ("-", rest),
        Some(rest) => ("", rest),
        None => ("", s.as_str()),
    };
    let (int, frac) = s.split_once('.').map_or((s, None), |(i, f)| (i, Some(f)));
    let grouped = match group {
        Group::None => int.to_string(),
        Group::Thousands => group_digits(int, 3, 3),
        Group::Indian => group_digits(int, 3, 2),
    };
    match frac {
        Some(f) => format!("{sign}{grouped}.{f}"),
        None => format!("{sign}{grouped}"),
    }
}

fn group_digits(int: &str, first: usize, rest: usize) -> String {
    if int.len() <= first {
        return int.to_string();
    }
    let (head, tail) = int.split_at(int.len() - first);
    let mut groups = vec![tail.to_string()];
    let mut head = head;
    while head.len() > rest {
        let (h, t) = head.split_at(head.len() - rest);
        groups.push(t.to_string());
        head = h;
    }
    groups.push(head.to_string());
    groups.reverse();
    groups.join(",")
}

fn format_date(serial: f64) -> String {
    if !(0.0..2_958_466.0).contains(&serial) {
        return "#".repeat(8);
    }
    let (y, m, d) = date::ymd_from_serial(serial);
    format!("{d:02}-{m:02}-{y:04}")
}

fn format_time(serial: f64) -> String {
    let minutes = ((serial.fract().abs() * 1440.0).round() as i64) % 1440;
    format!("{:02}:{:02}", minutes / 60, minutes % 60)
}

/// Recognise typed input like `12%`, `₹1,500`, `1,234.5`, `15-08-2025` or
/// `2025-08-15`, returning the number and the format it implies.
pub fn parse_formatted(input: &str) -> Option<(f64, NumberFormat)> {
    let s = input.trim();
    if s.is_empty() {
        return None;
    }
    if let Some(num) = s.strip_suffix('%') {
        let n: f64 = num.trim().parse().ok()?;
        let decimals = num.split_once('.').map_or(0, |(_, f)| f.len()) as u8;
        return Some((n / 100.0, NumberFormat::Percent { decimals }));
    }
    for symbol in ["₹", "$", "€", "£", "Rs."] {
        let (neg, rest) = match s.strip_prefix('-') {
            Some(r) => (true, r.trim_start()),
            None => (false, s),
        };
        if let Some(num) = rest.strip_prefix(symbol) {
            let n = parse_grouped(num.trim())?;
            let decimals = num.split_once('.').map_or(0, |(_, f)| f.len()) as u8;
            let symbol = if symbol == "Rs." { "₹" } else { symbol };
            return Some((if neg { -n } else { n }, NumberFormat::Currency { symbol: symbol.into(), decimals }));
        }
    }
    if s.contains(',') {
        if let Some(n) = parse_grouped(s) {
            let decimals = s.split_once('.').map_or(0, |(_, f)| f.len()) as u8;
            return Some((n, NumberFormat::Number { decimals, thousands: true }));
        }
    }
    parse_date(s).map(|serial| (serial, NumberFormat::Date))
}

/// "1,23,456.5" -> 123456.5. Every comma must be followed by digits.
fn parse_grouped(s: &str) -> Option<f64> {
    let (int, frac) = s.split_once('.').map_or((s, None), |(i, f)| (i, Some(f)));
    let int_part = int.strip_prefix('-').unwrap_or(int);
    if int_part.is_empty()
        || int_part.starts_with(',')
        || int_part.ends_with(',')
        || int_part.contains(",,")
        || !int_part.chars().all(|c| c.is_ascii_digit() || c == ',')
    {
        return None;
    }
    let cleaned = format!("{}{}", int.replace(',', ""), frac.map_or(String::new(), |f| format!(".{f}")));
    cleaned.parse().ok()
}

/// `dd-mm-yyyy`, `dd/mm/yyyy` (Indian/British order) or ISO `yyyy-mm-dd`.
fn parse_date(s: &str) -> Option<f64> {
    let parts: Vec<&str> = s.split(['-', '/', '.']).collect();
    if parts.len() != 3 || !parts.iter().all(|p| !p.is_empty() && p.chars().all(|c| c.is_ascii_digit())) {
        return None;
    }
    let nums: Vec<i64> = parts.iter().map(|p| p.parse().ok()).collect::<Option<_>>()?;
    let (y, m, d) = if parts[0].len() == 4 { (nums[0], nums[1], nums[2]) } else { (nums[2], nums[1], nums[0]) };
    let y = if parts[2].len() == 2 && parts[0].len() != 4 { 2000 + y } else { y };
    if !(1..=12).contains(&m) || !(1..=31).contains(&d) || !(1900..=9999).contains(&y) {
        return None;
    }
    let serial = date::serial_from_ymd(y, m, d);
    // Reject 31-02-2025 and similar, which DATE() would roll over.
    let (yy, mm, dd) = date::ymd_from_serial(serial);
    (yy == y && mm as i64 == m && dd as i64 == d).then_some(serial)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fmt(n: f64, nf: NumberFormat) -> String {
        format_value(&Value::Number(n), &nf)
    }

    #[test]
    fn formats_numbers() {
        assert_eq!(fmt(1234567.891, NumberFormat::Number { decimals: 2, thousands: true }), "1,234,567.89");
        assert_eq!(fmt(-0.004, NumberFormat::Number { decimals: 2, thousands: false }), "0.00");
        assert_eq!(fmt(1234567.0, NumberFormat::Currency { symbol: "₹".into(), decimals: 2 }), "₹12,34,567.00");
        assert_eq!(fmt(-1500.0, NumberFormat::Currency { symbol: "$".into(), decimals: 0 }), "-$1,500");
        assert_eq!(fmt(0.125, NumberFormat::Percent { decimals: 1 }), "12.5%");
        assert_eq!(fmt(45658.5, NumberFormat::Date), "01-01-2025");
        assert_eq!(fmt(45658.5, NumberFormat::Time), "12:00");
        assert_eq!(format_value(&Value::Text("x".into()), &NumberFormat::Date), "x");
    }

    #[test]
    fn parses_typed_formats() {
        assert_eq!(parse_formatted("12%"), Some((0.12, NumberFormat::Percent { decimals: 0 })));
        assert_eq!(
            parse_formatted("₹1,500.50"),
            Some((1500.5, NumberFormat::Currency { symbol: "₹".into(), decimals: 2 }))
        );
        assert_eq!(
            parse_formatted("1,23,456"),
            Some((123456.0, NumberFormat::Number { decimals: 0, thousands: true }))
        );
        assert_eq!(parse_formatted("01-01-2025"), Some((45658.0, NumberFormat::Date)));
        assert_eq!(parse_formatted("2025-01-01"), Some((45658.0, NumberFormat::Date)));
        assert_eq!(parse_formatted("31-02-2025"), None);
        assert_eq!(parse_formatted("1,,2"), None);
        assert_eq!(parse_formatted("hello"), None);
        assert_eq!(parse_formatted("A1,B2"), None);
    }
}
