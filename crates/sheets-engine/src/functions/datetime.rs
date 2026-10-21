//! Date and time functions. Dates are serial numbers in the 1900 date system;
//! all calendar math goes through [`crate::date`].

use super::{for_each_number, Arg};
use crate::date::{now_serial, serial_from_ymd, ymd_from_serial};
use crate::value::{ErrorKind, Value};

/// 9999-12-31, the last date Excel accepts.
pub(super) const MAX_SERIAL: f64 = 2_958_465.0;

/// (year, month, day), with Excel's serial 0 reading as 1900-01-00.
fn ymd(serial: f64) -> (i64, i64, i64) {
    if serial < 1.0 {
        return (1900, 1, 0);
    }
    let (y, m, d) = ymd_from_serial(serial);
    (y, m as i64, d as i64)
}

fn days_in_month(year: i64, month: i64) -> i64 {
    (serial_from_ymd(year, month + 1, 1) - serial_from_ymd(year, month, 1)) as i64
}

fn checked(serial: f64) -> Result<Value, ErrorKind> {
    if (0.0..MAX_SERIAL + 1.0).contains(&serial) {
        Ok(Value::Number(serial))
    } else {
        Err(ErrorKind::Num)
    }
}

/// A date/time argument: a serial number, or text such as "2025-01-31",
/// "31/01/2025", "14:30" or "2025-01-31 14:30". Out of range is `#NUM!`.
pub(super) fn serial_arg(arg: &Arg) -> Result<f64, ErrorKind> {
    let v = arg.scalar();
    let n = match (&v, v.as_number()) {
        (_, Ok(n)) => n,
        (Value::Text(s), Err(e)) => parse_datetime(s).ok_or(e)?,
        (_, Err(e)) => return Err(e),
    };
    if !(0.0..MAX_SERIAL + 1.0).contains(&n) {
        return Err(ErrorKind::Num);
    }
    Ok(n)
}

/// Parse "yyyy-mm-dd" or "dd/mm/yyyy" to a whole serial number.
pub(super) fn parse_date(s: &str) -> Option<f64> {
    let s = s.trim();
    let (sep, iso) = if s.contains('-') { ('-', true) } else { ('/', false) };
    let parts: Vec<&str> = s.split(sep).collect();
    if parts.len() != 3 || parts.iter().any(|p| p.is_empty() || p.len() > 4 || !p.bytes().all(|b| b.is_ascii_digit())) {
        return None;
    }
    let (y, m, d) = if iso { (parts[0], parts[1], parts[2]) } else { (parts[2], parts[1], parts[0]) };
    if y.len() != 4 {
        return None;
    }
    let (y, m, d): (i64, i64, i64) = (y.parse().ok()?, m.parse().ok()?, d.parse().ok()?);
    if y < 1900 || !(1..=12).contains(&m) || d < 1 || d > days_in_month(y, m) {
        return None;
    }
    Some(serial_from_ymd(y, m, d))
}

/// Parse "hh:mm", "hh:mm:ss", optionally followed by AM/PM, to a day fraction.
pub(super) fn parse_time(s: &str) -> Option<f64> {
    let s = s.trim();
    let upper = s.to_ascii_uppercase();
    let (clock, pm) = match upper.strip_suffix("AM").or_else(|| upper.strip_suffix("PM")) {
        Some(rest) => (rest.trim_end().to_string(), Some(upper.ends_with("PM"))),
        None => (upper.clone(), None),
    };
    let parts: Vec<&str> = clock.split(':').collect();
    if !(2..=3).contains(&parts.len()) || parts.iter().any(|p| p.is_empty() || !p.bytes().all(|b| b.is_ascii_digit())) {
        return None;
    }
    let mut h: u32 = parts[0].parse().ok()?;
    let m: u32 = parts[1].parse().ok()?;
    let sec: u32 = parts.get(2).map_or(Some(0), |p| p.parse().ok())?;
    if m > 59 || sec > 59 {
        return None;
    }
    match pm {
        Some(pm) => {
            if !(1..=12).contains(&h) {
                return None;
            }
            h = h % 12 + if pm { 12 } else { 0 };
        }
        None if h > 23 => return None,
        None => {}
    }
    Some((h * 3600 + m * 60 + sec) as f64 / 86_400.0)
}

/// A date, a time, or a date followed by a space (or `T`) and a time.
pub(super) fn parse_datetime(s: &str) -> Option<f64> {
    let s = s.trim();
    if let Some(n) = parse_date(s).or_else(|| parse_time(s)) {
        return Some(n);
    }
    let (date, time) = s.split_once([' ', 'T'])?;
    Some(parse_date(date)? + parse_time(time)?)
}

/// 0 = Sunday ... 6 = Saturday.
fn day_of_week(serial: f64) -> i64 {
    (serial.floor() as i64 + 6).rem_euclid(7)
}

pub fn date(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut y = args[0].number()?.trunc();
    let m = args[1].number()?.trunc();
    let d = args[2].number()?.trunc();
    if (0.0..1900.0).contains(&y) {
        y += 1900.0;
    }
    if !(1900.0..10000.0).contains(&y) || m.abs() > 1e6 || d.abs() > 1e8 {
        return Err(ErrorKind::Num);
    }
    checked(serial_from_ymd(y as i64, m as i64, d as i64))
}

pub fn today(_: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(now_serial().floor()))
}

pub fn now(_: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(now_serial()))
}

pub fn year(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(ymd(serial_arg(&args[0])?).0 as f64))
}

pub fn month(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(ymd(serial_arg(&args[0])?).1 as f64))
}

pub fn day(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(ymd(serial_arg(&args[0])?).2 as f64))
}

/// WEEKDAY(date, [type = 1]): 1 = Sun 1..Sat 7, 2 = Mon 1..Sun 7, 3 = Mon 0..Sun 6,
/// 11..17 = week starting Monday..Sunday numbered from 1.
pub fn weekday(args: &[Arg]) -> Result<Value, ErrorKind> {
    let dow = day_of_week(serial_arg(&args[0])?);
    let kind = args.get(1).map_or(Ok(1.0), Arg::number)?.trunc() as i64;
    let (first_day, base) = match kind {
        1 => (0, 1),
        2 => (1, 1),
        3 => (1, 0),
        11..=17 => ((kind - 10) % 7, 1),
        _ => return Err(ErrorKind::Num),
    };
    Ok(Value::Number(((dow - first_day).rem_euclid(7) + base) as f64))
}

/// WEEKNUM(date, [type = 1]): the week containing January 1 is week 1. Types:
/// 1 weeks start Sunday, 2 Monday, 11..17 Monday..Sunday, 21 ISO 8601.
pub fn weeknum(args: &[Arg]) -> Result<Value, ErrorKind> {
    let serial = serial_arg(&args[0])?.floor();
    let kind = args.get(1).map_or(Ok(1.0), Arg::number)?.trunc() as i64;
    let first_day = match kind {
        1 | 17 => 0,
        2 | 11 => 1,
        12..=16 => kind - 10,
        21 => return Ok(Value::Number(iso_week(serial) as f64)),
        _ => return Err(ErrorKind::Num),
    };
    let jan1 = serial_from_ymd(ymd(serial).0, 1, 1);
    let shift = (day_of_week(jan1) - first_day).rem_euclid(7);
    Ok(Value::Number(((serial - jan1) as i64 + shift).div_euclid(7) as f64 + 1.0))
}

fn iso_week(serial: f64) -> i64 {
    let monday_based = (day_of_week(serial) + 6) % 7;
    let thursday = serial - monday_based as f64 + 3.0;
    let jan1 = serial_from_ymd(ymd(thursday).0, 1, 1);
    (thursday - jan1) as i64 / 7 + 1
}

/// Same day of the month `months` away, clamped to the month's last day.
pub fn edate(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (y, m, d) = ymd(serial_arg(&args[0])?);
    let months = months_arg(&args[1])?;
    let day = d.min(days_in_month(y, m + months));
    checked(serial_from_ymd(y, m + months, day))
}

pub fn eomonth(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (y, m, _) = ymd(serial_arg(&args[0])?);
    let months = months_arg(&args[1])?;
    checked(serial_from_ymd(y, m + months + 1, 0))
}

fn months_arg(arg: &Arg) -> Result<i64, ErrorKind> {
    let months = arg.number()?.trunc();
    if months.abs() > 120_000.0 {
        return Err(ErrorKind::Num);
    }
    Ok(months as i64)
}

/// DATEDIF(start, end, unit) with unit Y, M, D, MD, YM or YD.
pub fn datedif(args: &[Arg]) -> Result<Value, ErrorKind> {
    let start = serial_arg(&args[0])?.floor();
    let end = serial_arg(&args[1])?.floor();
    let unit = args[2].text()?.to_ascii_uppercase();
    if start > end {
        return Err(ErrorKind::Num);
    }
    let (y1, m1, d1) = ymd(start);
    let (y2, m2, d2) = ymd(end);
    let months = (y2 - y1) * 12 + (m2 - m1) - (d2 < d1) as i64;
    let result = match unit.as_str() {
        "Y" => months / 12,
        "M" => months,
        "D" => (end - start) as i64,
        "YM" => months % 12,
        "MD" => {
            if d2 >= d1 {
                d2 - d1
            } else {
                // Days since the start's day-of-month in the previous month,
                // clamped to that month's length (Excel can return negatives here).
                let day = d1.min(days_in_month(y2, m2 - 1));
                (end - serial_from_ymd(y2, m2 - 1, day)) as i64
            }
        }
        "YD" => {
            let mut anniversary = serial_from_ymd(y2, m1, d1);
            if anniversary > end {
                anniversary = serial_from_ymd(y2 - 1, m1, d1);
            }
            (end - anniversary) as i64
        }
        _ => return Err(ErrorKind::Num),
    };
    Ok(Value::Number(result as f64))
}

/// DAYS(end, start)
pub fn days(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(serial_arg(&args[0])?.floor() - serial_arg(&args[1])?.floor()))
}

fn holidays(arg: Option<&Arg>) -> Result<Vec<f64>, ErrorKind> {
    let mut out = Vec::new();
    if let Some(arg) = arg {
        for_each_number(std::slice::from_ref(arg), |n| out.push(n.floor()))?;
    }
    Ok(out)
}

fn is_workday(serial: f64, holidays: &[f64]) -> bool {
    !matches!(day_of_week(serial), 0 | 6) && !holidays.contains(&serial)
}

/// Working days (Mon-Fri, not holidays) from start to end inclusive;
/// negative when end is before start.
pub fn networkdays(args: &[Arg]) -> Result<Value, ErrorKind> {
    let start = serial_arg(&args[0])?.floor();
    let end = serial_arg(&args[1])?.floor();
    let holidays = holidays(args.get(2))?;
    let (lo, hi) = if start <= end { (start as i64, end as i64) } else { (end as i64, start as i64) };
    let count = (lo..=hi).filter(|&s| is_workday(s as f64, &holidays)).count() as f64;
    Ok(Value::Number(if start <= end { count } else { -count }))
}

/// The date `days` working days before or after start (start itself is not counted).
pub fn workday(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut serial = serial_arg(&args[0])?.floor();
    let days = args[1].number()?.trunc();
    let holidays = holidays(args.get(2))?;
    if days.abs() > MAX_SERIAL {
        return Err(ErrorKind::Num);
    }
    let step = days.signum();
    let mut remaining = days.abs() as i64;
    while remaining > 0 {
        serial += step;
        if !(0.0..=MAX_SERIAL).contains(&serial) {
            return Err(ErrorKind::Num);
        }
        if is_workday(serial, &holidays) {
            remaining -= 1;
        }
    }
    Ok(Value::Number(serial))
}

/// Seconds since midnight, rounded to the nearest second like Excel.
fn seconds_of_day(arg: &Arg) -> Result<i64, ErrorKind> {
    let serial = serial_arg(arg)?;
    Ok(((serial - serial.floor()) * 86_400.0).round() as i64 % 86_400)
}

pub fn hour(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number((seconds_of_day(&args[0])? / 3600) as f64))
}

pub fn minute(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number((seconds_of_day(&args[0])? / 60 % 60) as f64))
}

pub fn second(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number((seconds_of_day(&args[0])? % 60) as f64))
}

/// TIME(hour, minute, second): a day fraction; whole days wrap around.
pub fn time(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut parts = [0.0; 3];
    for (part, arg) in parts.iter_mut().zip(args) {
        *part = arg.number()?.trunc();
        if part.abs() > 32_767.0 {
            return Err(ErrorKind::Num);
        }
    }
    let total = parts[0] * 3600.0 + parts[1] * 60.0 + parts[2];
    if total < 0.0 {
        return Err(ErrorKind::Num);
    }
    Ok(Value::Number(total % 86_400.0 / 86_400.0))
}

/// DATEVALUE("yyyy-mm-dd" or "dd/mm/yyyy"); a trailing time is ignored.
pub fn datevalue(args: &[Arg]) -> Result<Value, ErrorKind> {
    match args[0].scalar() {
        Value::Text(s) => {
            parse_datetime(&s).filter(|&n| n >= 1.0).map(|n| Value::Number(n.floor())).ok_or(ErrorKind::Value)
        }
        Value::Error(e) => Err(e),
        _ => Err(ErrorKind::Value),
    }
}

#[cfg(test)]
mod tests {
    use crate::date::serial_from_ymd;
    use crate::functions::test_util::{n, run};
    use crate::{ErrorKind, Value};

    fn d(y: i64, m: i64, day: i64) -> Value {
        n(serial_from_ymd(y, m, day))
    }

    #[test]
    fn building_dates() {
        assert_eq!(run("DATE(2025, 1, 1)"), n(45658.0));
        assert_eq!(run("DATE(2024, 14, 1)"), d(2025, 2, 1));
        assert_eq!(run("DATE(2024, 3, 0)"), d(2024, 2, 29));
        assert_eq!(run("DATE(99, 1, 1)"), d(1999, 1, 1));
        assert_eq!(run("DATE(10000, 1, 1)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("DATE(\"x\", 1, 1)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("TIME(12, 0, 0)"), n(0.5));
        assert_eq!(run("TIME(25, 0, 0)"), n(1.0 / 24.0));
        assert_eq!(run("TIME(0, -1, 0)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("DATEVALUE(\"2025-01-01\")"), n(45658.0));
        assert_eq!(run("DATEVALUE(\"31/12/2024\")"), n(45657.0));
        assert_eq!(run("DATEVALUE(\"2025-01-01 10:00\")"), n(45658.0));
        assert_eq!(run("DATEVALUE(\"2025-02-30\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("DATEVALUE(\"hello\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("DATEVALUE(45658)"), Value::Error(ErrorKind::Value));
    }

    #[test]
    fn now_and_today() {
        let Value::Number(today) = run("TODAY()") else { panic!("TODAY must return a number") };
        let Value::Number(now) = run("NOW()") else { panic!("NOW must return a number") };
        assert_eq!(today.fract(), 0.0);
        assert!(today > 45000.0 && now >= today && now < today + 1.0 + 1e-6);
    }

    #[test]
    fn date_parts() {
        assert_eq!(run("YEAR(45658)"), n(2025.0));
        assert_eq!(run("MONTH(DATE(2024, 7, 15))"), n(7.0));
        assert_eq!(run("DAY(DATE(2024, 7, 15))"), n(15.0));
        assert_eq!(run("DAY(\"2024-02-29\")"), n(29.0));
        assert_eq!(run("YEAR(0)"), n(1900.0));
        assert_eq!(run("YEAR(-1)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("MONTH(\"soon\")"), Value::Error(ErrorKind::Value));
        // 2025-01-01 is a Wednesday.
        assert_eq!(run("WEEKDAY(45658)"), n(4.0));
        assert_eq!(run("WEEKDAY(45658, 2)"), n(3.0));
        assert_eq!(run("WEEKDAY(45658, 3)"), n(2.0));
        assert_eq!(run("WEEKDAY(45658, 13)"), n(1.0));
        assert_eq!(run("WEEKDAY(45658, 4)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("WEEKNUM(DATE(2025, 1, 4))"), n(1.0));
        assert_eq!(run("WEEKNUM(DATE(2025, 1, 5))"), n(2.0));
        assert_eq!(run("WEEKNUM(DATE(2025, 1, 5), 2)"), n(1.0));
        assert_eq!(run("WEEKNUM(DATE(2025, 1, 6), 2)"), n(2.0));
        assert_eq!(run("WEEKNUM(DATE(2025, 12, 31))"), n(53.0));
        assert_eq!(run("WEEKNUM(DATE(2024, 12, 30), 21)"), n(1.0));
        assert_eq!(run("WEEKNUM(DATE(2021, 1, 3), 21)"), n(53.0));
        assert_eq!(run("WEEKNUM(1, 5)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("HOUR(0.75)"), n(18.0));
        assert_eq!(run("MINUTE(TIME(10, 45, 30))"), n(45.0));
        assert_eq!(run("SECOND(TIME(10, 45, 30))"), n(30.0));
        assert_eq!(run("HOUR(\"14:30\")"), n(14.0));
        assert_eq!(run("HOUR(\"2:30 PM\")"), n(14.0));
        assert_eq!(run("SECOND(-0.5)"), Value::Error(ErrorKind::Num));
    }

    #[test]
    fn date_arithmetic() {
        assert_eq!(run("EDATE(DATE(2024, 1, 31), 1)"), d(2024, 2, 29));
        assert_eq!(run("EDATE(DATE(2024, 3, 15), -3)"), d(2023, 12, 15));
        assert_eq!(run("EOMONTH(DATE(2024, 1, 15), 1)"), d(2024, 2, 29));
        assert_eq!(run("EOMONTH(DATE(2024, 1, 15), -1)"), d(2023, 12, 31));
        assert_eq!(run("EOMONTH(1, -2)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("DAYS(DATE(2025, 3, 1), DATE(2025, 2, 1))"), n(28.0));
        assert_eq!(run("DAYS(\"2025-01-02\", \"2025-01-01\")"), n(1.0));
        let dif = |unit: &str| run(&format!("DATEDIF(DATE(2020, 5, 20), DATE(2024, 3, 10), \"{unit}\")"));
        assert_eq!(dif("Y"), n(3.0));
        assert_eq!(dif("M"), n(45.0));
        assert_eq!(dif("D"), n(1390.0));
        assert_eq!(dif("YM"), n(9.0));
        assert_eq!(dif("MD"), n(19.0));
        assert_eq!(dif("YD"), n(295.0));
        assert_eq!(dif("ym"), n(9.0));
        assert_eq!(dif("W"), Value::Error(ErrorKind::Num));
        assert_eq!(run("DATEDIF(DATE(2024, 1, 2), DATE(2024, 1, 1), \"D\")"), Value::Error(ErrorKind::Num));
    }

    #[test]
    fn working_days() {
        // 2025-01-06 is a Monday.
        assert_eq!(run("NETWORKDAYS(DATE(2025, 1, 6), DATE(2025, 1, 19))"), n(10.0));
        assert_eq!(run("NETWORKDAYS(DATE(2025, 1, 19), DATE(2025, 1, 6))"), n(-10.0));
        assert_eq!(run("NETWORKDAYS(DATE(2025, 1, 6), DATE(2025, 1, 10), DATE(2025, 1, 8))"), n(4.0));
        assert_eq!(run("NETWORKDAYS(DATE(2025, 1, 6), DATE(2025, 1, 10), A1:A5)"), n(5.0));
        assert_eq!(run("NETWORKDAYS(\"x\", 1)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("WORKDAY(DATE(2025, 1, 10), 1)"), d(2025, 1, 13));
        assert_eq!(run("WORKDAY(DATE(2025, 1, 13), -1)"), d(2025, 1, 10));
        assert_eq!(run("WORKDAY(DATE(2025, 1, 6), 5, DATE(2025, 1, 8))"), d(2025, 1, 14));
        assert_eq!(run("WORKDAY(DATE(2025, 1, 11), 0)"), d(2025, 1, 11));
        assert_eq!(run("WORKDAY(2, -5)"), Value::Error(ErrorKind::Num));
    }
}
