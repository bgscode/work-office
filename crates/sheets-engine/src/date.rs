//! Excel date serial numbers (1900 date system).
//!
//! A date is stored as a plain number: days since 1899-12-30, so 1 is
//! 1900-01-01 and 45658 is 2025-01-01. The fraction is the time of day.
//! Excel wrongly treats 1900 as a leap year; serial 60 is the fictional
//! 1900-02-29, and we reproduce that so files match Excel exactly.

/// Days since 1970-01-01 for a proleptic Gregorian date (Howard Hinnant's algorithm).
fn days_from_civil(y: i64, m: u32, d: u32) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = (m as i64 + 9) % 12;
    let doy = (153 * mp + 2) / 5 + d as i64 - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    (if m <= 2 { y + 1 } else { y }, m, d)
}

const EPOCH: i64 = -25569; // days_from_civil(1899, 12, 30)

/// Serial number for a date. Months and days may overflow (month 13 = January
/// next year, day 0 = last day of previous month), as in Excel's DATE().
pub fn serial_from_ymd(year: i64, month: i64, day: i64) -> f64 {
    let total_months = year * 12 + (month - 1);
    let (y, m) = (total_months.div_euclid(12), total_months.rem_euclid(12) + 1);
    let days = days_from_civil(y, m as u32, 1) + day - 1;
    let serial = days - EPOCH;
    // Dates before 1900-03-01 are one lower in Excel because of the fake leap day.
    (if serial < 61 { serial - 1 } else { serial }) as f64
}

/// (year, month, day) for a serial number; the time fraction is ignored.
pub fn ymd_from_serial(serial: f64) -> (i64, u32, u32) {
    let s = serial.floor() as i64;
    if s == 60 {
        return (1900, 2, 29);
    }
    let days = if s < 60 { s + 1 } else { s } + EPOCH;
    civil_from_days(days)
}

/// Days since 1899-12-30 for a Unix timestamp in milliseconds.
pub fn serial_from_unix_ms(ms: f64) -> f64 {
    ms / 86_400_000.0 + 25569.0
}

static CLOCK: std::sync::OnceLock<fn() -> f64> = std::sync::OnceLock::new();

/// Install the source of the current *local* time, in milliseconds since the
/// Unix epoch. The WebAssembly build installs one backed by JavaScript's `Date`,
/// because `std::time` is unavailable there.
pub fn set_clock(local_unix_ms: fn() -> f64) {
    let _ = CLOCK.set(local_unix_ms);
}

/// The current date and time as a serial number (for TODAY and NOW).
pub fn now_serial() -> f64 {
    let ms = match CLOCK.get() {
        Some(clock) => clock(),
        None => system_unix_ms(),
    };
    serial_from_unix_ms(ms)
}

#[cfg(not(target_arch = "wasm32"))]
fn system_unix_ms() -> f64 {
    std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map_or(0.0, |d| d.as_millis() as f64)
}

#[cfg(target_arch = "wasm32")]
fn system_unix_ms() -> f64 {
    0.0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matches_excel_serials() {
        assert_eq!(serial_from_ymd(1900, 1, 1), 1.0);
        assert_eq!(serial_from_ymd(1900, 2, 28), 59.0);
        assert_eq!(serial_from_ymd(1900, 3, 1), 61.0);
        assert_eq!(serial_from_ymd(2025, 1, 1), 45658.0);
        assert_eq!(serial_from_ymd(2024, 14, 1), serial_from_ymd(2025, 2, 1));
        assert_eq!(serial_from_ymd(2024, 3, 0), serial_from_ymd(2024, 2, 29));
        for s in [1.0, 59.0, 60.0, 61.0, 45658.0, 2958465.0] {
            let (y, m, d) = ymd_from_serial(s);
            if s != 60.0 {
                assert_eq!(serial_from_ymd(y, m as i64, d as i64), s);
            }
        }
        assert_eq!(ymd_from_serial(45658.75), (2025, 1, 1));
    }
}
