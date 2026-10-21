//! Fill-handle series: what to put in the cells you drag over.
//!
//! Rules (following Excel's AutoFill):
//! - Numbers: two or more continue their linear trend (1, 2, 3 -> 4, 5);
//!   a single number is copied unless `increment_single` (Ctrl-drag, or a date).
//! - Text ending in a number ("Item 1", "Q1"): the number counts up.
//! - Day and month names (Mon, Monday, Jan, January): continue the cycle.
//! - Anything else, including formulas: the pattern repeats, and formulas
//!   have their relative references shifted by the distance moved.

use crate::value::{format_number, Value};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FillDirection {
    Down,
    Up,
    Right,
    Left,
}

const LISTS: &[&[&str]] = &[
    &["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
    &["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
    &["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    &[
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
    ],
];

/// `source` is in natural order (top-to-bottom / left-to-right). Returns
/// `count` inputs ordered by distance from the source (nearest first).
/// `shift(input, delta)` moves a formula by `delta` cells along the fill axis.
pub fn series(
    source: &[String],
    count: usize,
    backward: bool,
    increment_single: bool,
    shift: &dyn Fn(&str, i64) -> String,
) -> Vec<String> {
    if source.is_empty() || count == 0 {
        return Vec::new();
    }
    // Work as if filling forward: for up/left fills, reverse the source.
    let seq: Vec<&str> = if backward {
        source.iter().rev().map(String::as_str).collect()
    } else {
        source.iter().map(String::as_str).collect()
    };
    let n = seq.len();
    let has_formula = seq.iter().any(|s| s.starts_with('='));

    if !has_formula {
        if let Some(nums) = seq.iter().map(|s| number(s)).collect::<Option<Vec<f64>>>() {
            if n > 1 || increment_single {
                let step_if_single = if backward { -1.0 } else { 1.0 };
                let (a, b) = if n == 1 { (nums[0], step_if_single) } else { linear_fit(&nums) };
                return (0..count).map(|i| format_number(a + b * (n + i) as f64)).collect();
            }
        }
        // With one source cell there is no trend, so the direction sets the step.
        let single_step = if backward { -1 } else { 1 };
        if let Some(out) = numbered_text(&seq, count, single_step) {
            return out;
        }
        if let Some(out) = named_list(&seq, count, single_step) {
            return out;
        }
    }

    let sign = if backward { -1 } else { 1 };
    (0..count)
        .map(|i| {
            let src = seq[i % n];
            let delta = ((i / n + 1) * n) as i64 * sign;
            if src.starts_with('=') {
                shift(src, delta)
            } else {
                src.to_string()
            }
        })
        .collect()
}

fn number(input: &str) -> Option<f64> {
    match Value::from_input(input) {
        Value::Number(n) if !input.starts_with('\'') => Some(n),
        _ => None,
    }
}

/// Least-squares line through (0, y0), (1, y1)...: returns (intercept, slope).
fn linear_fit(ys: &[f64]) -> (f64, f64) {
    let n = ys.len() as f64;
    let mean_x = (n - 1.0) / 2.0;
    let mean_y = ys.iter().sum::<f64>() / n;
    let (mut num, mut den) = (0.0, 0.0);
    for (i, y) in ys.iter().enumerate() {
        let dx = i as f64 - mean_x;
        num += dx * (y - mean_y);
        den += dx * dx;
    }
    let slope = if den == 0.0 { 0.0 } else { num / den };
    (mean_y - slope * mean_x, slope)
}

/// "Item 7" -> ("Item ", 7, ""), "Q1-2025"? -> no (number must be last or first).
fn split_number(s: &str) -> Option<(&str, i64, &str)> {
    let end = s.rfind(|c: char| c.is_ascii_digit())? + 1;
    let start = s[..end].rfind(|c: char| !c.is_ascii_digit()).map_or(0, |i| i + 1);
    let (prefix, suffix) = (&s[..start], &s[end..]);
    if prefix.is_empty() && suffix.is_empty() {
        return None; // a plain number, handled elsewhere
    }
    Some((prefix, s[start..end].parse().ok()?, suffix))
}

fn numbered_text(seq: &[&str], count: usize, single_step: i64) -> Option<Vec<String>> {
    let parts: Vec<_> = seq.iter().map(|s| split_number(s)).collect::<Option<_>>()?;
    let (prefix, _, suffix) = parts[0];
    if !parts.iter().all(|(p, _, s)| *p == prefix && *s == suffix) {
        return None;
    }
    let first = parts[0].1;
    let last = parts[parts.len() - 1].1;
    let step =
        if parts.len() == 1 { single_step } else { ((last - first) as f64 / (parts.len() - 1) as f64).round() as i64 };
    Some((1..=count as i64).map(|i| format!("{prefix}{}{suffix}", (last + step * i).abs())).collect())
}

fn named_list(seq: &[&str], count: usize, single_step: i64) -> Option<Vec<String>> {
    let list = LISTS.iter().find(|list| seq.iter().all(|s| list.iter().any(|item| item.eq_ignore_ascii_case(s))))?;
    let index = |s: &str| list.iter().position(|item| item.eq_ignore_ascii_case(s)).unwrap() as i64;
    let len = list.len() as i64;
    let last = index(seq[seq.len() - 1]);
    let step = if seq.len() == 1 {
        single_step
    } else {
        (last - index(seq[0])).rem_euclid(len) / (seq.len() as i64 - 1).max(1)
    };
    let step = if step == 0 { 1 } else { step };
    let sample = seq[seq.len() - 1];
    let case = |s: &str| {
        if sample.chars().all(|c| c.is_uppercase()) {
            s.to_uppercase()
        } else if sample.chars().all(|c| c.is_lowercase()) {
            s.to_lowercase()
        } else {
            s.to_string()
        }
    };
    Some((1..=count as i64).map(|i| case(list[(last + step * i).rem_euclid(len) as usize])).collect())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fill(src: &[&str], count: usize, backward: bool) -> Vec<String> {
        let src: Vec<String> = src.iter().map(|s| s.to_string()).collect();
        series(&src, count, backward, false, &|s, d| format!("{s}@{d}"))
    }

    #[test]
    fn numbers() {
        assert_eq!(fill(&["1", "2", "3"], 3, false), ["4", "5", "6"]);
        assert_eq!(fill(&["10", "20"], 2, false), ["30", "40"]);
        assert_eq!(fill(&["1", "2", "3"], 2, true), ["0", "-1"]);
        assert_eq!(fill(&["7"], 2, false), ["7", "7"]);
        let src = vec!["7".to_string()];
        assert_eq!(series(&src, 2, false, true, &|s, _| s.to_string()), ["8", "9"]);
        assert_eq!(fill(&["0.1", "0.2"], 1, false), ["0.3"]);
    }

    #[test]
    fn text_and_lists() {
        assert_eq!(fill(&["Item 1"], 2, false), ["Item 2", "Item 3"]);
        assert_eq!(fill(&["Q1", "Q3"], 1, false), ["Q5"]);
        assert_eq!(fill(&["Mon"], 3, false), ["Tue", "Wed", "Thu"]);
        assert_eq!(fill(&["SATURDAY"], 1, false), ["SUNDAY"]);
        assert_eq!(fill(&["Jan", "Mar"], 2, false), ["May", "Jul"]);
        assert_eq!(fill(&["Dec"], 1, true), ["Nov"]);
        assert_eq!(fill(&["a", "b"], 3, false), ["a", "b", "a"]);
    }

    #[test]
    fn formulas_repeat_with_shift() {
        assert_eq!(fill(&["=A1", "x"], 3, false), ["=A1@2", "x", "=A1@4"]);
        assert_eq!(fill(&["=A5"], 2, true), ["=A5@-1", "=A5@-2"]);
    }
}
