use std::cmp::Ordering;

use super::conditional::wildcard_match;
use super::Arg;
use crate::value::{compare, ErrorKind, Value};

/// VLOOKUP(value, table, col_index, [approximate = TRUE])
pub fn vlookup(args: &[Arg]) -> Result<Value, ErrorKind> {
    table_lookup(args, false)
}

/// HLOOKUP(value, table, row_index, [approximate = TRUE])
pub fn hlookup(args: &[Arg]) -> Result<Value, ErrorKind> {
    table_lookup(args, true)
}

fn table_lookup(args: &[Arg], horizontal: bool) -> Result<Value, ErrorKind> {
    let needle = args[0].scalar();
    if let Value::Error(e) = needle {
        return Err(e);
    }
    let table = args[1].as_range();
    let index = args[2].number()?.trunc();
    let approximate = args.get(3).map_or(Ok(true), Arg::boolean)?;
    let (keys, width) = if horizontal { (table.cols, table.rows) } else { (table.rows, table.cols) };
    if index < 1.0 {
        return Err(ErrorKind::Value);
    }
    if index as u32 > width {
        return Err(ErrorKind::Ref);
    }
    let cell = |key: u32, offset: u32| {
        if horizontal { table.get(offset, key) } else { table.get(key, offset) }.cloned().unwrap_or_default()
    };
    let key_values: Vec<Value> = (0..keys).map(|k| cell(k, 0)).collect();
    let found = if approximate { approx_le(&needle, &key_values) } else { exact(&needle, &key_values) };
    found.map(|k| cell(k as u32, index as u32 - 1)).ok_or(ErrorKind::NA)
}

/// INDEX(range, row, [col]) — 1-based. A single row or column needs only one index.
pub fn index(args: &[Arg]) -> Result<Value, ErrorKind> {
    let range = args[0].as_range();
    let a = args[1].number()?.trunc() as i64;
    let b = args.get(2).map_or(Ok(None), |x| x.number().map(|n| Some(n.trunc() as i64)))?;
    let (row, col) = match b {
        Some(col) => (a, col),
        None if range.rows == 1 => (1, a),
        None => (a, 1),
    };
    if row < 1 || col < 1 {
        return Err(ErrorKind::Value);
    }
    range.get(row as u32 - 1, col as u32 - 1).cloned().ok_or(ErrorKind::Ref)
}

/// MATCH(value, range, [type = 1]): 1 = largest <= value (sorted ascending),
/// 0 = exact, -1 = smallest >= value (sorted descending). Returns a 1-based position.
pub fn match_(args: &[Arg]) -> Result<Value, ErrorKind> {
    let needle = args[0].scalar();
    if let Value::Error(e) = needle {
        return Err(e);
    }
    let range = args[1].as_range();
    if range.rows > 1 && range.cols > 1 {
        return Err(ErrorKind::NA);
    }
    let kind = args.get(2).map_or(Ok(1.0), Arg::number)?;
    let values = &range.values;
    let found = if kind == 0.0 {
        exact(&needle, values)
    } else if kind > 0.0 {
        approx_le(&needle, values)
    } else {
        approx_ge(&needle, values)
    };
    found.map(|i| Value::Number(i as f64 + 1.0)).ok_or(ErrorKind::NA)
}

/// XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
///
/// match_mode: 0 exact (default), -1 exact or next smaller, 1 exact or next
/// larger, 2 wildcard. search_mode: 1 first to last (default), -1 last to
/// first; 2 and -2 (binary search) are accepted and searched linearly.
/// Dynamic arrays are not supported, so a 2-D return_array yields only the
/// first value of the matched row or column.
pub fn xlookup(args: &[Arg]) -> Result<Value, ErrorKind> {
    let needle = args[0].scalar();
    if let Value::Error(e) = needle {
        return Err(e);
    }
    let keys = args[1].as_range();
    let results = args[2].as_range();
    let vertical = if keys.cols == 1 && results.rows == keys.rows {
        true
    } else if keys.rows == 1 && results.cols == keys.cols {
        false
    } else {
        return Err(ErrorKind::Value);
    };
    let match_mode = args.get(4).map_or(Ok(0.0), Arg::number)?.trunc();
    let search_mode = args.get(5).map_or(Ok(1.0), Arg::number)?.trunc();
    if ![-1.0, 0.0, 1.0, 2.0].contains(&match_mode) || ![-2.0, -1.0, 1.0, 2.0].contains(&search_mode) {
        return Err(ErrorKind::Value);
    }
    let mut order: Vec<usize> = (0..keys.values.len()).collect();
    if search_mode < 0.0 {
        order.reverse();
    }
    let (key_values, target) = (&keys.values, &needle);
    let candidates = || order.iter().copied().filter(move |&i| comparable(&key_values[i], target));
    let exact_match = || candidates().find(|&i| compare(&keys.values[i], &needle) == Ordering::Equal);
    let found = match (match_mode as i8, &needle) {
        (2, Value::Text(pattern)) => {
            let pattern = pattern.to_lowercase();
            candidates()
                .find(|&i| matches!(&keys.values[i], Value::Text(s) if wildcard_match(&pattern, &s.to_lowercase())))
        }
        (-1, _) => exact_match().or_else(|| {
            let smaller = candidates().filter(|&i| compare(&keys.values[i], &needle) == Ordering::Less);
            // Largest of the smaller values; on ties the first in search order wins.
            smaller.reduce(
                |best, i| if compare(&keys.values[i], &keys.values[best]) == Ordering::Greater { i } else { best },
            )
        }),
        (1, _) => exact_match().or_else(|| {
            let larger = candidates().filter(|&i| compare(&keys.values[i], &needle) == Ordering::Greater);
            larger
                .reduce(|best, i| if compare(&keys.values[i], &keys.values[best]) == Ordering::Less { i } else { best })
        }),
        _ => exact_match(),
    };
    match found {
        Some(i) => {
            let at = if vertical { results.get(i as u32, 0) } else { results.get(0, i as u32) };
            Ok(at.cloned().unwrap_or_default())
        }
        None => args.get(3).map_or(Err(ErrorKind::NA), |a| Ok(a.scalar())),
    }
}

/// CHOOSE(index, value1, ...): 1-based; out of range is `#VALUE!`.
pub fn choose(args: &[Arg]) -> Result<Value, ErrorKind> {
    let index = args[0].number()?.trunc();
    if index < 1.0 || index >= args.len() as f64 {
        return Err(ErrorKind::Value);
    }
    Ok(args[index as usize].scalar())
}

fn comparable(a: &Value, b: &Value) -> bool {
    matches!(
        (a, b),
        (Value::Number(_), Value::Number(_)) | (Value::Text(_), Value::Text(_)) | (Value::Bool(_), Value::Bool(_))
    )
}

fn exact(needle: &Value, values: &[Value]) -> Option<usize> {
    values.iter().position(|v| comparable(v, needle) && compare(v, needle) == Ordering::Equal)
}

/// Last position whose value is <= needle, scanning a list sorted ascending.
fn approx_le(needle: &Value, values: &[Value]) -> Option<usize> {
    let mut best = None;
    for (i, v) in values.iter().enumerate() {
        if !comparable(v, needle) {
            continue;
        }
        if compare(v, needle) == Ordering::Greater {
            break;
        }
        best = Some(i);
    }
    best
}

/// Last position whose value is >= needle, scanning a list sorted descending.
fn approx_ge(needle: &Value, values: &[Value]) -> Option<usize> {
    let mut best = None;
    for (i, v) in values.iter().enumerate() {
        if !comparable(v, needle) {
            continue;
        }
        if compare(v, needle) == Ordering::Less {
            break;
        }
        best = Some(i);
    }
    best
}

#[cfg(test)]
mod tests {
    use crate::functions::test_util::{n, run, t};
    use crate::{ErrorKind, Value};

    #[test]
    fn lookups() {
        assert_eq!(run("VLOOKUP(3, A1:B5, 2, FALSE)"), t("c"));
        assert_eq!(run("VLOOKUP(3.7, A1:B5, 2)"), t("c"));
        assert_eq!(run("VLOOKUP(9, A1:B5, 2, FALSE)"), Value::Error(ErrorKind::NA));
        assert_eq!(run("VLOOKUP(1, A1:B5, 3, FALSE)"), Value::Error(ErrorKind::Ref));
        assert_eq!(run("HLOOKUP(1, A1:B1, 1, FALSE)"), n(1.0));
        assert_eq!(run("INDEX(A1:B5, 2, 2)"), t("b"));
        assert_eq!(run("INDEX(A1:A5, 4)"), n(4.0));
        assert_eq!(run("INDEX(A1:A5, 9)"), Value::Error(ErrorKind::Ref));
        assert_eq!(run("MATCH(\"C\", B1:B5, 0)"), n(3.0));
        assert_eq!(run("MATCH(4.5, A1:A5)"), n(4.0));
        assert_eq!(run("INDEX(B1:B5, MATCH(5, A1:A5, 0))"), t("e"));
    }

    #[test]
    fn xlookup_and_choose() {
        assert_eq!(run("XLOOKUP(3, A1:A5, B1:B5)"), t("c"));
        assert_eq!(run("XLOOKUP(\"D\", B1:B5, A1:A5)"), n(4.0));
        assert_eq!(run("XLOOKUP(9, A1:A5, B1:B5)"), Value::Error(ErrorKind::NA));
        assert_eq!(run("XLOOKUP(9, A1:A5, B1:B5, \"none\")"), t("none"));
        assert_eq!(run("XLOOKUP(3.5, A1:A5, B1:B5, \"none\", -1)"), t("c"));
        assert_eq!(run("XLOOKUP(3.5, A1:A5, B1:B5, \"none\", 1)"), t("d"));
        assert_eq!(run("XLOOKUP(0, A1:A5, B1:B5, \"none\", -1)"), t("none"));
        assert_eq!(run("XLOOKUP(\"?\", B1:B5, A1:A5, 0, 2)"), n(1.0));
        assert_eq!(run("XLOOKUP(\"?\", B1:B5, A1:A5, 0, 2, -1)"), n(5.0));
        assert_eq!(run("XLOOKUP(\"?\", B1:B5, A1:A5, 0, 0)"), n(0.0));
        // 2-D return array: first value of the matched row.
        assert_eq!(run("XLOOKUP(2, A1:A5, A1:B5)"), n(2.0));
        // Horizontal lookup.
        assert_eq!(run("XLOOKUP(\"a\", A1:B1, A1:B1)"), t("a"));
        assert_eq!(run("XLOOKUP(1, A1:A5, B1:B4)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("XLOOKUP(1, A1:A5, B1:B5, 0, 5)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("XLOOKUP(1/0, A1:A5, B1:B5)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("CHOOSE(2, \"x\", \"y\", \"z\")"), t("y"));
        assert_eq!(run("CHOOSE(3.9, 1, 2, 3)"), n(3.0));
        assert_eq!(run("CHOOSE(4, 1, 2, 3)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("CHOOSE(0, 1)"), Value::Error(ErrorKind::Value));
    }
}
