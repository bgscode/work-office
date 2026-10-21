//! SUMIF / COUNTIF / AVERAGEIF, their multi-condition *IFS versions, and the
//! criteria syntax they share:
//! `5`, `">5"`, `"<>apple"`, `"=a*"` (wildcards `*` and `?`).

use std::cmp::Ordering;

use super::Arg;
use crate::value::{compare, ErrorKind, Value};

pub fn countif(args: &[Arg]) -> Result<Value, ErrorKind> {
    let crit = Criteria::new(&args[1].scalar())?;
    let range = args[0].as_range();
    Ok(Value::Number(range.values.iter().filter(|v| crit.matches(v)).count() as f64))
}

pub fn sumif(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (total, _) = sum_matching(args)?;
    Ok(Value::Number(total))
}

pub fn averageif(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (total, count) = sum_matching(args)?;
    if count == 0 {
        return Err(ErrorKind::Div0);
    }
    Ok(Value::Number(total / count as f64))
}

/// Sum the numbers in the sum range (or the test range itself) where the test range matches.
fn sum_matching(args: &[Arg]) -> Result<(f64, usize), ErrorKind> {
    let crit = Criteria::new(&args[1].scalar())?;
    let test = args[0].as_range();
    let sums = args.get(2).map_or_else(|| test.clone(), Arg::as_range);
    let (mut total, mut count) = (0.0, 0);
    for row in 0..test.rows {
        for col in 0..test.cols {
            if !crit.matches(test.get(row, col).unwrap()) {
                continue;
            }
            match sums.get(row, col) {
                Some(Value::Number(n)) => {
                    total += n;
                    count += 1;
                }
                Some(Value::Error(e)) => return Err(*e),
                _ => {}
            }
        }
    }
    Ok((total, count))
}

/// SUMIFS(sum_range, criteria_range1, criteria1, ...)
pub fn sumifs(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(numbers_where(args)?.iter().sum()))
}

/// AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
pub fn averageifs(args: &[Arg]) -> Result<Value, ErrorKind> {
    let nums = numbers_where(args)?;
    if nums.is_empty() {
        return Err(ErrorKind::Div0);
    }
    Ok(Value::Number(nums.iter().sum::<f64>() / nums.len() as f64))
}

/// MAXIFS(max_range, criteria_range1, criteria1, ...); 0 when nothing matches.
pub fn maxifs(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(numbers_where(args)?.into_iter().reduce(f64::max).unwrap_or(0.0)))
}

/// MINIFS(min_range, criteria_range1, criteria1, ...); 0 when nothing matches.
pub fn minifs(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(numbers_where(args)?.into_iter().reduce(f64::min).unwrap_or(0.0)))
}

/// COUNTIFS(criteria_range1, criteria1, ...)
pub fn countifs(args: &[Arg]) -> Result<Value, ErrorKind> {
    let first = args[0].as_range();
    let mask = matching_cells(args, first.rows, first.cols)?;
    Ok(Value::Number(mask.into_iter().filter(|&ok| ok).count() as f64))
}

/// Numbers in `args[0]` at the positions where every (range, criteria) pair
/// in the rest of `args` matches. Text, booleans and blanks are skipped.
fn numbers_where(args: &[Arg]) -> Result<Vec<f64>, ErrorKind> {
    let target = args[0].as_range();
    let mask = matching_cells(&args[1..], target.rows, target.cols)?;
    let mut out = Vec::new();
    for (v, ok) in target.values.iter().zip(mask) {
        match v {
            Value::Number(n) if ok => out.push(*n),
            Value::Error(e) if ok => return Err(*e),
            _ => {}
        }
    }
    Ok(out)
}

/// One flag per cell: does every (range, criteria) pair match there?
/// All ranges must be `rows` x `cols`, otherwise `#VALUE!`.
fn matching_cells(pairs: &[Arg], rows: u32, cols: u32) -> Result<Vec<bool>, ErrorKind> {
    if pairs.is_empty() || !pairs.len().is_multiple_of(2) {
        return Err(ErrorKind::Value);
    }
    let mut mask = vec![true; rows as usize * cols as usize];
    for pair in pairs.chunks(2) {
        let range = pair[0].as_range();
        if range.rows != rows || range.cols != cols {
            return Err(ErrorKind::Value);
        }
        let crit = Criteria::new(&pair[1].scalar())?;
        for (ok, v) in mask.iter_mut().zip(&range.values) {
            *ok = *ok && crit.matches(v);
        }
    }
    Ok(mask)
}

struct Criteria {
    op: Op,
    target: Value,
}

#[derive(Clone, Copy, PartialEq)]
enum Op {
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

impl Criteria {
    fn new(criteria: &Value) -> Result<Self, ErrorKind> {
        let text = match criteria {
            Value::Error(e) => return Err(*e),
            Value::Text(s) => s.clone(),
            other => return Ok(Criteria { op: Op::Eq, target: other.clone() }),
        };
        let (op, rest) = [("<=", Op::Le), (">=", Op::Ge), ("<>", Op::Ne), ("<", Op::Lt), (">", Op::Gt), ("=", Op::Eq)]
            .iter()
            .find_map(|(p, op)| text.strip_prefix(p).map(|rest| (*op, rest)))
            .unwrap_or((Op::Eq, text.as_str()));
        let target = match Value::from_input(rest) {
            Value::Empty => Value::Text(String::new()),
            v => v,
        };
        Ok(Criteria { op, target })
    }

    fn matches(&self, v: &Value) -> bool {
        let eq = self.equals(v);
        match self.op {
            Op::Eq => eq,
            Op::Ne => !eq,
            _ => {
                let same_type = matches!(
                    (v, &self.target),
                    (Value::Number(_), Value::Number(_)) | (Value::Text(_), Value::Text(_))
                );
                if !same_type {
                    return false;
                }
                let ord = compare(v, &self.target);
                match self.op {
                    Op::Lt => ord == Ordering::Less,
                    Op::Le => ord != Ordering::Greater,
                    Op::Gt => ord == Ordering::Greater,
                    _ => ord != Ordering::Less,
                }
            }
        }
    }

    fn equals(&self, v: &Value) -> bool {
        match (&self.target, v) {
            (Value::Text(t), Value::Empty) => t.is_empty(),
            (Value::Text(t), Value::Text(s)) => wildcard_match(&t.to_lowercase(), &s.to_lowercase()),
            (Value::Number(a), Value::Number(b)) => a == b,
            (Value::Bool(a), Value::Bool(b)) => a == b,
            (Value::Error(a), Value::Error(b)) => a == b,
            _ => false,
        }
    }
}

/// `*` matches any run of characters, `?` exactly one, `~` escapes the next one.
pub(super) fn wildcard_match(pattern: &str, s: &str) -> bool {
    fn go(p: &[char], s: &[char]) -> bool {
        match p.first() {
            None => s.is_empty(),
            Some('*') => (0..=s.len()).any(|i| go(&p[1..], &s[i..])),
            Some('?') => !s.is_empty() && go(&p[1..], &s[1..]),
            Some('~') if p.len() > 1 => s.first() == Some(&p[1]) && go(&p[2..], &s[1..]),
            Some(c) => s.first() == Some(c) && go(&p[1..], &s[1..]),
        }
    }
    let p: Vec<char> = pattern.chars().collect();
    let s: Vec<char> = s.chars().collect();
    go(&p, &s)
}

#[cfg(test)]
mod tests {
    use crate::functions::test_util::{n, run};
    use crate::{ErrorKind, Value};

    #[test]
    fn conditional_aggregates() {
        assert_eq!(run("COUNTIF(A1:A5, \">2\")"), n(3.0));
        assert_eq!(run("COUNTIF(A1:A5, 3)"), n(1.0));
        assert_eq!(run("COUNTIF(B1:B5, \"<>c\")"), n(4.0));
        assert_eq!(run("COUNTIF(B1:B5, \"?\")"), n(5.0));
        assert_eq!(run("COUNTIF(B1:B5, \"A\")"), n(1.0));
        assert_eq!(run("SUMIF(A1:A5, \"<=2\")"), n(3.0));
        assert_eq!(run("SUMIF(B1:B5, \"b\", A1:A5)"), n(2.0));
        assert_eq!(run("AVERAGEIF(A1:A5, \">3\")"), n(4.5));
        assert_eq!(run("AVERAGEIF(A1:A5, \">9\")"), Value::Error(ErrorKind::Div0));
    }

    #[test]
    fn multi_condition_aggregates() {
        assert_eq!(run("SUMIFS(A1:A5, A1:A5, \">1\", B1:B5, \"<>d\")"), n(10.0));
        assert_eq!(run("SUMIFS(A1:A5, B1:B5, \"c\")"), n(3.0));
        assert_eq!(run("COUNTIFS(A1:A5, \">=2\", A1:A5, \"<=4\")"), n(3.0));
        assert_eq!(run("COUNTIFS(B1:B5, \"?\")"), n(5.0));
        assert_eq!(run("AVERAGEIFS(A1:A5, B1:B5, \"<c\")"), n(1.5));
        assert_eq!(run("AVERAGEIFS(A1:A5, B1:B5, \"z\")"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("MAXIFS(A1:A5, B1:B5, \"<>e\")"), n(4.0));
        assert_eq!(run("MINIFS(A1:A5, A1:A5, \">2\")"), n(3.0));
        assert_eq!(run("MAXIFS(A1:A5, A1:A5, \">9\")"), n(0.0));
        // Different shapes, and a criteria range without its criteria.
        assert_eq!(run("SUMIFS(A1:A5, B1:B4, \"a\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("COUNTIFS(A1:A5, 1, B1:B5)"), Value::Error(ErrorKind::Value));
        assert_eq!(run("SUMIFS(A1:A5, B1:B5, 1/0)"), Value::Error(ErrorKind::Div0));
    }
}
