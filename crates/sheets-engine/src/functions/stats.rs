//! Statistics: MEDIAN, MODE, LARGE/SMALL, RANK, STDEV/VAR and their
//! population variants. Like SUM, they read numbers only from ranges and
//! coerce scalar arguments.

use super::{for_each_number, Arg};
use crate::value::{ErrorKind, Value};

fn numbers(args: &[Arg]) -> Result<Vec<f64>, ErrorKind> {
    let mut out = Vec::new();
    for_each_number(args, |n| out.push(n))?;
    Ok(out)
}

fn sorted(args: &[Arg]) -> Result<Vec<f64>, ErrorKind> {
    let mut nums = numbers(args)?;
    nums.sort_by(f64::total_cmp);
    Ok(nums)
}

pub fn median(args: &[Arg]) -> Result<Value, ErrorKind> {
    let nums = sorted(args)?;
    let len = nums.len();
    if len == 0 {
        return Err(ErrorKind::Num);
    }
    let mid = len / 2;
    Ok(Value::Number(if len % 2 == 1 { nums[mid] } else { (nums[mid - 1] + nums[mid]) / 2.0 }))
}

/// Most frequent number; on a tie, the one that appears first. `#N/A` when
/// no number repeats.
pub fn mode(args: &[Arg]) -> Result<Value, ErrorKind> {
    let nums = numbers(args)?;
    let mut best: Option<(f64, usize)> = None;
    for (i, &x) in nums.iter().enumerate() {
        let count = nums[i..].iter().filter(|&&y| y == x).count();
        if count > 1 && best.is_none_or(|(_, c)| count > c) {
            best = Some((x, count));
        }
    }
    best.map(|(x, _)| Value::Number(x)).ok_or(ErrorKind::NA)
}

/// The k-th value counted from the top (LARGE) or bottom (SMALL); k rounds up.
fn kth(args: &[Arg], from_top: bool) -> Result<Value, ErrorKind> {
    let mut nums = sorted(&args[..1])?;
    if from_top {
        nums.reverse();
    }
    let k = args[1].number()?.ceil();
    if k < 1.0 || k > nums.len() as f64 {
        return Err(ErrorKind::Num);
    }
    Ok(Value::Number(nums[k as usize - 1]))
}

pub fn large(args: &[Arg]) -> Result<Value, ErrorKind> {
    kth(args, true)
}

pub fn small(args: &[Arg]) -> Result<Value, ErrorKind> {
    kth(args, false)
}

/// RANK(number, ref, [order = 0]): 0 ranks the largest as 1, non-zero the
/// smallest. Ties share the best rank. `#N/A` if the number is not in ref.
pub fn rank(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?;
    let nums = numbers(&args[1..2])?;
    let ascending = args.get(2).map_or(Ok(0.0), Arg::number)? != 0.0;
    if !nums.contains(&x) {
        return Err(ErrorKind::NA);
    }
    let better = nums.iter().filter(|&&y| if ascending { y < x } else { y > x }).count();
    Ok(Value::Number(better as f64 + 1.0))
}

/// Sample (n - 1) or population (n) variance; `#DIV/0!` when the divisor is 0.
fn variance(args: &[Arg], sample: bool) -> Result<f64, ErrorKind> {
    let nums = numbers(args)?;
    let n = nums.len();
    let divisor = if sample { n.saturating_sub(1) } else { n };
    if divisor == 0 {
        return Err(ErrorKind::Div0);
    }
    let mean = nums.iter().sum::<f64>() / n as f64;
    Ok(nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / divisor as f64)
}

pub fn var(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(variance(args, true)?))
}

pub fn varp(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(variance(args, false)?))
}

pub fn stdev(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(variance(args, true)?.sqrt()))
}

pub fn stdevp(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(variance(args, false)?.sqrt()))
}

#[cfg(test)]
mod tests {
    use crate::functions::test_util::{n, run};
    use crate::{ErrorKind, Value};

    fn approx(formula: &str, expected: f64) {
        match run(formula) {
            Value::Number(x) => assert!((x - expected).abs() < 1e-12, "{formula} = {x}, expected {expected}"),
            other => panic!("{formula} = {other:?}"),
        }
    }

    #[test]
    fn central_tendency() {
        assert_eq!(run("MEDIAN(A1:A5)"), n(3.0));
        assert_eq!(run("MEDIAN(A1:A4)"), n(2.5));
        assert_eq!(run("MEDIAN(5, 1, 3, B1:B5)"), n(3.0));
        assert_eq!(run("MEDIAN(B1:B5)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("MEDIAN(\"x\")"), Value::Error(ErrorKind::Value));
        assert_eq!(run("MODE(1, 2, 2, 3, 3)"), n(2.0));
        assert_eq!(run("MODE.SNGL(A1:A5, 4)"), n(4.0));
        assert_eq!(run("MODE(A1:A5)"), Value::Error(ErrorKind::NA));
    }

    #[test]
    fn ordering() {
        assert_eq!(run("LARGE(A1:A5, 1)"), n(5.0));
        assert_eq!(run("LARGE(A1:B5, 2)"), n(4.0));
        assert_eq!(run("SMALL(A1:A5, 2)"), n(2.0));
        assert_eq!(run("SMALL(A1:A5, 6)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("LARGE(A1:A5, 0)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("RANK(4, A1:A5)"), n(2.0));
        assert_eq!(run("RANK.EQ(4, A1:A5, 1)"), n(4.0));
        assert_eq!(run("RANK(9, A1:A5)"), Value::Error(ErrorKind::NA));
    }

    #[test]
    fn spread() {
        approx("VAR(A1:A5)", 2.5);
        approx("VAR.S(A1:A5)", 2.5);
        approx("VARP(A1:A5)", 2.0);
        approx("VAR.P(A1:A5)", 2.0);
        approx("STDEV(A1:A5)", 2.5f64.sqrt());
        approx("STDEV.S(A1:A5)", 2.5f64.sqrt());
        approx("STDEVP(A1:A5)", 2.0f64.sqrt());
        approx("STDEV.P(A1:A5)", 2.0f64.sqrt());
        assert_eq!(run("STDEV(1)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("VARP(B1:B5)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("STDEVP(7)"), n(0.0));
    }
}
