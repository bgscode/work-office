use std::sync::atomic::{AtomicU64, Ordering as AtomicOrdering};

use super::{all_values, for_each_number, Arg};
use crate::value::{ErrorKind, Value};

pub fn sum(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut total = 0.0;
    for_each_number(args, |n| total += n)?;
    Ok(Value::Number(total))
}

pub fn product(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut total = 1.0;
    let mut seen = false;
    for_each_number(args, |n| {
        total *= n;
        seen = true;
    })?;
    Ok(Value::Number(if seen { total } else { 0.0 }))
}

pub fn average(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (mut total, mut count) = (0.0, 0);
    for_each_number(args, |n| {
        total += n;
        count += 1;
    })?;
    if count == 0 {
        return Err(ErrorKind::Div0);
    }
    Ok(Value::Number(total / count as f64))
}

pub fn min(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut best: Option<f64> = None;
    for_each_number(args, |n| best = Some(best.map_or(n, |b| b.min(n))))?;
    Ok(Value::Number(best.unwrap_or(0.0)))
}

pub fn max(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut best: Option<f64> = None;
    for_each_number(args, |n| best = Some(best.map_or(n, |b| b.max(n))))?;
    Ok(Value::Number(best.unwrap_or(0.0)))
}

/// Counts numbers; never fails, errors are just not counted.
pub fn count(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut n = 0;
    for arg in args {
        match arg {
            Arg::Value(v) => n += v.as_number().is_ok() as usize,
            Arg::Range(r) => n += r.values.iter().filter(|v| matches!(v, Value::Number(_))).count(),
        }
    }
    Ok(Value::Number(n as f64))
}

pub fn counta(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(all_values(args).filter(|v| !v.is_empty()).count() as f64))
}

pub fn countblank(args: &[Arg]) -> Result<Value, ErrorKind> {
    let blank = |v: &&Value| matches!(v, Value::Empty) || matches!(v, Value::Text(s) if s.is_empty());
    Ok(Value::Number(all_values(args).filter(blank).count() as f64))
}

pub fn abs(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(args[0].number()?.abs()))
}

pub fn int(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(args[0].number()?.floor()))
}

pub fn sqrt(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?;
    if x < 0.0 {
        return Err(ErrorKind::Num);
    }
    Ok(Value::Number(x.sqrt()))
}

pub fn power(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(args[0].number()?.powf(args[1].number()?)))
}

/// Result has the sign of the divisor, as in Excel: MOD(-3, 2) = 1.
pub fn mod_(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (n, d) = (args[0].number()?, args[1].number()?);
    if d == 0.0 {
        return Err(ErrorKind::Div0);
    }
    Ok(Value::Number(n - d * (n / d).floor()))
}

pub fn pi(_: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(std::f64::consts::PI))
}

/// SUMPRODUCT(array1, ...): arrays must have the same shape; non-numbers count as 0.
pub fn sumproduct(args: &[Arg]) -> Result<Value, ErrorKind> {
    let arrays: Vec<_> = args.iter().map(Arg::as_range).collect();
    let (rows, cols) = (arrays[0].rows, arrays[0].cols);
    if arrays.iter().any(|a| a.rows != rows || a.cols != cols) {
        return Err(ErrorKind::Value);
    }
    let mut total = 0.0;
    for i in 0..arrays[0].values.len() {
        let mut product = 1.0;
        for a in &arrays {
            match &a.values[i] {
                Value::Number(n) => product *= n,
                Value::Error(e) => return Err(*e),
                _ => product = 0.0,
            }
        }
        total += product;
    }
    Ok(Value::Number(total))
}

/// CEILING(number, significance), the Excel 2-argument form.
pub fn ceiling(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (x, sig) = (args[0].number()?, args[1].number()?);
    if x > 0.0 && sig < 0.0 {
        return Err(ErrorKind::Num);
    }
    if sig == 0.0 {
        return Ok(Value::Number(0.0));
    }
    Ok(Value::Number(tidy(tidy(x / sig).ceil() * sig)))
}

/// FLOOR(number, significance), the Excel 2-argument form.
pub fn floor(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (x, sig) = (args[0].number()?, args[1].number()?);
    if x > 0.0 && sig < 0.0 {
        return Err(ErrorKind::Num);
    }
    if sig == 0.0 {
        return if x == 0.0 { Ok(Value::Number(0.0)) } else { Err(ErrorKind::Div0) };
    }
    Ok(Value::Number(tidy(tidy(x / sig).floor() * sig)))
}

/// Round to 15 significant digits to hide binary noise (1.1 / 0.1 = 11.000000000000002).
fn tidy(x: f64) -> f64 {
    format!("{:.14e}", x).parse().unwrap_or(x)
}

pub fn sign(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?;
    Ok(Value::Number(if x > 0.0 {
        1.0
    } else if x < 0.0 {
        -1.0
    } else {
        0.0
    }))
}

pub fn exp(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(args[0].number()?.exp()))
}

pub fn ln(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?;
    if x <= 0.0 {
        return Err(ErrorKind::Num);
    }
    Ok(Value::Number(x.ln()))
}

/// LOG(number, [base = 10])
pub fn log(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?;
    let base = args.get(1).map_or(Ok(10.0), Arg::number)?;
    if x <= 0.0 || base <= 0.0 {
        return Err(ErrorKind::Num);
    }
    if base == 1.0 {
        return Err(ErrorKind::Div0);
    }
    Ok(Value::Number(if base == 10.0 { x.log10() } else { x.ln() / base.ln() }))
}

pub fn log10(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?;
    if x <= 0.0 {
        return Err(ErrorKind::Num);
    }
    Ok(Value::Number(x.log10()))
}

/// TRUNC(number, [digits = 0])
pub fn trunc(args: &[Arg]) -> Result<Value, ErrorKind> {
    let digits = args.get(1).map_or(Ok(0.0), Arg::number)?;
    round_to(args[0].number()?, digits, f64::trunc)
}

/// Rounds away from zero to the nearest even integer.
pub fn even(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?;
    Ok(Value::Number(((tidy(x.abs()) / 2.0).ceil() * 2.0).copysign(x)))
}

/// Rounds away from zero to the nearest odd integer; ODD(0) is 1.
pub fn odd(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?;
    let mut n = tidy(x.abs()).ceil();
    if n % 2.0 == 0.0 {
        n += 1.0;
    }
    Ok(Value::Number(if x < 0.0 { -n } else { n }))
}

pub fn fact(args: &[Arg]) -> Result<Value, ErrorKind> {
    let x = args[0].number()?.trunc();
    if x < 0.0 {
        return Err(ErrorKind::Num);
    }
    if x > 170.0 {
        return Err(ErrorKind::Num);
    }
    Ok(Value::Number((1..=x as u32).fold(1.0, |acc, k| acc * k as f64)))
}

/// Integer arguments for GCD/LCM: truncated, `#NUM!` when negative or too large.
fn integers(args: &[Arg]) -> Result<Vec<u64>, ErrorKind> {
    let mut out = Vec::new();
    let mut bad = false;
    for_each_number(args, |n| {
        let n = n.trunc();
        if !(0.0..9.007_199_254_740_992e15).contains(&n) {
            bad = true;
        }
        out.push(n as u64);
    })?;
    if bad {
        return Err(ErrorKind::Num);
    }
    Ok(out)
}

fn gcd_pair(mut a: u64, mut b: u64) -> u64 {
    while b != 0 {
        (a, b) = (b, a % b);
    }
    a
}

pub fn gcd(args: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(integers(args)?.into_iter().fold(0, gcd_pair) as f64))
}

pub fn lcm(args: &[Arg]) -> Result<Value, ErrorKind> {
    let mut acc = 1.0;
    for n in integers(args)? {
        if n == 0 {
            return Ok(Value::Number(0.0));
        }
        acc = acc / gcd_pair(acc as u64, n) as f64 * n as f64;
        if acc >= 9.007_199_254_740_992e15 {
            return Err(ErrorKind::Num);
        }
    }
    Ok(Value::Number(acc))
}

/// A number in [0, 1). Seeded from the clock plus a call counter and mixed
/// with SplitMix64: fine for spreadsheets, not for cryptography.
pub(crate) fn random_unit() -> f64 {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    const GOLDEN: u64 = 0x9E37_79B9_7F4A_7C15;
    let count = COUNTER.fetch_add(1, AtomicOrdering::Relaxed);
    let mut z = crate::date::now_serial().to_bits() ^ count.wrapping_mul(GOLDEN);
    z = z.wrapping_add(GOLDEN);
    z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
    z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
    z ^= z >> 31;
    (z >> 11) as f64 / (1u64 << 53) as f64
}

pub fn rand(_: &[Arg]) -> Result<Value, ErrorKind> {
    Ok(Value::Number(random_unit()))
}

/// RANDBETWEEN(bottom, top): a random integer in [ceil(bottom), floor(top)].
pub fn randbetween(args: &[Arg]) -> Result<Value, ErrorKind> {
    let lo = args[0].number()?.ceil();
    let hi = args[1].number()?.floor();
    if lo > hi {
        return Err(ErrorKind::Num);
    }
    Ok(Value::Number((lo + (random_unit() * (hi - lo + 1.0)).floor()).min(hi)))
}

pub fn round(args: &[Arg]) -> Result<Value, ErrorKind> {
    round_with(args, f64::round)
}

pub fn roundup(args: &[Arg]) -> Result<Value, ErrorKind> {
    round_with(args, |x| x.abs().ceil().copysign(x))
}

pub fn rounddown(args: &[Arg]) -> Result<Value, ErrorKind> {
    round_with(args, f64::trunc)
}

fn round_with(args: &[Arg], op: fn(f64) -> f64) -> Result<Value, ErrorKind> {
    round_to(args[0].number()?, args[1].number()?, op)
}

fn round_to(x: f64, digits: f64, op: fn(f64) -> f64) -> Result<Value, ErrorKind> {
    let digits = digits.trunc() as i32;
    let factor = 10f64.powi(digits);
    // Trim binary noise first so ROUND(2.675, 2) gives 2.68 like Excel.
    let scaled: f64 = format!("{:.12e}", x * factor).parse().unwrap_or(x * factor);
    Ok(Value::Number(op(scaled) / factor))
}

#[cfg(test)]
mod tests {
    use crate::functions::test_util::{n, run};
    use crate::ErrorKind;
    use crate::Value;

    #[test]
    fn aggregates() {
        assert_eq!(run("SUM(A1:A5)"), n(15.0));
        assert_eq!(run("SUM(A1:C5, 10, \"5\")"), n(30.0));
        assert_eq!(run("PRODUCT(A1:A4)"), n(24.0));
        assert_eq!(run("AVERAGE(A1:A5)"), n(3.0));
        assert_eq!(run("AVERAGE(B1:B5)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("MIN(A1:A5, -1)"), n(-1.0));
        assert_eq!(run("MAX(A1:B5)"), n(5.0));
        assert_eq!(run("COUNT(A1:D5)"), n(5.0));
        assert_eq!(run("COUNTA(A1:D5)"), n(12.0));
        assert_eq!(run("COUNTBLANK(D1:D3)"), n(3.0));
    }

    #[test]
    fn rounding_and_arithmetic() {
        assert_eq!(run("ROUND(2.675, 2)"), n(2.68));
        assert_eq!(run("ROUND(1234, -2)"), n(1200.0));
        assert_eq!(run("ROUNDUP(-1.21, 1)"), n(-1.3));
        assert_eq!(run("ROUNDDOWN(1.29, 1)"), n(1.2));
        assert_eq!(run("INT(-1.5)"), n(-2.0));
        assert_eq!(run("MOD(-3, 2)"), n(1.0));
        assert_eq!(run("MOD(1, 0)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("SQRT(-1)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("POWER(2, 10)"), n(1024.0));
        assert_eq!(run("ABS(-3)"), n(3.0));
    }

    #[test]
    fn sumproduct() {
        assert_eq!(run("SUMPRODUCT(A1:A5, A1:A5)"), n(55.0));
        assert_eq!(run("SUMPRODUCT(A1:A3)"), n(6.0));
        assert_eq!(run("SUMPRODUCT(A1:A5, B1:B5)"), n(0.0));
        assert_eq!(run("SUMPRODUCT(3, 4)"), n(12.0));
        assert_eq!(run("SUMPRODUCT(A1:A5, A1:A4)"), Value::Error(ErrorKind::Value));
    }

    #[test]
    fn ceiling_floor_and_integers() {
        assert_eq!(run("CEILING(2.5, 1)"), n(3.0));
        assert_eq!(run("CEILING(1.1, 0.1)"), n(1.1));
        assert_eq!(run("CEILING(-2.5, 2)"), n(-2.0));
        assert_eq!(run("CEILING(-2.5, -2)"), n(-4.0));
        assert_eq!(run("CEILING(5, 0)"), n(0.0));
        assert_eq!(run("CEILING(2.5, -1)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("FLOOR(2.5, 1)"), n(2.0));
        assert_eq!(run("FLOOR(-2.5, 2)"), n(-4.0));
        assert_eq!(run("FLOOR(-2.5, -2)"), n(-2.0));
        assert_eq!(run("FLOOR(7, 0)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("FLOOR(2.5, -1)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("TRUNC(8.9)"), n(8.0));
        assert_eq!(run("TRUNC(-8.96, 1)"), n(-8.9));
        assert_eq!(run("EVEN(1.5)"), n(2.0));
        assert_eq!(run("EVEN(3)"), n(4.0));
        assert_eq!(run("EVEN(-1)"), n(-2.0));
        assert_eq!(run("EVEN(0)"), n(0.0));
        assert_eq!(run("ODD(1.5)"), n(3.0));
        assert_eq!(run("ODD(3)"), n(3.0));
        assert_eq!(run("ODD(2)"), n(3.0));
        assert_eq!(run("ODD(-2)"), n(-3.0));
        assert_eq!(run("ODD(0)"), n(1.0));
        assert_eq!(run("SIGN(-4)"), n(-1.0));
        assert_eq!(run("SIGN(0)"), n(0.0));
        assert_eq!(run("SIGN(A2)"), n(1.0));
        assert_eq!(run("FACT(5)"), n(120.0));
        assert_eq!(run("FACT(0)"), n(1.0));
        assert_eq!(run("FACT(-1)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("GCD(24, 36)"), n(12.0));
        assert_eq!(run("GCD(A1:A5)"), n(1.0));
        assert_eq!(run("GCD(0, 5)"), n(5.0));
        assert_eq!(run("GCD(-2, 4)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("LCM(4, 6)"), n(12.0));
        assert_eq!(run("LCM(A1:A5)"), n(60.0));
        assert_eq!(run("LCM(0, 6)"), n(0.0));
        assert_eq!(run("LCM(\"x\", 6)"), Value::Error(ErrorKind::Value));
    }

    #[test]
    fn exponents_and_logs() {
        assert_eq!(run("EXP(0)"), n(1.0));
        assert_eq!(run("LN(1)"), n(0.0));
        assert_eq!(run("LN(0)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("LOG(1000)"), n(3.0));
        assert_eq!(run("LOG(8, 2)"), n(3.0));
        assert_eq!(run("LOG(8, 1)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run("LOG(-1)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("LOG10(0.01)"), n(-2.0));
        assert_eq!(run("LOG10(0)"), Value::Error(ErrorKind::Num));
        assert_eq!(run("EXP(1000)"), Value::Error(ErrorKind::Num));
    }

    #[test]
    fn random_numbers() {
        for _ in 0..200 {
            let Value::Number(x) = run("RAND()") else { panic!("RAND must return a number") };
            assert!((0.0..1.0).contains(&x));
            let Value::Number(k) = run("RANDBETWEEN(-2, 3)") else { panic!("RANDBETWEEN must return a number") };
            assert!((-2.0..=3.0).contains(&k) && k.fract() == 0.0);
        }
        assert_ne!(run("RAND()"), run("RAND()"));
        assert_eq!(run("RANDBETWEEN(1.5, 2.5)"), n(2.0));
        assert_eq!(run("RANDBETWEEN(5, 1)"), Value::Error(ErrorKind::Num));
    }
}
