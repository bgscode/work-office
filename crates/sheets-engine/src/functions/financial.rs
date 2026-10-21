//! Financial functions: loans and annuities (PMT, IPMT, PPMT, FV, PV, NPER,
//! RATE), cash flows (NPV, XNPV, IRR, XIRR), depreciation (SLN, DB, DDB) and
//! interest-rate conversion (EFFECT, NOMINAL).
//!
//! Sign convention as in Excel: money paid out is negative, money received is
//! positive. `type` is 0 for payments at the end of each period (default) and
//! 1 for payments at the start.

use super::{Arg, RangeValues};
use crate::value::{ErrorKind, Value};

fn opt(args: &[Arg], i: usize, default: f64) -> Result<f64, ErrorKind> {
    match args.get(i) {
        Some(a) => a.number(),
        None => Ok(default),
    }
}

/// `type` argument: any non-zero value means "start of period".
fn due(args: &[Arg], i: usize) -> Result<f64, ErrorKind> {
    Ok(if opt(args, i, 0.0)? != 0.0 { 1.0 } else { 0.0 })
}

fn num(x: f64) -> Result<Value, ErrorKind> {
    if x.is_finite() {
        Ok(Value::Number(x))
    } else {
        Err(ErrorKind::Num)
    }
}

pub(crate) fn pmt_of(rate: f64, nper: f64, pv: f64, fv: f64, due: f64) -> f64 {
    if rate == 0.0 {
        return -(pv + fv) / nper;
    }
    let f = (1.0 + rate).powf(nper);
    -(rate * (fv + pv * f)) / ((1.0 + rate * due) * (f - 1.0))
}

pub(crate) fn fv_of(rate: f64, nper: f64, pmt: f64, pv: f64, due: f64) -> f64 {
    if rate == 0.0 {
        return -(pv + pmt * nper);
    }
    let f = (1.0 + rate).powf(nper);
    -(pv * f + pmt * (1.0 + rate * due) * (f - 1.0) / rate)
}

/// PMT(rate, nper, pv, [fv], [type])
pub fn pmt(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (rate, nper, pv) = (args[0].number()?, args[1].number()?, args[2].number()?);
    if nper == 0.0 {
        return Err(ErrorKind::Num);
    }
    num(pmt_of(rate, nper, pv, opt(args, 3, 0.0)?, due(args, 4)?))
}

/// FV(rate, nper, pmt, [pv], [type])
pub fn fv(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (rate, nper, pmt) = (args[0].number()?, args[1].number()?, args[2].number()?);
    num(fv_of(rate, nper, pmt, opt(args, 3, 0.0)?, due(args, 4)?))
}

/// PV(rate, nper, pmt, [fv], [type])
pub fn pv(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (rate, nper, pmt) = (args[0].number()?, args[1].number()?, args[2].number()?);
    let (fv, t) = (opt(args, 3, 0.0)?, due(args, 4)?);
    if rate == 0.0 {
        return num(-(fv + pmt * nper));
    }
    let f = (1.0 + rate).powf(nper);
    num(-(fv + pmt * (1.0 + rate * t) * (f - 1.0) / rate) / f)
}

/// NPER(rate, pmt, pv, [fv], [type])
pub fn nper(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (rate, pmt, pv) = (args[0].number()?, args[1].number()?, args[2].number()?);
    let (fv, t) = (opt(args, 3, 0.0)?, due(args, 4)?);
    if rate == 0.0 {
        if pmt == 0.0 {
            return Err(ErrorKind::Num);
        }
        return num(-(pv + fv) / pmt);
    }
    let a = pmt * (1.0 + rate * t);
    let ratio = (a - fv * rate) / (a + pv * rate);
    if ratio <= 0.0 || rate <= -1.0 {
        return Err(ErrorKind::Num);
    }
    num(ratio.ln() / (1.0 + rate).ln())
}

/// Interest part of payment `per` (1-based), as LibreOffice and Excel compute it.
fn ipmt_of(rate: f64, per: f64, nper: f64, pv: f64, fv: f64, due: f64) -> f64 {
    let pmt = pmt_of(rate, nper, pv, fv, due);
    let interest = if per == 1.0 {
        if due == 1.0 {
            0.0
        } else {
            -pv
        }
    } else if due == 1.0 {
        fv_of(rate, per - 2.0, pmt, pv, 1.0) - pmt
    } else {
        fv_of(rate, per - 1.0, pmt, pv, 0.0)
    };
    interest * rate
}

fn period_args(args: &[Arg]) -> Result<(f64, f64, f64, f64, f64, f64), ErrorKind> {
    let (rate, per, nper, pv) = (args[0].number()?, args[1].number()?, args[2].number()?, args[3].number()?);
    if per < 1.0 || per > nper || nper <= 0.0 {
        return Err(ErrorKind::Num);
    }
    Ok((rate, per, nper, pv, opt(args, 4, 0.0)?, due(args, 5)?))
}

/// IPMT(rate, per, nper, pv, [fv], [type])
pub fn ipmt(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (rate, per, nper, pv, fv, t) = period_args(args)?;
    num(ipmt_of(rate, per, nper, pv, fv, t))
}

/// PPMT(rate, per, nper, pv, [fv], [type])
pub fn ppmt(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (rate, per, nper, pv, fv, t) = period_args(args)?;
    num(pmt_of(rate, nper, pv, fv, t) - ipmt_of(rate, per, nper, pv, fv, t))
}

/// Numbers from the arguments the way NPV reads them: ranges skip text,
/// booleans and blanks; direct arguments are coerced.
fn cash_flows(args: &[Arg]) -> Result<Vec<f64>, ErrorKind> {
    let mut out = Vec::new();
    super::for_each_number(args, |n| out.push(n))?;
    Ok(out)
}

/// Every cell of a range argument as a number; text or blanks are `#VALUE!`
/// (XNPV/XIRR need the values and dates to line up one to one).
fn strict_numbers(r: &RangeValues) -> Result<Vec<f64>, ErrorKind> {
    r.values
        .iter()
        .map(|v| match v {
            Value::Number(n) => Ok(*n),
            Value::Error(e) => Err(*e),
            _ => Err(ErrorKind::Value),
        })
        .collect()
}

/// NPV(rate, value1, [value2], ...): values at the end of periods 1, 2, ...
pub fn npv(args: &[Arg]) -> Result<Value, ErrorKind> {
    let rate = args[0].number()?;
    if rate == -1.0 {
        return Err(ErrorKind::Div0);
    }
    let flows = cash_flows(&args[1..])?;
    num(flows.iter().enumerate().map(|(i, v)| v / (1.0 + rate).powi(i as i32 + 1)).sum())
}

fn xnpv_of(rate: f64, values: &[f64], dates: &[f64]) -> f64 {
    values.iter().zip(dates).map(|(v, d)| v / (1.0 + rate).powf((d - dates[0]) / 365.0)).sum()
}

fn dated_flows(args: &[Arg]) -> Result<(Vec<f64>, Vec<f64>), ErrorKind> {
    let values = strict_numbers(&args[0].as_range())?;
    let dates: Vec<f64> = strict_numbers(&args[1].as_range())?.into_iter().map(f64::trunc).collect();
    if values.len() != dates.len() || values.is_empty() {
        return Err(ErrorKind::Num);
    }
    if dates.iter().any(|d| *d < dates[0] || *d < 0.0) {
        return Err(ErrorKind::Num);
    }
    Ok((values, dates))
}

/// XNPV(rate, values, dates): cash flows on arbitrary dates, discounted on a 365-day year.
/// (`dated_flows` takes the values and dates, so skip the rate.)
pub fn xnpv(args: &[Arg]) -> Result<Value, ErrorKind> {
    let rate = args[0].number()?;
    if rate <= -1.0 {
        return Err(ErrorKind::Num);
    }
    let (values, dates) = dated_flows(&args[1..])?;
    num(xnpv_of(rate, &values, &dates))
}

/// Find a root of `f` near `guess`: Newton's method with a numeric
/// derivative, then a bracketing search if Newton wanders off.
fn solve(guess: f64, f: impl Fn(f64) -> f64) -> Result<f64, ErrorKind> {
    let mut x = guess;
    for _ in 0..100 {
        let y = f(x);
        if !y.is_finite() {
            break;
        }
        if y.abs() < 1e-10 {
            return Ok(x);
        }
        let h = 1e-7 * x.abs().max(1e-3);
        let slope = (f(x + h) - f(x - h)) / (2.0 * h);
        if slope == 0.0 || !slope.is_finite() {
            break;
        }
        let next = x - y / slope;
        if (next - x).abs() < 1e-12 {
            return if f(next).abs() < 1e-6 { Ok(next) } else { Err(ErrorKind::Num) };
        }
        // Rates at or below -100% are meaningless here.
        x = if next <= -1.0 { (x - 1.0) / 2.0 } else { next };
    }
    // Fallback: scan for a sign change, then bisect.
    let grid: Vec<f64> = (0..400).map(|i| -0.99 + i as f64 * 0.025).collect();
    for w in grid.windows(2) {
        let (mut lo, mut hi) = (w[0], w[1]);
        let (flo, fhi) = (f(lo), f(hi));
        if !flo.is_finite() || !fhi.is_finite() || flo.signum() == fhi.signum() {
            continue;
        }
        for _ in 0..200 {
            let mid = (lo + hi) / 2.0;
            if f(mid).signum() == f(lo).signum() {
                lo = mid;
            } else {
                hi = mid;
            }
        }
        return Ok((lo + hi) / 2.0);
    }
    Err(ErrorKind::Num)
}

fn has_both_signs(values: &[f64]) -> bool {
    values.iter().any(|v| *v > 0.0) && values.iter().any(|v| *v < 0.0)
}

/// IRR(values, [guess]): the rate at which the NPV of the flows (period 0, 1, ...) is zero.
pub fn irr(args: &[Arg]) -> Result<Value, ErrorKind> {
    let flows = cash_flows(&args[..1])?;
    if !has_both_signs(&flows) {
        return Err(ErrorKind::Num);
    }
    let npv = |r: f64| flows.iter().enumerate().map(|(i, v)| v / (1.0 + r).powi(i as i32)).sum::<f64>();
    num(solve(opt(args, 1, 0.1)?, npv)?)
}

/// XIRR(values, dates, [guess])
pub fn xirr(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (values, dates) = dated_flows(args)?;
    if !has_both_signs(&values) {
        return Err(ErrorKind::Num);
    }
    num(solve(opt(args, 2, 0.1)?, |r| xnpv_of(r, &values, &dates))?)
}

/// RATE(nper, pmt, pv, [fv], [type], [guess]): interest rate per period.
pub fn rate(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (nper, pmt, pv) = (args[0].number()?, args[1].number()?, args[2].number()?);
    let (fv, t, guess) = (opt(args, 3, 0.0)?, due(args, 4)?, opt(args, 5, 0.1)?);
    if nper <= 0.0 {
        return Err(ErrorKind::Num);
    }
    let balance = |r: f64| {
        if r.abs() < 1e-12 {
            pv + pmt * nper + fv
        } else {
            let f = (1.0 + r).powf(nper);
            pv * f + pmt * (1.0 + r * t) * (f - 1.0) / r + fv
        }
    };
    num(solve(guess, balance)?)
}

/// SLN(cost, salvage, life): straight-line depreciation per period.
pub fn sln(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (cost, salvage, life) = (args[0].number()?, args[1].number()?, args[2].number()?);
    if life == 0.0 {
        return Err(ErrorKind::Div0);
    }
    num((cost - salvage) / life)
}

/// DB(cost, salvage, life, period, [month]): fixed-declining balance, with
/// the rate rounded to three decimals as Excel does.
pub fn db(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (cost, salvage, life, period) = (args[0].number()?, args[1].number()?, args[2].number()?, args[3].number()?);
    let month = opt(args, 4, 12.0)?.trunc();
    let period = period.trunc();
    let last = if month < 12.0 { life + 1.0 } else { life };
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period < 1.0 || period > last || !(1.0..=12.0).contains(&month) {
        return Err(ErrorKind::Num);
    }
    if cost == 0.0 {
        return Ok(Value::Number(0.0));
    }
    let rate = ((1.0 - (salvage / cost).powf(1.0 / life)) * 1000.0).round() / 1000.0;
    let mut total = cost * rate * month / 12.0;
    if period == 1.0 {
        return num(total);
    }
    let mut dep = 0.0;
    for p in 2..=period as u32 {
        dep =
            if p as f64 == life + 1.0 { (cost - total) * rate * (12.0 - month) / 12.0 } else { (cost - total) * rate };
        total += dep;
    }
    num(dep)
}

/// DDB(cost, salvage, life, period, [factor]): double-declining balance.
pub fn ddb(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (cost, salvage, life, period) = (args[0].number()?, args[1].number()?, args[2].number()?, args[3].number()?);
    let factor = opt(args, 4, 2.0)?;
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || period > life || factor <= 0.0 {
        return Err(ErrorKind::Num);
    }
    let mut rate = factor / life;
    let old = if rate >= 1.0 {
        rate = 1.0;
        if period == 1.0 {
            cost
        } else {
            0.0
        }
    } else {
        cost * (1.0 - rate).powf(period - 1.0)
    };
    let new = cost * (1.0 - rate).powf(period);
    let dep = if new < salvage { old - salvage } else { old - new };
    num(dep.max(0.0))
}

/// EFFECT(nominal_rate, npery)
pub fn effect(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (rate, n) = (args[0].number()?, args[1].number()?.trunc());
    if rate <= 0.0 || n < 1.0 {
        return Err(ErrorKind::Num);
    }
    num((1.0 + rate / n).powf(n) - 1.0)
}

/// NOMINAL(effect_rate, npery)
pub fn nominal(args: &[Arg]) -> Result<Value, ErrorKind> {
    let (rate, n) = (args[0].number()?, args[1].number()?.trunc());
    if rate <= 0.0 || n < 1.0 {
        return Err(ErrorKind::Num);
    }
    num(n * ((1.0 + rate).powf(1.0 / n) - 1.0))
}

#[cfg(test)]
mod tests {
    use crate::{CellRef, ErrorKind, Sheet, Value};

    /// Evaluate `formula` on a sheet holding the given column A and B values.
    fn run_with(a: &[&str], b: &[&str], formula: &str) -> Value {
        let mut s = Sheet::new();
        let mut edits: Vec<(CellRef, String)> = Vec::new();
        for (i, v) in a.iter().enumerate() {
            edits.push((CellRef::new(i as u32, 0), v.to_string()));
        }
        for (i, v) in b.iter().enumerate() {
            edits.push((CellRef::new(i as u32, 1), v.to_string()));
        }
        edits.push((CellRef::new(0, 5), format!("={formula}")));
        s.set_inputs(edits);
        s.value(CellRef::new(0, 5)).clone()
    }

    fn approx_with(a: &[&str], b: &[&str], formula: &str, expected: f64, tolerance: f64) {
        match run_with(a, b, formula) {
            Value::Number(x) => {
                assert!((x - expected).abs() < tolerance, "{formula} = {x}, expected {expected}")
            }
            other => panic!("{formula} = {other:?}"),
        }
    }

    fn approx(formula: &str, expected: f64) {
        approx_with(&[], &[], formula, expected, 0.005);
    }

    #[test]
    fn loans() {
        approx("PMT(0.08/12, 10, 10000)", -1037.03);
        approx("PMT(0.08/12, 10, 10000, 0, 1)", -1030.16);
        approx("PMT(0, 10, 1000)", -100.0);
        approx("FV(0.06/12, 10, -200, -500, 1)", 2581.40);
        approx("FV(0.12/12, 12, -1000)", 12682.50);
        approx("PV(0.08/12, 12*20, 500, 0, 0)", -59777.15);
        // Microsoft's documented example: 59.6738657 with payments at the start (type 1).
        approx("NPER(0.12/12, -100, -1000, 10000, 1)", 59.6739);
        approx("NPER(0.12/12, -100, -1000, 10000)", 60.0821);
        approx("NPER(0.12/12, -100, -1000)", -9.578);
        approx("IPMT(0.1/12, 1, 3*12, 8000)", -66.67);
        approx("IPMT(0.1, 3, 3, 8000)", -292.45);
        approx("PPMT(0.1/12, 1, 2*12, 2000)", -75.62);
        approx("PPMT(0.08, 10, 10, 200000)", -27598.05);
        approx("RATE(4*12, -200, 8000)", 0.00770147);
        assert_eq!(run_with(&[], &[], "IPMT(0.1, 0, 3, 8000)"), Value::Error(ErrorKind::Num));
        assert_eq!(run_with(&[], &[], "PMT(0.1, 0, 100)"), Value::Error(ErrorKind::Num));
    }

    #[test]
    fn cash_flows() {
        approx("NPV(0.1, -10000, 3000, 4200, 6800)", 1188.44);
        let flows = ["-70000", "12000", "15000", "18000", "21000", "26000"];
        approx_with(&flows[..5], &[], "IRR(A1:A5)", -0.021245, 1e-5);
        approx_with(&flows, &[], "IRR(A1:A6)", 0.086631, 1e-5);
        approx_with(&flows[..3], &[], "IRR(A1:A3, -0.1)", -0.443507, 1e-5);
        assert_eq!(run_with(&["1", "2"], &[], "IRR(A1:A2)"), Value::Error(ErrorKind::Num));
        let values = ["-10000", "2750", "4250", "3250", "2750"];
        let dates = ["=DATE(2008,1,1)", "=DATE(2008,3,1)", "=DATE(2008,10,30)", "=DATE(2009,2,15)", "=DATE(2009,4,1)"];
        approx_with(&values, &dates, "XNPV(0.09, A1:A5, B1:B5)", 2086.65, 0.01);
        approx_with(&values, &dates, "XIRR(A1:A5, B1:B5)", 0.373362535, 1e-6);
        // Different numbers of values and dates.
        assert_eq!(run_with(&values, &dates, "XNPV(0.09, A1:A5, B1:B4)"), Value::Error(ErrorKind::Num));
    }

    #[test]
    fn depreciation_and_rates() {
        approx("SLN(30000, 7500, 10)", 2250.0);
        approx("DB(1000000, 100000, 6, 1, 7)", 186083.33);
        approx("DB(1000000, 100000, 6, 2, 7)", 259639.42);
        approx("DB(1000000, 100000, 6, 7, 7)", 15845.10);
        approx("DDB(2400, 300, 10*365, 1)", 1.32);
        approx("DDB(2400, 300, 10*12, 1, 2)", 40.0);
        approx("DDB(2400, 300, 10, 1, 2)", 480.0);
        approx("DDB(2400, 300, 10, 2, 1.5)", 306.0);
        approx("DDB(2400, 300, 10, 10)", 22.12);
        approx_with(&[], &[], "EFFECT(0.0525, 4)", 0.0535427, 1e-6);
        approx_with(&[], &[], "NOMINAL(0.053543, 4)", 0.0525003, 1e-6);
        assert_eq!(run_with(&[], &[], "EFFECT(0, 4)"), Value::Error(ErrorKind::Num));
        assert_eq!(run_with(&[], &[], "SLN(1, 0, 0)"), Value::Error(ErrorKind::Div0));
        assert_eq!(run_with(&[], &[], "DDB(2400, 300, 10, 11)"), Value::Error(ErrorKind::Num));
    }
}
