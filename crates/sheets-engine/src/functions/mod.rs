//! Built-in spreadsheet functions.
//!
//! ## Adding a function
//!
//! 1. Write `fn my_func(args: &[Arg]) -> Result<Value, ErrorKind>` in the
//!    module that fits (math, stats, text, logical, lookup, conditional,
//!    datetime, info).
//! 2. Register it in [`FUNCTIONS`] below with its argument limits, syntax and
//!    a one-sentence summary (both are shown by formula autocomplete).
//! 3. Add a test in that module and a row in `crates/sheets-engine/FUNCTIONS.md`
//!    (a test fails if a registered function is missing from that file).

mod conditional;
mod datetime;
mod financial;
mod format;
mod info;
mod logical;
mod lookup;
mod math;
mod reference;
mod stats;
mod text;

use crate::value::{ErrorKind, Value};

/// A rectangular block of values from a range argument, row-major.
#[derive(Debug, Clone, PartialEq)]
pub struct RangeValues {
    pub rows: u32,
    pub cols: u32,
    pub values: Vec<Value>,
}

impl RangeValues {
    pub fn get(&self, row: u32, col: u32) -> Option<&Value> {
        if row < self.rows && col < self.cols {
            self.values.get((row * self.cols + col) as usize)
        } else {
            None
        }
    }
}

/// An evaluated function argument.
#[derive(Debug, Clone, PartialEq)]
pub enum Arg {
    Value(Value),
    Range(RangeValues),
}

impl Arg {
    /// The argument as one value. A 1x1 range is fine; a bigger one is `#VALUE!`.
    pub fn scalar(&self) -> Value {
        match self {
            Arg::Value(v) => v.clone(),
            Arg::Range(r) if r.values.len() == 1 => r.values[0].clone(),
            Arg::Range(_) => Value::Error(ErrorKind::Value),
        }
    }

    pub fn number(&self) -> Result<f64, ErrorKind> {
        self.scalar().as_number()
    }

    pub fn text(&self) -> Result<String, ErrorKind> {
        self.scalar().as_text()
    }

    pub fn boolean(&self) -> Result<bool, ErrorKind> {
        self.scalar().as_bool()
    }

    /// Treat a scalar like a 1x1 range so lookup functions accept both.
    pub fn as_range(&self) -> RangeValues {
        match self {
            Arg::Range(r) => r.clone(),
            Arg::Value(v) => RangeValues { rows: 1, cols: 1, values: vec![v.clone()] },
        }
    }
}

/// Visit every number the way SUM/AVERAGE/MIN/MAX do: scalar arguments are
/// coerced (`SUM("3")` is 3), while text and booleans inside ranges are skipped.
/// Errors anywhere are returned.
pub fn for_each_number(args: &[Arg], mut f: impl FnMut(f64)) -> Result<(), ErrorKind> {
    for arg in args {
        match arg {
            Arg::Value(v) => f(v.as_number()?),
            Arg::Range(r) => {
                for v in &r.values {
                    match v {
                        Value::Number(n) => f(*n),
                        Value::Error(e) => return Err(*e),
                        _ => {}
                    }
                }
            }
        }
    }
    Ok(())
}

/// Every value in all arguments, flattened.
pub fn all_values(args: &[Arg]) -> impl Iterator<Item = &Value> {
    args.iter().flat_map(|a| match a {
        Arg::Value(v) => std::slice::from_ref(v).iter(),
        Arg::Range(r) => r.values.iter(),
    })
}

pub type FnImpl = fn(&[Arg]) -> Result<Value, ErrorKind>;

pub struct Function {
    pub name: &'static str,
    pub min_args: usize,
    /// `None` means any number of arguments.
    pub max_args: Option<usize>,
    pub call: FnImpl,
    /// Signature shown by formula autocomplete, e.g. `SUM(number1, [number2], ...)`.
    pub syntax: &'static str,
    /// One short sentence describing what the function does.
    pub summary: &'static str,
}

const fn f(
    name: &'static str,
    min_args: usize,
    max_args: Option<usize>,
    call: FnImpl,
    syntax: &'static str,
    summary: &'static str,
) -> Function {
    Function { name, min_args, max_args, call, syntax, summary }
}

const ANY: Option<usize> = None;

/// Registry of every built-in function, sorted by byte order of the name
/// (so `STDEV` < `STDEV.P` < `STDEVP`); [`lookup`] relies on that.
pub static FUNCTIONS: &[Function] = &[
    f("ABS", 1, Some(1), math::abs, "ABS(number)", "Absolute value of a number."),
    f(
        "ADDRESS",
        2,
        Some(5),
        reference::address,
        "ADDRESS(row, column, [abs_num], [a1], [sheet_text])",
        "A cell reference as text, like $A$1.",
    ),
    f(
        "AGGREGATE",
        3,
        ANY,
        reference::special,
        "AGGREGATE(function_num, options, ref1, ...)",
        "A total that can skip errors, hidden rows and nested subtotals.",
    ),
    f("AND", 1, ANY, logical::and, "AND(logical1, [logical2], ...)", "TRUE if every argument is TRUE."),
    f("AVERAGE", 1, ANY, math::average, "AVERAGE(number1, [number2], ...)", "Arithmetic mean of the numbers."),
    f(
        "AVERAGEIF",
        2,
        Some(3),
        conditional::averageif,
        "AVERAGEIF(range, criteria, [average_range])",
        "Mean of the cells that meet a condition.",
    ),
    f(
        "AVERAGEIFS",
        3,
        ANY,
        conditional::averageifs,
        "AVERAGEIFS(average_range, criteria_range1, criteria1, ...)",
        "Mean of the cells that meet every condition.",
    ),
    f(
        "CEILING",
        2,
        Some(2),
        math::ceiling,
        "CEILING(number, significance)",
        "Rounds up to the nearest multiple of significance.",
    ),
    f("CHAR", 1, Some(1), text::char, "CHAR(number)", "Character for a code from 1 to 255."),
    f(
        "CHOOSE",
        2,
        ANY,
        lookup::choose,
        "CHOOSE(index, value1, [value2], ...)",
        "Picks a value from the list by position.",
    ),
    f(
        "CHOOSECOLS",
        2,
        ANY,
        reference::special,
        "CHOOSECOLS(array, col_num1, [col_num2], ...)",
        "The chosen columns of an array.",
    ),
    f(
        "CHOOSEROWS",
        2,
        ANY,
        reference::special,
        "CHOOSEROWS(array, row_num1, [row_num2], ...)",
        "The chosen rows of an array.",
    ),
    f("CLEAN", 1, Some(1), text::clean, "CLEAN(text)", "Removes non-printable characters."),
    f("CODE", 1, Some(1), text::code, "CODE(text)", "Numeric code of the first character."),
    f(
        "COLUMN",
        0,
        Some(1),
        reference::special,
        "COLUMN([reference])",
        "Column number of a reference, or of this cell.",
    ),
    f("COLUMNS", 1, Some(1), reference::special, "COLUMNS(array)", "Number of columns in a reference."),
    f("CONCAT", 1, ANY, text::concat, "CONCAT(text1, [text2], ...)", "Joins text, including whole ranges."),
    f("CONCATENATE", 1, ANY, text::concat, "CONCATENATE(text1, [text2], ...)", "Joins text."),
    f("COUNT", 1, ANY, math::count, "COUNT(value1, [value2], ...)", "Counts numbers."),
    f("COUNTA", 1, ANY, math::counta, "COUNTA(value1, [value2], ...)", "Counts non-empty cells."),
    f("COUNTBLANK", 1, Some(1), math::countblank, "COUNTBLANK(range)", "Counts empty cells."),
    f("COUNTIF", 2, Some(2), conditional::countif, "COUNTIF(range, criteria)", "Counts cells that meet a condition."),
    f(
        "COUNTIFS",
        2,
        ANY,
        conditional::countifs,
        "COUNTIFS(criteria_range1, criteria1, ...)",
        "Counts cells that meet every condition.",
    ),
    f("DATE", 3, Some(3), datetime::date, "DATE(year, month, day)", "Date serial number from year, month and day."),
    f(
        "DATEDIF",
        3,
        Some(3),
        datetime::datedif,
        "DATEDIF(start_date, end_date, unit)",
        "Difference between two dates in years, months or days.",
    ),
    f("DATEVALUE", 1, Some(1), datetime::datevalue, "DATEVALUE(date_text)", "Converts date text to a serial number."),
    f("DAY", 1, Some(1), datetime::day, "DAY(date)", "Day of the month, 1 to 31."),
    f("DAYS", 2, Some(2), datetime::days, "DAYS(end_date, start_date)", "Number of days between two dates."),
    f(
        "DB",
        4,
        Some(5),
        financial::db,
        "DB(cost, salvage, life, period, [month])",
        "Depreciation by the fixed-declining balance method.",
    ),
    f(
        "DDB",
        4,
        Some(5),
        financial::ddb,
        "DDB(cost, salvage, life, period, [factor])",
        "Depreciation by the double-declining balance method.",
    ),
    f(
        "DROP",
        2,
        Some(3),
        reference::special,
        "DROP(array, rows, [columns])",
        "An array without its first (or last, if negative) rows or columns.",
    ),
    f("EDATE", 2, Some(2), datetime::edate, "EDATE(start_date, months)", "Same day a number of months away."),
    f("EFFECT", 2, Some(2), financial::effect, "EFFECT(nominal_rate, npery)", "Effective annual interest rate."),
    f(
        "EOMONTH",
        2,
        Some(2),
        datetime::eomonth,
        "EOMONTH(start_date, months)",
        "Last day of the month a number of months away.",
    ),
    f("EVEN", 1, Some(1), math::even, "EVEN(number)", "Rounds away from zero to an even integer."),
    f("EXACT", 2, Some(2), text::exact, "EXACT(text1, text2)", "TRUE if two texts are identical, case-sensitive."),
    f("EXP", 1, Some(1), math::exp, "EXP(number)", "e raised to a power."),
    f(
        "EXPAND",
        2,
        Some(4),
        reference::special,
        "EXPAND(array, rows, [columns], [pad_with])",
        "An array grown to the given size.",
    ),
    f("FACT", 1, Some(1), math::fact, "FACT(number)", "Factorial of a number."),
    f("FALSE", 0, Some(0), logical::false_, "FALSE()", "The logical value FALSE."),
    f(
        "FILTER",
        2,
        Some(3),
        reference::special,
        "FILTER(array, include, [if_empty])",
        "The rows (or columns) of an array that meet a condition. Spills.",
    ),
    f(
        "FIND",
        2,
        Some(3),
        text::find,
        "FIND(find_text, within_text, [start])",
        "Position of text within text, case-sensitive.",
    ),
    f(
        "FLOOR",
        2,
        Some(2),
        math::floor,
        "FLOOR(number, significance)",
        "Rounds down to the nearest multiple of significance.",
    ),
    f("FV", 3, Some(5), financial::fv, "FV(rate, nper, pmt, [pv], [type])", "Future value of an investment."),
    f("GCD", 1, ANY, math::gcd, "GCD(number1, [number2], ...)", "Greatest common divisor."),
    f(
        "HLOOKUP",
        3,
        Some(4),
        lookup::hlookup,
        "HLOOKUP(value, table, row_index, [approximate])",
        "Looks up a value in the first row of a table.",
    ),
    f("HOUR", 1, Some(1), datetime::hour, "HOUR(time)", "Hour of a time, 0 to 23."),
    f("HSTACK", 1, ANY, reference::special, "HSTACK(array1, [array2], ...)", "Arrays placed side by side."),
    f(
        "IF",
        2,
        Some(3),
        logical::if_,
        "IF(condition, value_if_true, [value_if_false])",
        "Chooses a value based on a condition.",
    ),
    f(
        "IFERROR",
        2,
        Some(2),
        logical::iferror,
        "IFERROR(value, value_if_error)",
        "Replaces any error with another value.",
    ),
    f("IFNA", 2, Some(2), logical::ifna, "IFNA(value, value_if_na)", "Replaces #N/A with another value."),
    f(
        "IFS",
        2,
        ANY,
        logical::ifs,
        "IFS(condition1, value1, [condition2, value2], ...)",
        "Value for the first condition that is TRUE.",
    ),
    f("INDEX", 2, Some(3), lookup::index, "INDEX(range, row, [col])", "Value at a row and column of a range."),
    f("INDIRECT", 1, Some(2), reference::special, "INDIRECT(ref_text, [a1])", "The reference named by a text string."),
    f("INT", 1, Some(1), math::int, "INT(number)", "Rounds down to an integer."),
    f(
        "IPMT",
        4,
        Some(6),
        financial::ipmt,
        "IPMT(rate, per, nper, pv, [fv], [type])",
        "Interest part of a loan payment.",
    ),
    f("IRR", 1, Some(2), financial::irr, "IRR(values, [guess])", "Internal rate of return of cash flows."),
    f("ISBLANK", 1, Some(1), info::isblank, "ISBLANK(value)", "TRUE if the cell is empty."),
    f("ISERR", 1, Some(1), info::iserr, "ISERR(value)", "TRUE for any error except #N/A."),
    f("ISERROR", 1, Some(1), info::iserror, "ISERROR(value)", "TRUE for any error."),
    f("ISEVEN", 1, Some(1), info::iseven, "ISEVEN(number)", "TRUE if the number is even."),
    f("ISLOGICAL", 1, Some(1), info::islogical, "ISLOGICAL(value)", "TRUE if the value is TRUE or FALSE."),
    f("ISNA", 1, Some(1), info::isna, "ISNA(value)", "TRUE if the value is #N/A."),
    f("ISNONTEXT", 1, Some(1), info::isnontext, "ISNONTEXT(value)", "TRUE if the value is not text."),
    f("ISNUMBER", 1, Some(1), info::isnumber, "ISNUMBER(value)", "TRUE if the value is a number."),
    f("ISODD", 1, Some(1), info::isodd, "ISODD(number)", "TRUE if the number is odd."),
    f("ISTEXT", 1, Some(1), info::istext, "ISTEXT(value)", "TRUE if the value is text."),
    f("LARGE", 2, Some(2), stats::large, "LARGE(data, k)", "The k-th largest number."),
    f("LCM", 1, ANY, math::lcm, "LCM(number1, [number2], ...)", "Least common multiple."),
    f("LEFT", 1, Some(2), text::left, "LEFT(text, [count])", "First characters of a text."),
    f("LEN", 1, Some(1), text::len, "LEN(text)", "Number of characters in a text."),
    f(
        "LET",
        3,
        ANY,
        reference::special,
        "LET(name1, value1, [name2, value2, ...], calculation)",
        "Names values for use in a calculation.",
    ),
    f("LN", 1, Some(1), math::ln, "LN(number)", "Natural logarithm."),
    f("LOG", 1, Some(2), math::log, "LOG(number, [base])", "Logarithm to a base, 10 by default."),
    f("LOG10", 1, Some(1), math::log10, "LOG10(number)", "Base-10 logarithm."),
    f("LOWER", 1, Some(1), text::lower, "LOWER(text)", "Converts text to lower case."),
    f("MATCH", 2, Some(3), lookup::match_, "MATCH(value, range, [type])", "Position of a value in a range."),
    f("MAX", 1, ANY, math::max, "MAX(number1, [number2], ...)", "Largest number."),
    f(
        "MAXIFS",
        3,
        ANY,
        conditional::maxifs,
        "MAXIFS(max_range, criteria_range1, criteria1, ...)",
        "Largest number among cells that meet every condition.",
    ),
    f("MEDIAN", 1, ANY, stats::median, "MEDIAN(number1, [number2], ...)", "Middle number of the values."),
    f("MID", 3, Some(3), text::mid, "MID(text, start, count)", "Characters from the middle of a text."),
    f("MIN", 1, ANY, math::min, "MIN(number1, [number2], ...)", "Smallest number."),
    f(
        "MINIFS",
        3,
        ANY,
        conditional::minifs,
        "MINIFS(min_range, criteria_range1, criteria1, ...)",
        "Smallest number among cells that meet every condition.",
    ),
    f("MINUTE", 1, Some(1), datetime::minute, "MINUTE(time)", "Minute of a time, 0 to 59."),
    f("MOD", 2, Some(2), math::mod_, "MOD(number, divisor)", "Remainder after division."),
    f("MODE", 1, ANY, stats::mode, "MODE(number1, [number2], ...)", "Most frequent number."),
    f("MODE.SNGL", 1, ANY, stats::mode, "MODE.SNGL(number1, [number2], ...)", "Most frequent number."),
    f("MONTH", 1, Some(1), datetime::month, "MONTH(date)", "Month of a date, 1 to 12."),
    f("N", 1, Some(1), text::n, "N(value)", "Converts a value to a number; text becomes 0."),
    f("NA", 0, Some(0), info::na, "NA()", "The #N/A error."),
    f(
        "NETWORKDAYS",
        2,
        Some(3),
        datetime::networkdays,
        "NETWORKDAYS(start_date, end_date, [holidays])",
        "Working days between two dates, inclusive.",
    ),
    f("NOMINAL", 2, Some(2), financial::nominal, "NOMINAL(effect_rate, npery)", "Nominal annual interest rate."),
    f("NOT", 1, Some(1), logical::not, "NOT(logical)", "Reverses TRUE and FALSE."),
    f("NOW", 0, Some(0), datetime::now, "NOW()", "Current date and time."),
    f("NPER", 3, Some(5), financial::nper, "NPER(rate, pmt, pv, [fv], [type])", "Number of periods for an investment."),
    f("NPV", 2, ANY, financial::npv, "NPV(rate, value1, [value2], ...)", "Net present value of cash flows."),
    f("ODD", 1, Some(1), math::odd, "ODD(number)", "Rounds away from zero to an odd integer."),
    f(
        "OFFSET",
        3,
        Some(5),
        reference::special,
        "OFFSET(reference, rows, cols, [height], [width])",
        "A reference shifted from a starting reference.",
    ),
    f("OR", 1, ANY, logical::or, "OR(logical1, [logical2], ...)", "TRUE if any argument is TRUE."),
    f("PERCENTILE", 2, Some(2), stats::percentile_inc, "PERCENTILE(array, k)", "The k-th percentile of values."),
    f(
        "PERCENTILE.EXC",
        2,
        Some(2),
        stats::percentile_exc,
        "PERCENTILE.EXC(array, k)",
        "The k-th percentile, k strictly between 0 and 1.",
    ),
    f(
        "PERCENTILE.INC",
        2,
        Some(2),
        stats::percentile_inc,
        "PERCENTILE.INC(array, k)",
        "The k-th percentile of values.",
    ),
    f("PI", 0, Some(0), math::pi, "PI()", "The number pi."),
    f("PMT", 3, Some(5), financial::pmt, "PMT(rate, nper, pv, [fv], [type])", "Payment for a loan."),
    f("POWER", 2, Some(2), math::power, "POWER(number, power)", "A number raised to a power."),
    f(
        "PPMT",
        4,
        Some(6),
        financial::ppmt,
        "PPMT(rate, per, nper, pv, [fv], [type])",
        "Principal part of a loan payment.",
    ),
    f("PRODUCT", 1, ANY, math::product, "PRODUCT(number1, [number2], ...)", "Multiplies numbers."),
    f("PROPER", 1, Some(1), text::proper, "PROPER(text)", "Capitalizes the first letter of each word."),
    f("PV", 3, Some(5), financial::pv, "PV(rate, nper, pmt, [fv], [type])", "Present value of an investment."),
    f("QUARTILE", 2, Some(2), stats::quartile_inc, "QUARTILE(array, quart)", "A quartile of values."),
    f("QUARTILE.EXC", 2, Some(2), stats::quartile_exc, "QUARTILE.EXC(array, quart)", "A quartile, excluding 0 and 1."),
    f("QUARTILE.INC", 2, Some(2), stats::quartile_inc, "QUARTILE.INC(array, quart)", "A quartile of values."),
    f("RAND", 0, Some(0), math::rand, "RAND()", "Random number from 0 up to 1."),
    f(
        "RANDARRAY",
        0,
        Some(5),
        reference::special,
        "RANDARRAY([rows], [columns], [min], [max], [whole_number])",
        "An array of random numbers. Spills.",
    ),
    f("RANDBETWEEN", 2, Some(2), math::randbetween, "RANDBETWEEN(bottom, top)", "Random integer between two numbers."),
    f("RANK", 2, Some(3), stats::rank, "RANK(number, ref, [order])", "Rank of a number in a list."),
    f("RANK.EQ", 2, Some(3), stats::rank, "RANK.EQ(number, ref, [order])", "Rank of a number in a list."),
    f("RATE", 3, Some(6), financial::rate, "RATE(nper, pmt, pv, [fv], [type], [guess])", "Interest rate per period."),
    f(
        "REPLACE",
        4,
        Some(4),
        text::replace,
        "REPLACE(old_text, start, count, new_text)",
        "Replaces characters at a position.",
    ),
    f("REPT", 2, Some(2), text::rept, "REPT(text, times)", "Repeats text."),
    f("RIGHT", 1, Some(2), text::right, "RIGHT(text, [count])", "Last characters of a text."),
    f("ROUND", 2, Some(2), math::round, "ROUND(number, digits)", "Rounds to a number of digits."),
    f("ROUNDDOWN", 2, Some(2), math::rounddown, "ROUNDDOWN(number, digits)", "Rounds toward zero."),
    f("ROUNDUP", 2, Some(2), math::roundup, "ROUNDUP(number, digits)", "Rounds away from zero."),
    f("ROW", 0, Some(1), reference::special, "ROW([reference])", "Row number of a reference, or of this cell."),
    f("ROWS", 1, Some(1), reference::special, "ROWS(array)", "Number of rows in a reference."),
    f(
        "SEARCH",
        2,
        Some(3),
        text::search,
        "SEARCH(find_text, within_text, [start])",
        "Position of text within text, ignoring case; supports wildcards.",
    ),
    f("SECOND", 1, Some(1), datetime::second, "SECOND(time)", "Second of a time, 0 to 59."),
    f(
        "SEQUENCE",
        1,
        Some(4),
        reference::special,
        "SEQUENCE(rows, [columns], [start], [step])",
        "An array of sequential numbers. Spills.",
    ),
    f("SIGN", 1, Some(1), math::sign, "SIGN(number)", "1, 0 or -1 depending on the sign."),
    f("SLN", 3, Some(3), financial::sln, "SLN(cost, salvage, life)", "Straight-line depreciation for one period."),
    f("SMALL", 2, Some(2), stats::small, "SMALL(data, k)", "The k-th smallest number."),
    f(
        "SORT",
        1,
        Some(4),
        reference::special,
        "SORT(array, [sort_index], [sort_order], [by_col])",
        "An array sorted by one of its columns. Spills.",
    ),
    f(
        "SORTBY",
        2,
        ANY,
        reference::special,
        "SORTBY(array, by_array1, [sort_order1], ...)",
        "An array sorted by other arrays. Spills.",
    ),
    f("SQRT", 1, Some(1), math::sqrt, "SQRT(number)", "Square root."),
    f("STDEV", 1, ANY, stats::stdev, "STDEV(number1, [number2], ...)", "Standard deviation of a sample."),
    f(
        "STDEV.P",
        1,
        ANY,
        stats::stdevp,
        "STDEV.P(number1, [number2], ...)",
        "Standard deviation of a whole population.",
    ),
    f("STDEV.S", 1, ANY, stats::stdev, "STDEV.S(number1, [number2], ...)", "Standard deviation of a sample."),
    f("STDEVP", 1, ANY, stats::stdevp, "STDEVP(number1, [number2], ...)", "Standard deviation of a whole population."),
    f(
        "SUBSTITUTE",
        3,
        Some(4),
        text::substitute,
        "SUBSTITUTE(text, old_text, new_text, [instance])",
        "Replaces occurrences of a text.",
    ),
    f(
        "SUBTOTAL",
        2,
        ANY,
        reference::special,
        "SUBTOTAL(function_num, ref1, ...)",
        "A subtotal that skips nested subtotals (and hidden rows for 101-111).",
    ),
    f("SUM", 1, ANY, math::sum, "SUM(number1, [number2], ...)", "Adds numbers."),
    f(
        "SUMIF",
        2,
        Some(3),
        conditional::sumif,
        "SUMIF(range, criteria, [sum_range])",
        "Adds the cells that meet a condition.",
    ),
    f(
        "SUMIFS",
        3,
        ANY,
        conditional::sumifs,
        "SUMIFS(sum_range, criteria_range1, criteria1, ...)",
        "Adds the cells that meet every condition.",
    ),
    f(
        "SUMPRODUCT",
        1,
        ANY,
        math::sumproduct,
        "SUMPRODUCT(array1, [array2], ...)",
        "Sum of the products of matching items.",
    ),
    f(
        "SWITCH",
        3,
        ANY,
        logical::switch,
        "SWITCH(expression, value1, result1, [value2, result2], ..., [default])",
        "Result for the first value equal to the expression.",
    ),
    f("T", 1, Some(1), text::t, "T(value)", "The value if it is text, otherwise empty text."),
    f(
        "TAKE",
        2,
        Some(3),
        reference::special,
        "TAKE(array, rows, [columns])",
        "The first (or last, if negative) rows or columns of an array.",
    ),
    f("TEXT", 2, Some(2), text::text, "TEXT(value, format)", "Formats a number as text using a format code."),
    f(
        "TEXTJOIN",
        3,
        ANY,
        text::textjoin,
        "TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)",
        "Joins text with a delimiter.",
    ),
    f(
        "TEXTSPLIT",
        2,
        Some(4),
        reference::special,
        "TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty])",
        "Text split into columns and rows. Spills.",
    ),
    f(
        "TIME",
        3,
        Some(3),
        datetime::time,
        "TIME(hour, minute, second)",
        "Time serial number from hours, minutes and seconds.",
    ),
    f("TOCOL", 1, Some(3), reference::special, "TOCOL(array, [ignore], [scan_by_column])", "An array as one column."),
    f("TODAY", 0, Some(0), datetime::today, "TODAY()", "Current date."),
    f("TOROW", 1, Some(3), reference::special, "TOROW(array, [ignore], [scan_by_column])", "An array as one row."),
    f(
        "TRANSPOSE",
        1,
        Some(1),
        reference::special,
        "TRANSPOSE(array)",
        "Rows turned into columns and columns into rows. Spills.",
    ),
    f("TRIM", 1, Some(1), text::trim, "TRIM(text)", "Removes extra spaces."),
    f("TRUE", 0, Some(0), logical::true_, "TRUE()", "The logical value TRUE."),
    f("TRUNC", 1, Some(2), math::trunc, "TRUNC(number, [digits])", "Cuts off the fractional part."),
    f(
        "UNIQUE",
        1,
        Some(3),
        reference::special,
        "UNIQUE(array, [by_col], [exactly_once])",
        "The distinct rows of an array. Spills.",
    ),
    f("UPPER", 1, Some(1), text::upper, "UPPER(text)", "Converts text to upper case."),
    f("VALUE", 1, Some(1), text::value, "VALUE(text)", "Converts text to a number."),
    f("VAR", 1, ANY, stats::var, "VAR(number1, [number2], ...)", "Variance of a sample."),
    f("VAR.P", 1, ANY, stats::varp, "VAR.P(number1, [number2], ...)", "Variance of a whole population."),
    f("VAR.S", 1, ANY, stats::var, "VAR.S(number1, [number2], ...)", "Variance of a sample."),
    f("VARP", 1, ANY, stats::varp, "VARP(number1, [number2], ...)", "Variance of a whole population."),
    f(
        "VLOOKUP",
        3,
        Some(4),
        lookup::vlookup,
        "VLOOKUP(value, table, col_index, [approximate])",
        "Looks up a value in the first column of a table.",
    ),
    f("VSTACK", 1, ANY, reference::special, "VSTACK(array1, [array2], ...)", "Arrays placed one below the other."),
    f("WEEKDAY", 1, Some(2), datetime::weekday, "WEEKDAY(date, [type])", "Day of the week as a number."),
    f("WEEKNUM", 1, Some(2), datetime::weeknum, "WEEKNUM(date, [type])", "Week number of the year."),
    f(
        "WORKDAY",
        2,
        Some(3),
        datetime::workday,
        "WORKDAY(start_date, days, [holidays])",
        "Date a number of working days away.",
    ),
    f(
        "XIRR",
        2,
        Some(3),
        financial::xirr,
        "XIRR(values, dates, [guess])",
        "Internal rate of return for irregular cash flows.",
    ),
    f(
        "XLOOKUP",
        3,
        Some(6),
        lookup::xlookup,
        "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
        "Finds a value and returns the matching item from another range.",
    ),
    f(
        "XMATCH",
        2,
        Some(4),
        reference::special,
        "XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])",
        "Position of a value in a row or column (exact, next smaller/larger or wildcard).",
    ),
    f("XNPV", 3, Some(3), financial::xnpv, "XNPV(rate, values, dates)", "Net present value for irregular cash flows."),
    f("XOR", 1, ANY, logical::xor, "XOR(logical1, [logical2], ...)", "TRUE if an odd number of arguments are TRUE."),
    f("YEAR", 1, Some(1), datetime::year, "YEAR(date)", "Year of a date."),
];

/// Find a function by its upper-case name.
/// A number in [0, 1) (RAND, RANDARRAY).
pub(crate) fn random() -> f64 {
    math::random_unit()
}

pub fn lookup(name: &str) -> Option<&'static Function> {
    FUNCTIONS.binary_search_by(|f| f.name.cmp(name)).ok().map(|i| &FUNCTIONS[i])
}

#[cfg(test)]
pub(crate) mod test_util {
    use crate::formula::{eval, parse, CellSource};
    use crate::{CellRef, Value};

    /// A1:A5 = 1..5, B1:B5 = a..e, C1 = TRUE, D1 = "", everything else empty.
    pub struct Fixture;

    impl CellSource for Fixture {
        fn value(&self, at: CellRef) -> Value {
            match (at.col, at.row) {
                (0, r @ 0..=4) => Value::Number(r as f64 + 1.0),
                (1, r @ 0..=4) => Value::Text(((b'a' + r as u8) as char).to_string()),
                (2, 0) => Value::Bool(true),
                (3, 0) => Value::Text(String::new()),
                _ => Value::Empty,
            }
        }
    }

    pub fn run(formula: &str) -> Value {
        eval(&parse(formula).unwrap_or_else(|e| panic!("{formula}: {e:?}")), &Fixture)
    }

    pub fn n(x: f64) -> Value {
        Value::Number(x)
    }

    pub fn t(s: &str) -> Value {
        Value::Text(s.into())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn registry_is_sorted_and_unique() {
        for pair in FUNCTIONS.windows(2) {
            assert!(pair[0].name < pair[1].name, "{} must come before {}", pair[0].name, pair[1].name);
        }
    }

    #[test]
    fn every_function_has_metadata() {
        for f in FUNCTIONS {
            assert!(f.syntax.starts_with(&format!("{}(", f.name)) && f.syntax.ends_with(')'), "{} syntax", f.name);
            assert!(!f.summary.is_empty() && f.summary.ends_with('.'), "{} summary", f.name);
            assert!(f.max_args.is_none_or(|max| max >= f.min_args), "{} argument limits", f.name);
            assert!(std::ptr::eq(lookup(f.name).unwrap(), f), "{} lookup", f.name);
        }
        assert!(lookup("STDEV.S").is_some() && lookup("NOPE").is_none());
    }

    #[test]
    fn every_function_is_documented() {
        let docs = include_str!("../../FUNCTIONS.md");
        for f in FUNCTIONS {
            assert!(docs.contains(&format!("`{}`", f.name)), "{} is missing from FUNCTIONS.md", f.name);
        }
    }
}
