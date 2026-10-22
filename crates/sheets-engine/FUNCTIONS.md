# Supported functions

Every function registered in [`src/functions/mod.rs`](src/functions/mod.rs) must
be listed here — `cargo test` fails otherwise. Behaviour follows Microsoft Excel
unless a note says otherwise. Each registry entry also carries the syntax and a
one-sentence summary used by formula autocomplete.

## Math

| Function | Syntax | Notes |
|---|---|---|
| `SUM` | `SUM(value1, ...)` | Text and booleans inside ranges are ignored |
| `PRODUCT` | `PRODUCT(value1, ...)` | |
| `SUMPRODUCT` | `SUMPRODUCT(array1, [array2], ...)` | Arrays must be the same shape (`#VALUE!` otherwise); non-numbers count as 0 |
| `AVERAGE` | `AVERAGE(value1, ...)` | `#DIV/0!` when there are no numbers |
| `MIN` | `MIN(value1, ...)` | 0 when there are no numbers |
| `MAX` | `MAX(value1, ...)` | 0 when there are no numbers |
| `COUNT` | `COUNT(value1, ...)` | Counts numbers |
| `COUNTA` | `COUNTA(value1, ...)` | Counts non-empty cells |
| `COUNTBLANK` | `COUNTBLANK(range)` | Empty cells and `""` |
| `ABS` | `ABS(number)` | |
| `SIGN` | `SIGN(number)` | |
| `INT` | `INT(number)` | Rounds down |
| `TRUNC` | `TRUNC(number, [digits])` | Toward zero |
| `MOD` | `MOD(number, divisor)` | Sign follows the divisor |
| `POWER` | `POWER(number, power)` | |
| `SQRT` | `SQRT(number)` | `#NUM!` for negatives |
| `EXP` | `EXP(number)` | |
| `LN` | `LN(number)` | `#NUM!` for numbers <= 0 |
| `LOG` | `LOG(number, [base])` | Base defaults to 10; base 1 is `#DIV/0!` |
| `LOG10` | `LOG10(number)` | |
| `PI` | `PI()` | |
| `ROUND` | `ROUND(number, digits)` | Half away from zero; negative digits allowed |
| `ROUNDUP` | `ROUNDUP(number, digits)` | Away from zero |
| `ROUNDDOWN` | `ROUNDDOWN(number, digits)` | Toward zero |
| `CEILING` | `CEILING(number, significance)` | Excel 2-argument form; positive number with negative significance is `#NUM!` |
| `FLOOR` | `FLOOR(number, significance)` | Excel 2-argument form; significance 0 is `#DIV/0!` |
| `EVEN` | `EVEN(number)` | Away from zero |
| `ODD` | `ODD(number)` | Away from zero; `ODD(0)` is 1 |
| `FACT` | `FACT(number)` | `#NUM!` for negatives |
| `GCD` | `GCD(number1, ...)` | Accepts ranges; `#NUM!` for negatives |
| `LCM` | `LCM(number1, ...)` | Accepts ranges; `#NUM!` for negatives |
| `RAND` | `RAND()` | See note on volatile functions below |
| `RANDBETWEEN` | `RANDBETWEEN(bottom, top)` | `#NUM!` when bottom > top |

## Statistics

Like `SUM`, these read only numbers from ranges and coerce scalar arguments.

| Function | Syntax | Notes |
|---|---|---|
| `MEDIAN` | `MEDIAN(number1, ...)` | `#NUM!` when there are no numbers |
| `MODE` | `MODE(number1, ...)` | First most frequent value; `#N/A` when nothing repeats |
| `MODE.SNGL` | `MODE.SNGL(number1, ...)` | Same as `MODE` |
| `LARGE` | `LARGE(data, k)` | `#NUM!` when k is out of range |
| `SMALL` | `SMALL(data, k)` | `#NUM!` when k is out of range |
| `RANK` | `RANK(number, ref, [order])` | order 0 (default) ranks the largest as 1; `#N/A` if number is not in ref |
| `RANK.EQ` | `RANK.EQ(number, ref, [order])` | Same as `RANK` |
| `STDEV` | `STDEV(number1, ...)` | Sample; `#DIV/0!` with fewer than 2 numbers |
| `STDEV.S` | `STDEV.S(number1, ...)` | Same as `STDEV` |
| `STDEVP` | `STDEVP(number1, ...)` | Population |
| `STDEV.P` | `STDEV.P(number1, ...)` | Same as `STDEVP` |
| `VAR` | `VAR(number1, ...)` | Sample |
| `VAR.S` | `VAR.S(number1, ...)` | Same as `VAR` |
| `VARP` | `VARP(number1, ...)` | Population |
| `VAR.P` | `VAR.P(number1, ...)` | Same as `VARP` |

## Conditional

Criteria can be a value (`5`, `"apple"`) or a string with an operator
(`">5"`, `"<>apple"`, `"<=10"`). Text supports wildcards `*`, `?` and `~`.
In the `*IFS` functions every criteria range must have the same shape as the
first range, otherwise the result is `#VALUE!`.

| Function | Syntax | Notes |
|---|---|---|
| `SUMIF` | `SUMIF(range, criteria, [sum_range])` | |
| `SUMIFS` | `SUMIFS(sum_range, criteria_range1, criteria1, ...)` | |
| `COUNTIF` | `COUNTIF(range, criteria)` | |
| `COUNTIFS` | `COUNTIFS(criteria_range1, criteria1, ...)` | |
| `AVERAGEIF` | `AVERAGEIF(range, criteria, [average_range])` | |
| `AVERAGEIFS` | `AVERAGEIFS(average_range, criteria_range1, criteria1, ...)` | `#DIV/0!` when nothing matches |
| `MAXIFS` | `MAXIFS(max_range, criteria_range1, criteria1, ...)` | 0 when nothing matches |
| `MINIFS` | `MINIFS(min_range, criteria_range1, criteria1, ...)` | 0 when nothing matches |

## Logical

| Function | Syntax | Notes |
|---|---|---|
| `IF` | `IF(condition, value_if_true, [value_if_false])` | |
| `IFS` | `IFS(condition1, value1, ...)` | `#N/A` when no condition is TRUE |
| `IFERROR` | `IFERROR(value, value_if_error)` | |
| `IFNA` | `IFNA(value, value_if_na)` | |
| `SWITCH` | `SWITCH(expression, value1, result1, ..., [default])` | Values match only if they have the same type; `#N/A` without a match or default |
| `AND` | `AND(logical1, ...)` | |
| `OR` | `OR(logical1, ...)` | |
| `XOR` | `XOR(logical1, ...)` | TRUE when an odd number of arguments are TRUE |
| `NOT` | `NOT(logical)` | |
| `TRUE` | `TRUE()` | |
| `FALSE` | `FALSE()` | |

All arguments are evaluated before the function runs (no short-circuiting), but
an error in a branch that is not chosen does not affect the result.

## Text

| Function | Syntax | Notes |
|---|---|---|
| `LEN` | `LEN(text)` | |
| `UPPER` | `UPPER(text)` | |
| `LOWER` | `LOWER(text)` | |
| `PROPER` | `PROPER(text)` | |
| `TRIM` | `TRIM(text)` | |
| `CLEAN` | `CLEAN(text)` | Removes characters 0–31 |
| `LEFT` | `LEFT(text, [count])` | |
| `RIGHT` | `RIGHT(text, [count])` | |
| `MID` | `MID(text, start, count)` | |
| `FIND` | `FIND(find_text, within_text, [start])` | Case-sensitive; `#VALUE!` if not found |
| `SEARCH` | `SEARCH(find_text, within_text, [start])` | Ignores case; wildcards `*`, `?`, `~` |
| `SUBSTITUTE` | `SUBSTITUTE(text, old_text, new_text, [instance])` | Case-sensitive |
| `REPLACE` | `REPLACE(old_text, start, count, new_text)` | |
| `REPT` | `REPT(text, times)` | `#VALUE!` beyond 32,767 characters |
| `CONCAT` | `CONCAT(text1, ...)` (accepts ranges) | |
| `CONCATENATE` | `CONCATENATE(text1, ...)` | |
| `TEXTJOIN` | `TEXTJOIN(delimiter, ignore_empty, text1, ...)` | |
| `EXACT` | `EXACT(text1, text2)` | Case-sensitive |
| `VALUE` | `VALUE(text)` | Understands `1,234.5`, `$12`, `50%`, `(3)`, ISO / `dd/mm/yyyy` dates and `hh:mm[:ss]` times |
| `TEXT` | `TEXT(value, format)` | See below |
| `CHAR` | `CHAR(number)` | Windows-1252 character set, 1–255 |
| `CODE` | `CODE(text)` | Windows-1252 code; 63 (`?`) for characters outside it |
| `T` | `T(value)` | |
| `N` | `N(value)` | |

`TEXT` supports `General`, `@`, digit placeholders `0` and `#`, `.`, thousands
`,` (and trailing `,` to divide by 1000), `%`, quoted literals, `\` escapes,
up to three `;` sections (positive;negative;zero) and date/time codes `d`,
`dd`, `ddd`, `dddd`, `m`, `mm`, `mmm`, `mmmm`, `yy`, `yyyy`, `h`, `hh`, `m`/`mm`
(minutes, after `h` or before `s`), `s`, `ss`, `AM/PM` and `A/P` — for example
`"0.00"`, `"#,##0"`, `"0.00%"`, `"dd/mm/yyyy"`, `"dd-mmm-yyyy"`, `"dddd"`.
Month and day names are English. Scientific notation, fractions, elapsed time
(`[h]`), fractional seconds and conditional sections (`[>100]`) are not
supported yet; colour tags like `[Red]` are ignored.

## Date & time

Dates are serial numbers (1 = 1900-01-01, including Excel's fictional
1900-02-29). Date arguments may also be text in `yyyy-mm-dd` or `dd/mm/yyyy`
form, optionally followed by a time.

| Function | Syntax | Notes |
|---|---|---|
| `DATE` | `DATE(year, month, day)` | Months and days overflow like Excel; years 0–1899 add 1900 |
| `DATEVALUE` | `DATEVALUE(date_text)` | Accepts `yyyy-mm-dd` and `dd/mm/yyyy` (day first, not the US order) |
| `TIME` | `TIME(hour, minute, second)` | Wraps past 24 hours |
| `TODAY` | `TODAY()` | Local date; see note below |
| `NOW` | `NOW()` | Local date and time; see note below |
| `YEAR` | `YEAR(date)` | |
| `MONTH` | `MONTH(date)` | |
| `DAY` | `DAY(date)` | |
| `HOUR` | `HOUR(time)` | |
| `MINUTE` | `MINUTE(time)` | |
| `SECOND` | `SECOND(time)` | |
| `WEEKDAY` | `WEEKDAY(date, [type])` | Types 1, 2, 3 and 11–17 |
| `WEEKNUM` | `WEEKNUM(date, [type])` | Types 1, 2, 11–17 and 21 (ISO 8601) |
| `EDATE` | `EDATE(start_date, months)` | Clamps to the end of shorter months |
| `EOMONTH` | `EOMONTH(start_date, months)` | |
| `DAYS` | `DAYS(end_date, start_date)` | |
| `DATEDIF` | `DATEDIF(start_date, end_date, unit)` | Units `Y`, `M`, `D`, `MD`, `YM`, `YD`; `MD` never goes negative (Excel can) |
| `NETWORKDAYS` | `NETWORKDAYS(start_date, end_date, [holidays])` | Saturday and Sunday are weekends |
| `WORKDAY` | `WORKDAY(start_date, days, [holidays])` | |

**Volatile functions.** `TODAY`, `NOW`, `RAND` and `RANDBETWEEN` are evaluated
when their cell is (re)calculated. They do not refresh on their own yet: the
value changes only when the cell or one of its inputs is edited.

## Lookup

| Function | Syntax | Notes |
|---|---|---|
| `VLOOKUP` | `VLOOKUP(value, table, col_index, [approximate])` | `approximate` defaults to TRUE, as in Excel |
| `HLOOKUP` | `HLOOKUP(value, table, row_index, [approximate])` | |
| `XLOOKUP` | `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | match_mode 0 exact, -1 next smaller, 1 next larger, 2 wildcard; search_mode 1 or -1 (2/-2 are accepted and searched linearly). Returns a single value: with a 2-D return_array, the first value of the matched row/column (no dynamic arrays yet) |
| `INDEX` | `INDEX(range, row, [col])` | |
| `MATCH` | `MATCH(value, range, [type])` | `type`: 1 (default), 0 exact, -1 |
| `CHOOSE` | `CHOOSE(index, value1, ...)` | Returns a single value, so a chosen multi-cell range gives `#VALUE!` |

## Information

| Function | Syntax |
|---|---|
| `ISBLANK` | `ISBLANK(value)` |
| `ISNUMBER` | `ISNUMBER(value)` |
| `ISTEXT` | `ISTEXT(value)` |
| `ISNONTEXT` | `ISNONTEXT(value)` |
| `ISLOGICAL` | `ISLOGICAL(value)` |
| `ISERROR` | `ISERROR(value)` |
| `ISERR` | `ISERR(value)` |
| `ISNA` | `ISNA(value)` |
| `ISEVEN` | `ISEVEN(number)` |
| `ISODD` | `ISODD(number)` |
| `NA` | `NA()` |

## Financial

Money paid out is negative, money received positive (Excel's convention). `type` 0 = payments at period end (default), 1 = start.

| Function | Syntax |
|---|---|
| `PMT` | `PMT(rate, nper, pv, [fv], [type])` |
| `PPMT` | `PPMT(rate, per, nper, pv, [fv], [type])` |
| `IPMT` | `IPMT(rate, per, nper, pv, [fv], [type])` |
| `FV` | `FV(rate, nper, pmt, [pv], [type])` |
| `PV` | `PV(rate, nper, pmt, [fv], [type])` |
| `NPER` | `NPER(rate, pmt, pv, [fv], [type])` |
| `RATE` | `RATE(nper, pmt, pv, [fv], [type], [guess])` |
| `NPV` | `NPV(rate, value1, ...)` |
| `XNPV` | `XNPV(rate, values, dates)` |
| `IRR` | `IRR(values, [guess])` |
| `XIRR` | `XIRR(values, dates, [guess])` |
| `SLN` | `SLN(cost, salvage, life)` |
| `DB` | `DB(cost, salvage, life, period, [month])` |
| `DDB` | `DDB(cost, salvage, life, period, [factor])` |
| `EFFECT` | `EFFECT(nominal_rate, npery)` |
| `NOMINAL` | `NOMINAL(effect_rate, npery)` |

## Reference

| Function | Syntax | Notes |
|---|---|---|
| `ROW` | `ROW([reference])` | Without an argument: this cell's row |
| `COLUMN` | `COLUMN([reference])` | Without an argument: this cell's column |
| `ROWS` | `ROWS(array)` | |
| `COLUMNS` | `COLUMNS(array)` | |
| `ADDRESS` | `ADDRESS(row, column, [abs_num], [a1], [sheet_text])` | |
| `OFFSET` | `OFFSET(reference, rows, cols, [height], [width])` | Volatile |
| `INDIRECT` | `INDIRECT(ref_text, [a1])` | Volatile; A1 style |
| `SUBTOTAL` | `SUBTOTAL(function_num, ref1, ...)` | Skips nested SUBTOTAL/AGGREGATE; 101–111 also skip hidden rows |
| `AGGREGATE` | `AGGREGATE(function_num, options, ref1, ...)` | Functions 1–13 with options 0–7 |

## Percentiles

| Function | Syntax |
|---|---|
| `PERCENTILE` | `PERCENTILE(array, k)` |
| `PERCENTILE.INC` | `PERCENTILE.INC(array, k)` |
| `PERCENTILE.EXC` | `PERCENTILE.EXC(array, k)` |
| `QUARTILE` | `QUARTILE(array, quart)` |
| `QUARTILE.INC` | `QUARTILE.INC(array, quart)` |
| `QUARTILE.EXC` | `QUARTILE.EXC(array, quart)` |

## Dynamic arrays

A formula whose result is several values **spills** into the cells below and to
the right, like Excel 365. If any of those cells holds something, the formula
shows `#SPILL!` until they are cleared. Operators work element by element
(`=A2:A10*B2:B10`, `=B2:B10>100`), and one-value functions such as `UPPER`,
`ROUND`, `LEFT`, `IF` and `TEXT` apply to each element of an array argument.
`SUM(A2:A10*B2:B10)` and `SUM(FILTER(...))` work without Ctrl+Shift+Enter.

| Function | Syntax | Notes |
|---|---|---|
| `FILTER` | `FILTER(array, include, [if_empty])` | `#CALC!` when nothing matches and no `if_empty` |
| `SORT` | `SORT(array, [sort_index], [sort_order], [by_col])` | Blanks and errors last |
| `SORTBY` | `SORTBY(array, by_array1, [sort_order1], ...)` | |
| `UNIQUE` | `UNIQUE(array, [by_col], [exactly_once])` | Text compared without case |
| `SEQUENCE` | `SEQUENCE(rows, [columns], [start], [step])` | |
| `RANDARRAY` | `RANDARRAY([rows], [columns], [min], [max], [whole_number])` | Volatile |
| `TRANSPOSE` | `TRANSPOSE(array)` | |
| `TAKE` | `TAKE(array, rows, [columns])` | Negative counts from the end |
| `DROP` | `DROP(array, rows, [columns])` | Negative counts from the end |
| `CHOOSEROWS` | `CHOOSEROWS(array, row_num1, ...)` | |
| `CHOOSECOLS` | `CHOOSECOLS(array, col_num1, ...)` | |
| `VSTACK` | `VSTACK(array1, ...)` | Shorter rows padded with `#N/A` |
| `HSTACK` | `HSTACK(array1, ...)` | |
| `TOCOL` | `TOCOL(array, [ignore], [scan_by_column])` | ignore: 1 blanks, 2 errors, 3 both |
| `TOROW` | `TOROW(array, [ignore], [scan_by_column])` | |
| `EXPAND` | `EXPAND(array, rows, [columns], [pad_with])` | |
| `TEXTSPLIT` | `TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty])` | Delimiters may be arrays of several strings |
| `XMATCH` | `XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])` | match_mode 0, -1, 1, 2 (wildcards) |
| `LET` | `LET(name1, value1, ..., calculation)` | Names may hold arrays |

## Not yet supported

`LAMBDA` and its helpers (`MAP`, `REDUCE`, `BYROW`…), array constants
(`{1,2;3,4}`), `TIMEVALUE`, database functions (`DSUM`…) and a few statistics
(`CORREL`, `FORECAST`…). Good first issues — see
[CONTRIBUTING.md](../../CONTRIBUTING.md).
