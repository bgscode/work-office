//! Excel number format codes: `#,##0.00`, `0.0%`, `dd-mmm-yyyy`, `[h]:mm`,
//! `[Red](#,##0)`, `0.00E+00`, `# ?/?`, `"₹"#,##,##0`, `@" kg"`, accounting
//! formats and conditions like `[>=1000]#,##0,"K";0`.
//!
//! A code has up to four sections (positive; negative; zero; text) separated
//! by `;`. Each section may start with a colour and/or a condition in
//! brackets. Parsed codes are cached, so formatting a column of 100,000 cells
//! parses its code once.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use crate::date;
use crate::value::format_number;

/// Display text and the colour a section asked for (`[Red]` etc.).
#[derive(Debug, Clone, PartialEq)]
pub struct Formatted {
    pub text: String,
    pub color: Option<&'static str>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Digit {
    Zero,
    Hash,
    Question,
}

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Lit(String),
    Digit(Digit),
    Point,
    /// `E+` / `E-`: whether a plus sign is shown for positive exponents.
    Exp(bool),
    Slash,
    Percent,
    Text,
    General,
    Year(u8),
    Month(u8),
    Day(u8),
    Hour(u8),
    Minute(u8),
    Second(u8),
    /// Fractions of a second after `ss`: `.0`, `.00`, `.000`.
    SubSecond(u8),
    AmPm(bool),
    /// `[h]`, `[mm]`, `[ss]`: elapsed time, not wrapped at 24h/60m/60s.
    Elapsed(char, u8),
    /// A comma after the last digit placeholder: divide by 1000.
    Scale,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Cond {
    Lt(f64),
    Le(f64),
    Gt(f64),
    Ge(f64),
    Eq(f64),
    Ne(f64),
}

impl Cond {
    fn test(self, n: f64) -> bool {
        match self {
            Cond::Lt(v) => n < v,
            Cond::Le(v) => n <= v,
            Cond::Gt(v) => n > v,
            Cond::Ge(v) => n >= v,
            Cond::Eq(v) => n == v,
            Cond::Ne(v) => n != v,
        }
    }
}

#[derive(Debug, Clone, Default)]
struct Section {
    toks: Vec<Tok>,
    color: Option<&'static str>,
    cond: Option<Cond>,
    thousands: bool,
    is_date: bool,
    has_text: bool,
    has_digits: bool,
}

#[derive(Debug, Clone)]
struct Parsed {
    sections: Vec<Section>,
}

const COLORS: [(&str, &str); 8] = [
    ("black", "#000000"),
    ("blue", "#0000FF"),
    ("cyan", "#00FFFF"),
    ("green", "#008000"),
    ("magenta", "#FF00FF"),
    ("red", "#FF0000"),
    ("white", "#FFFFFF"),
    ("yellow", "#FFFF00"),
];

/// Excel's 56-colour palette for `[Color1]`..`[Color56]` (first 16 shown; others fall back to black).
const PALETTE: [&str; 16] = [
    "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", "#800000", "#008000",
    "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
];

fn parse_bracket(inner: &str, sec: &mut Section) {
    let lower = inner.to_ascii_lowercase();
    if let Some((_, c)) = COLORS.iter().find(|(n, _)| *n == lower) {
        sec.color = Some(c);
        return;
    }
    if let Some(n) = lower.strip_prefix("color").and_then(|n| n.parse::<usize>().ok()) {
        sec.color = Some(PALETTE.get(n.saturating_sub(1)).copied().unwrap_or("#000000"));
        return;
    }
    // Conditions: [>=100], [<0], [=0], [<>5]
    type MakeCond = fn(f64) -> Cond;
    let ops: [(&str, MakeCond); 6] =
        [(">=", Cond::Ge), ("<=", Cond::Le), ("<>", Cond::Ne), (">", Cond::Gt), ("<", Cond::Lt), ("=", Cond::Eq)];
    for (op, make) in ops {
        if let Some(v) = inner.strip_prefix(op).and_then(|v| v.trim().parse::<f64>().ok()) {
            sec.cond = Some(make(v));
            return;
        }
    }
    // Elapsed time: [h], [hh], [m], [mm], [s], [ss]
    if !lower.is_empty() && lower.chars().all(|c| c == lower.chars().next().unwrap()) {
        let c = lower.chars().next().unwrap();
        if matches!(c, 'h' | 'm' | 's') {
            sec.toks.push(Tok::Elapsed(c, lower.len() as u8));
            sec.is_date = true;
            return;
        }
    }
    // Currency/locale: [$₹-4009], [$€-x-euro2], [$-409] (locale only).
    if let Some(rest) = inner.strip_prefix('$') {
        let symbol = rest.split('-').next().unwrap_or("");
        if !symbol.is_empty() {
            sec.toks.push(Tok::Lit(symbol.to_string()));
        }
    }
}

fn parse_section(src: &str) -> Section {
    let mut sec = Section::default();
    let chars: Vec<char> = src.chars().collect();
    let mut i = 0;
    let push_lit = |sec: &mut Section, s: &str| {
        if let Some(Tok::Lit(last)) = sec.toks.last_mut() {
            last.push_str(s);
        } else {
            sec.toks.push(Tok::Lit(s.to_string()));
        }
    };
    while i < chars.len() {
        let c = chars[i];
        let lower = c.to_ascii_lowercase();
        let run = |ch: char| {
            let mut n = 0;
            while i + n < chars.len() && chars[i + n].to_ascii_lowercase() == ch {
                n += 1;
            }
            n
        };
        match c {
            '"' => {
                let mut j = i + 1;
                let mut s = String::new();
                while j < chars.len() && chars[j] != '"' {
                    s.push(chars[j]);
                    j += 1;
                }
                push_lit(&mut sec, &s);
                i = j + 1;
                continue;
            }
            '\\' => {
                if let Some(n) = chars.get(i + 1) {
                    push_lit(&mut sec, &n.to_string());
                }
                i += 2;
                continue;
            }
            // `_x`: a space as wide as x.
            '_' => {
                push_lit(&mut sec, " ");
                i += 2;
                continue;
            }
            // `*x`: repeat x to fill the cell; we don't know the width, so it's left out.
            '*' => {
                i += 2;
                continue;
            }
            '[' => {
                let end = chars[i..].iter().position(|ch| *ch == ']').map_or(chars.len(), |p| i + p);
                let inner: String = chars[i + 1..end.min(chars.len())].iter().collect();
                parse_bracket(&inner, &mut sec);
                i = end + 1;
                continue;
            }
            '0' => sec.toks.push(Tok::Digit(Digit::Zero)),
            '#' => sec.toks.push(Tok::Digit(Digit::Hash)),
            '?' => sec.toks.push(Tok::Digit(Digit::Question)),
            '.' => {
                // `ss.000`: fractions of a second.
                if matches!(sec.toks.last(), Some(Tok::Second(_)) | Some(Tok::Elapsed('s', _)))
                    && chars.get(i + 1) == Some(&'0')
                {
                    let mut n = 0;
                    while chars.get(i + 1 + n) == Some(&'0') {
                        n += 1;
                    }
                    sec.toks.push(Tok::SubSecond(n as u8));
                    i += 1 + n;
                    continue;
                }
                sec.toks.push(Tok::Point)
            }
            ',' => {
                let prev_digit = matches!(sec.toks.last(), Some(Tok::Digit(_)) | Some(Tok::Scale));
                let next_digit = matches!(chars.get(i + 1), Some('0' | '#' | '?'));
                if prev_digit && next_digit {
                    sec.thousands = true;
                } else if prev_digit {
                    sec.toks.push(Tok::Scale);
                } else {
                    push_lit(&mut sec, ",");
                }
            }
            '%' => sec.toks.push(Tok::Percent),
            '@' => {
                sec.toks.push(Tok::Text);
                sec.has_text = true;
            }
            '/' if sec.toks.iter().any(|t| matches!(t, Tok::Digit(_)))
                && matches!(chars.get(i + 1), Some('0'..='9' | '#' | '?')) =>
            {
                sec.toks.push(Tok::Slash)
            }
            'E' | 'e'
                if matches!(chars.get(i + 1), Some('+' | '-'))
                    && sec.toks.iter().any(|t| matches!(t, Tok::Digit(_))) =>
            {
                sec.toks.push(Tok::Exp(chars[i + 1] == '+'));
                i += 2;
                continue;
            }
            _ if lower == 'y' => {
                let n = run('y');
                sec.toks.push(Tok::Year(if n <= 2 { 2 } else { 4 }));
                sec.is_date = true;
                i += n;
                continue;
            }
            _ if lower == 'd' => {
                let n = run('d');
                sec.toks.push(Tok::Day(n.min(4) as u8));
                sec.is_date = true;
                i += n;
                continue;
            }
            _ if lower == 'h' => {
                let n = run('h');
                sec.toks.push(Tok::Hour(n.min(2) as u8));
                sec.is_date = true;
                i += n;
                continue;
            }
            _ if lower == 's' => {
                let n = run('s');
                sec.toks.push(Tok::Second(n.min(2) as u8));
                sec.is_date = true;
                i += n;
                continue;
            }
            _ if lower == 'm' => {
                let n = run('m');
                sec.toks.push(Tok::Month(n.min(5) as u8));
                sec.is_date = true;
                i += n;
                continue;
            }
            _ if lower == 'a' => {
                let rest: String = chars[i..].iter().take(5).collect::<String>().to_ascii_lowercase();
                if rest.starts_with("am/pm") {
                    sec.toks.push(Tok::AmPm(true));
                    sec.is_date = true;
                    i += 5;
                    continue;
                }
                if rest.starts_with("a/p") {
                    sec.toks.push(Tok::AmPm(false));
                    sec.is_date = true;
                    i += 3;
                    continue;
                }
                push_lit(&mut sec, &c.to_string());
            }
            _ if lower == 'g' && chars[i..].iter().take(7).collect::<String>().eq_ignore_ascii_case("general") => {
                sec.toks.push(Tok::General);
                i += 7;
                continue;
            }
            _ => push_lit(&mut sec, &c.to_string()),
        }
        i += 1;
    }
    // `m`/`mm` right after hours or right before seconds means minutes.
    let n = sec.toks.len();
    for k in 0..n {
        if let Tok::Month(len) = sec.toks[k] {
            if len > 2 {
                continue;
            }
            let prev = sec.toks[..k].iter().rev().find(|t| !matches!(t, Tok::Lit(_)));
            let next = sec.toks[k + 1..].iter().find(|t| !matches!(t, Tok::Lit(_)));
            if matches!(prev, Some(Tok::Hour(_) | Tok::Elapsed('h', _)))
                || matches!(next, Some(Tok::Second(_) | Tok::Elapsed('s', _)))
            {
                sec.toks[k] = Tok::Minute(len);
            }
        }
    }
    sec.has_digits = sec.toks.iter().any(|t| matches!(t, Tok::Digit(_)));
    sec
}

/// Split at `;` outside quotes, brackets and escapes.
fn split_sections(code: &str) -> Vec<String> {
    let mut out = vec![String::new()];
    let (mut quote, mut bracket, mut escape) = (false, false, false);
    for c in code.chars() {
        if escape {
            out.last_mut().unwrap().push(c);
            escape = false;
            continue;
        }
        match c {
            '\\' if !quote => escape = true,
            '"' => quote = !quote,
            '[' if !quote => bracket = true,
            ']' if !quote => bracket = false,
            ';' if !quote && !bracket => {
                out.push(String::new());
                continue;
            }
            _ => {}
        }
        out.last_mut().unwrap().push(c);
        if c == '\\' && !quote {
            // keep the backslash for parse_section
        }
    }
    out
}

thread_local! {
    static CACHE: RefCell<HashMap<String, Rc<Parsed>>> = RefCell::new(HashMap::new());
}

fn parsed(code: &str) -> Rc<Parsed> {
    CACHE.with(|c| {
        if let Some(p) = c.borrow().get(code) {
            return p.clone();
        }
        let p = Rc::new(Parsed { sections: split_sections(code).iter().map(|s| parse_section(s)).collect() });
        let mut map = c.borrow_mut();
        if map.len() > 2000 {
            map.clear();
        }
        map.insert(code.to_string(), p.clone());
        p
    })
}

/// Format a number with an Excel format code.
pub fn format_number_code(n: f64, code: &str) -> Formatted {
    let p = parsed(code);
    let secs = &p.sections;
    if !n.is_finite() {
        return Formatted { text: "#NUM!".into(), color: None };
    }
    // Choose the section: conditions first, then the sign rules.
    let has_cond = secs.iter().any(|s| s.cond.is_some());
    let (sec, abs) = if has_cond {
        let numeric: Vec<&Section> = secs.iter().filter(|s| !s.has_text || s.has_digits).take(3).collect();
        match numeric.iter().position(|s| s.cond.is_some_and(|c| c.test(n))) {
            Some(i) => (numeric[i], n < 0.0 && numeric[i].cond.is_some_and(|c| matches!(c, Cond::Lt(_) | Cond::Le(_)))),
            None => match numeric.iter().find(|s| s.cond.is_none()) {
                Some(s) => (*s, false),
                None => return Formatted { text: "#".repeat(5), color: None },
            },
        }
    } else if n < 0.0 && secs.len() >= 2 {
        (&secs[1], true)
    } else if n == 0.0 && secs.len() >= 3 {
        (&secs[2], false)
    } else {
        (&secs[0], false)
    };
    let value = if abs { n.abs() } else { n };
    let text = format_section(sec, value);
    Formatted { text, color: sec.color }
}

/// Format text with a code: the fourth section, or a section with `@`.
pub fn format_text_code(s: &str, code: &str) -> Formatted {
    let p = parsed(code);
    let sec = if p.sections.len() >= 4 { Some(&p.sections[3]) } else { p.sections.iter().find(|s| s.has_text) };
    match sec {
        Some(sec) => {
            let mut out = String::new();
            for t in &sec.toks {
                match t {
                    Tok::Text => out.push_str(s),
                    Tok::Lit(l) => out.push_str(l),
                    _ => {}
                }
            }
            Formatted { text: out, color: sec.color }
        }
        None => Formatted { text: s.to_string(), color: None },
    }
}

fn format_section(sec: &Section, n: f64) -> String {
    if sec.toks.iter().any(|t| matches!(t, Tok::General)) && !sec.has_digits {
        let body = format_number(n);
        return sec
            .toks
            .iter()
            .map(|t| {
                if matches!(t, Tok::General) {
                    body.clone()
                } else if let Tok::Lit(l) = t {
                    l.clone()
                } else {
                    String::new()
                }
            })
            .collect();
    }
    if sec.is_date {
        return format_date(sec, n);
    }
    if !sec.has_digits {
        // Only literals (e.g. `"-"` for zero in accounting formats), with a sign if negative.
        let lits: String =
            sec.toks.iter().filter_map(|t| if let Tok::Lit(l) = t { Some(l.as_str()) } else { None }).collect();
        return if sec.toks.iter().any(|t| matches!(t, Tok::Percent)) { format!("{lits}%") } else { lits };
    }
    format_digits(sec, n)
}

// ---------------------------------------------------------------- numbers

fn format_digits(sec: &Section, n: f64) -> String {
    let mut n = n;
    for t in &sec.toks {
        match t {
            Tok::Percent => n *= 100.0,
            Tok::Scale => n /= 1000.0,
            _ => {}
        }
    }
    if let Some(slash) = sec.toks.iter().position(|t| matches!(t, Tok::Slash)) {
        return format_fraction(sec, n, slash);
    }
    let negative = n < 0.0;
    let n = n.abs();
    let point = sec.toks.iter().position(|t| matches!(t, Tok::Point));
    let exp = sec.toks.iter().position(|t| matches!(t, Tok::Exp(_)));
    let int_end = point.or(exp).unwrap_or(sec.toks.len());
    let frac_digits: Vec<Digit> = match point {
        Some(p) => sec.toks[p + 1..exp.unwrap_or(sec.toks.len())]
            .iter()
            .filter_map(|t| if let Tok::Digit(d) = t { Some(*d) } else { None })
            .collect(),
        None => Vec::new(),
    };
    let int_count = sec.toks[..int_end].iter().filter(|t| matches!(t, Tok::Digit(_))).count();

    // Scientific: mantissa with `int_count` integer digits.
    let (mantissa, exponent) = if let Some(e) = exp {
        let _ = e;
        if n == 0.0 {
            (0.0, 0i32)
        } else {
            let shift = int_count.max(1) as i32 - 1;
            let mut ex = n.log10().floor() as i32 - shift;
            let mut m = n / 10f64.powi(ex);
            let rounded = round_to(m, frac_digits.len());
            if rounded >= 10f64.powi(shift + 1) {
                ex += 1;
                m = n / 10f64.powi(ex);
            }
            (m, ex)
        }
    } else {
        (n, 0)
    };

    let rounded = round_to(mantissa, frac_digits.len());
    let s = format!("{:.*}", frac_digits.len(), rounded);
    let (int_str, frac_str) = s.split_once('.').map_or((s.as_str(), ""), |(a, b)| (a, b));
    let int_str = if int_str == "0" { "" } else { int_str };

    // Integer part: fill placeholders right to left; extra digits go in front.
    let int_toks = &sec.toks[..int_end];
    let mut digits: Vec<char> = int_str.chars().collect();
    let mut out_rev: Vec<String> = Vec::new();
    let placeholder_count = int_toks.iter().filter(|t| matches!(t, Tok::Digit(_))).count();
    let mut seen = 0;
    let mut pos_from_right = 0; // digit positions emitted, for thousands separators
    for t in int_toks.iter().rev() {
        match t {
            Tok::Digit(d) => {
                seen += 1;
                let last_placeholder = seen == placeholder_count;
                let mut piece = String::new();
                let mut take = |piece: &mut String, ch: char| {
                    if sec.thousands && pos_from_right > 0 && pos_from_right % 3 == 0 {
                        piece.insert(0, ',');
                    }
                    piece.insert(0, ch);
                    pos_from_right += 1;
                };
                if let Some(ch) = digits.pop() {
                    take(&mut piece, ch);
                } else {
                    match d {
                        Digit::Zero => take(&mut piece, '0'),
                        Digit::Question => piece.insert(0, ' '),
                        Digit::Hash => {}
                    }
                }
                if last_placeholder {
                    while let Some(ch) = digits.pop() {
                        take(&mut piece, ch);
                    }
                }
                out_rev.push(piece);
            }
            Tok::Lit(l) => out_rev.push(l.clone()),
            Tok::Percent => out_rev.push("%".into()),
            Tok::Text => {}
            _ => {}
        }
    }
    let mut out: String = out_rev.into_iter().rev().collect();
    if placeholder_count == 0 && !int_str.is_empty() {
        out.push_str(int_str);
    }

    // Fraction part.
    if let Some(p) = point {
        let mut frac: Vec<char> = frac_str.chars().collect();
        // `#` drops trailing zeros, `?` turns them into spaces.
        for (k, d) in frac_digits.iter().enumerate().rev() {
            if frac.get(k) == Some(&'0') && k + 1 >= frac.len() {
                match d {
                    Digit::Hash => {
                        frac.pop();
                    }
                    Digit::Question => frac[k] = ' ',
                    Digit::Zero => break,
                }
            } else {
                break;
            }
        }
        let frac: String = frac.into_iter().collect();
        out.push('.');
        let mut fi = frac.chars();
        for t in &sec.toks[p + 1..exp.unwrap_or(sec.toks.len())] {
            match t {
                Tok::Digit(_) => {
                    if let Some(ch) = fi.next() {
                        out.push(ch);
                    }
                }
                Tok::Lit(l) => out.push_str(l),
                Tok::Percent => out.push('%'),
                _ => {}
            }
        }
    }

    // Exponent.
    if let Some(e) = exp {
        let Tok::Exp(plus) = sec.toks[e] else { unreachable!() };
        let exp_digits = sec.toks[e + 1..].iter().filter(|t| matches!(t, Tok::Digit(_))).count().max(1);
        let sign = if exponent < 0 {
            "-"
        } else if plus {
            "+"
        } else {
            ""
        };
        out.push('E');
        out.push_str(sign);
        out.push_str(&format!("{:0width$}", exponent.abs(), width = exp_digits));
        for t in &sec.toks[e + 1..] {
            match t {
                Tok::Lit(l) => out.push_str(l),
                Tok::Percent => out.push('%'),
                _ => {}
            }
        }
    }
    let shows_digit = out.chars().any(|c| c.is_ascii_digit() && c != '0');
    if negative && shows_digit {
        // Put the minus before the first digit-ish character, after leading literals like "₹".
        format!("-{out}")
    } else {
        out
    }
}

fn round_to(x: f64, decimals: usize) -> f64 {
    let f = 10f64.powi(decimals.min(15) as i32);
    // Half away from zero, like Excel; the tiny nudge undoes binary noise (2.675 is 2.67499...).
    (x * f + 1e-9 * x.signum()).round() / f
}

fn format_fraction(sec: &Section, n: f64, slash: usize) -> String {
    let negative = n < 0.0;
    let n = n.abs();
    // Denominator: fixed digits (`/8`, `/100`) or the number of placeholders (`/??` = up to 99).
    let after = &sec.toks[slash + 1..];
    let fixed: String = after
        .iter()
        .take_while(|t| matches!(t, Tok::Lit(l) if l.chars().all(|c| c.is_ascii_digit())))
        .map(|t| if let Tok::Lit(l) = t { l.clone() } else { String::new() })
        .collect();
    let den_places = after.iter().take_while(|t| matches!(t, Tok::Digit(_))).count();
    let before = &sec.toks[..slash];
    // A space (literal) between digit groups means "whole number, then fraction".
    let split = before.iter().rposition(|t| matches!(t, Tok::Lit(l) if l.contains(' ')));
    let has_whole = split.is_some_and(|s| before[..s].iter().any(|t| matches!(t, Tok::Digit(_))));
    let whole = if has_whole { n.trunc() } else { 0.0 };
    let rest = n - whole;
    let (num, den) = if let Ok(d) = fixed.parse::<u64>() {
        ((rest * d as f64).round() as u64, d.max(1))
    } else {
        best_fraction(rest, 10u64.pow(den_places.clamp(1, 6) as u32) - 1)
    };
    let (whole, num) = if num == den && den > 0 { (whole + 1.0, 0) } else { (whole, num) };
    let mut out = String::new();
    if has_whole && whole > 0.0 {
        out.push_str(&format!("{}", whole as u64));
    }
    if num > 0 {
        if !out.is_empty() {
            out.push(' ');
        }
        out.push_str(&format!("{num}/{den}"));
    } else if out.is_empty() {
        out.push('0');
    }
    if negative && out != "0" {
        out.insert(0, '-');
    }
    out
}

/// Closest fraction with a denominator up to `max_den` (Stern-Brocot search).
fn best_fraction(x: f64, max_den: u64) -> (u64, u64) {
    let (mut best_n, mut best_d, mut best_err) = (0u64, 1u64, x.abs());
    for d in 1..=max_den.min(10_000) {
        let n = (x * d as f64).round();
        let err = (x - n / d as f64).abs();
        if err < best_err - 1e-12 {
            best_n = n as u64;
            best_d = d;
            best_err = err;
            if err < 1e-12 {
                break;
            }
        }
    }
    (best_n, best_d)
}

// ---------------------------------------------------------------- dates and times

const MONTHS: [&str; 12] = [
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
];
const DAYS: [&str; 7] = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

fn format_date(sec: &Section, serial: f64) -> String {
    if serial < 0.0 {
        return "#".repeat(10);
    }
    let sub_places = sec.toks.iter().find_map(|t| if let Tok::SubSecond(n) = t { Some(*n) } else { None }).unwrap_or(0);
    // Round to the smallest unit shown, as Excel does.
    let unit = 86_400.0 * 10f64.powi(sub_places as i32);
    let total = (serial * unit).round() / unit;
    let day = total.floor();
    let secs_f = (total - day) * 86_400.0;
    let secs = secs_f.floor() as i64;
    let sub = secs_f - secs as f64;
    let (y, m, d) = date::ymd_from_serial(day);
    let (h, mi, s) = (secs / 3600, (secs / 60) % 60, secs % 60);
    let ampm = sec.toks.iter().any(|t| matches!(t, Tok::AmPm(_)));
    let weekday = ((day as i64) + 6).rem_euclid(7) as usize; // serial 1 = Sunday 1900-01-01
    let mut out = String::new();
    for t in &sec.toks {
        match t {
            Tok::Lit(l) => out.push_str(l),
            Tok::Year(2) => out.push_str(&format!("{:02}", y.rem_euclid(100))),
            Tok::Year(_) => out.push_str(&format!("{y:04}")),
            Tok::Month(1) => out.push_str(&m.to_string()),
            Tok::Month(2) => out.push_str(&format!("{m:02}")),
            Tok::Month(3) => out.push_str(&MONTHS[(m as usize + 11) % 12][..3]),
            Tok::Month(4) => out.push_str(MONTHS[(m as usize + 11) % 12]),
            Tok::Month(_) => out.push_str(&MONTHS[(m as usize + 11) % 12][..1]),
            Tok::Day(1) => out.push_str(&d.to_string()),
            Tok::Day(2) => out.push_str(&format!("{d:02}")),
            Tok::Day(3) => out.push_str(&DAYS[weekday][..3]),
            Tok::Day(_) => out.push_str(DAYS[weekday]),
            Tok::Hour(len) => {
                let hh = if ampm { (h + 11) % 12 + 1 } else { h };
                out.push_str(&if *len >= 2 { format!("{hh:02}") } else { hh.to_string() });
            }
            Tok::Minute(len) => out.push_str(&if *len >= 2 { format!("{mi:02}") } else { mi.to_string() }),
            Tok::Second(len) => out.push_str(&if *len >= 2 { format!("{s:02}") } else { s.to_string() }),
            Tok::SubSecond(n) => {
                let frac = format!("{:.*}", *n as usize, sub);
                out.push_str(frac.trim_start_matches('0'));
            }
            Tok::AmPm(full) => {
                let pm = h >= 12;
                out.push_str(match (full, pm) {
                    (true, false) => "AM",
                    (true, true) => "PM",
                    (false, false) => "A",
                    (false, true) => "P",
                });
            }
            Tok::Elapsed(unit, len) => {
                let whole_secs = (total * 86_400.0).floor() as i64;
                let v = match unit {
                    'h' => whole_secs / 3600,
                    'm' => whole_secs / 60,
                    _ => whole_secs,
                };
                out.push_str(&format!("{v:0width$}", width = *len as usize));
            }
            Tok::Digit(Digit::Zero) => out.push('0'),
            _ => {}
        }
    }
    out
}

// ---------------------------------------------------------------- tests

#[cfg(test)]
mod tests {
    use super::*;

    fn f(n: f64, code: &str) -> String {
        format_number_code(n, code).text
    }

    #[test]
    fn numbers() {
        assert_eq!(f(1234.567, "#,##0.00"), "1,234.57");
        assert_eq!(f(1234.5, "0"), "1235");
        assert_eq!(f(0.5, "#.##"), ".5");
        assert_eq!(f(0.5, "0.00"), "0.50");
        assert_eq!(f(5.0, "000"), "005");
        assert_eq!(f(-1234.5, "#,##0.00"), "-1,234.50");
        assert_eq!(f(1234567.0, "#,##0,\"K\""), "1,235K");
        assert_eq!(f(1234567.0, "0.0,,\"M\""), "1.2M");
        assert_eq!(f(0.125, "0.0%"), "12.5%");
        assert_eq!(f(12345.678, "0.00E+00"), "1.23E+04");
        assert_eq!(f(0.000123, "0.0E+0"), "1.2E-4");
        assert_eq!(f(5551234.0, "000-0000"), "555-1234");
        assert_eq!(f(3.5, "0.0?"), "3.5 ");
    }

    #[test]
    fn sections_colours_and_conditions() {
        assert_eq!(f(-5.0, "#,##0;(#,##0)"), "(5)");
        assert_eq!(format_number_code(-5.0, "0;[Red]-0").color, Some("#FF0000"));
        assert_eq!(f(0.0, "0;-0;\"zero\""), "zero");
        assert_eq!(f(1500.0, "[>=1000]#,##0,\"K\";0"), "2K");
        assert_eq!(f(15.0, "[>=1000]#,##0,\"K\";0"), "15");
        // Accounting: the zero section shows a dash.
        assert_eq!(f(0.0, r#"_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)"#).trim(), "-");
        assert_eq!(f(-12.0, r#"_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)"#).trim(), "(12.00)");
    }

    #[test]
    fn currency_and_literals() {
        assert_eq!(f(1234567.0, "[$₹-4009] #,##0"), "₹ 1,234,567");
        assert_eq!(f(12.0, "0\" kg\""), "12 kg");
        assert_eq!(f(12.0, "\\$0"), "$12");
        assert_eq!(format_text_code("abc", "0;0;0;\"Name: \"@").text, "Name: abc");
        assert_eq!(format_text_code("abc", "@\"!\"").text, "abc!");
    }

    #[test]
    fn fractions() {
        assert_eq!(f(1.5, "# ?/?"), "1 1/2");
        assert_eq!(f(0.333, "?/?"), "1/3");
        assert_eq!(f(2.75, "# ?/8"), "2 6/8");
        assert_eq!(f(3.0, "# ?/?"), "3");
    }

    #[test]
    fn dates_and_times() {
        // 2025-08-15 14:30:45 = serial 45884.6047...
        let serial = date::serial_from_ymd(2025, 8, 15) + (14.0 * 3600.0 + 30.0 * 60.0 + 45.0) / 86400.0;
        assert_eq!(f(serial, "dd-mmm-yyyy"), "15-Aug-2025");
        assert_eq!(f(serial, "d/m/yy"), "15/8/25");
        assert_eq!(f(serial, "dddd, mmmm d, yyyy"), "Friday, August 15, 2025");
        assert_eq!(f(serial, "hh:mm:ss"), "14:30:45");
        assert_eq!(f(serial, "h:mm AM/PM"), "2:30 PM");
        assert_eq!(f(serial, "mmm-yy"), "Aug-25");
        assert_eq!(f(1.5, "[h]:mm"), "36:00");
        assert_eq!(f(0.5 + 0.25 / 86400.0, "mm:ss.00"), "00:00.25");
    }
}
