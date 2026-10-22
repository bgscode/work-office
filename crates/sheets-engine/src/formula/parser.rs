//! Recursive-descent parser. Precedence follows Excel, lowest first:
//!
//! 1. comparison `= <> < <= > >=`
//! 2. concatenation `&`
//! 3. `+ -`
//! 4. `* /`
//! 5. `^` (left-associative, as in Excel: 2^3^2 = 64)
//! 6. postfix `%`
//! 7. prefix `-` / `+` (binds tighter than `^`, as in Excel: -2^2 = 4)
//! 8. primary: number, "text", TRUE/FALSE, A1, A1:B2, A:A, 1:1, Sheet2!A1,
//!    'My Sheet'!A1:B5, FUNC(args), (expr)

use super::ast::{BinOp, Expr, UnaryOp};
use super::lexer::{tokenize, Token};
use crate::cell::{parse_col, CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::value::ErrorKind;

/// Parse formula text without the leading `=`.
pub fn parse(src: &str) -> Result<Expr, ErrorKind> {
    let tokens = tokenize(src)?;
    let mut p = Parser { tokens, pos: 0 };
    let expr = p.comparison()?;
    if p.pos != p.tokens.len() {
        return Err(ErrorKind::Parse);
    }
    Ok(expr)
}

struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.pos)
    }

    fn next(&mut self) -> Option<Token> {
        let t = self.tokens.get(self.pos).cloned();
        self.pos += 1;
        t
    }

    fn eat(&mut self, t: &Token) -> bool {
        if self.peek() == Some(t) {
            self.pos += 1;
            true
        } else {
            false
        }
    }

    fn comparison(&mut self) -> Result<Expr, ErrorKind> {
        let mut left = self.concat()?;
        loop {
            let op = match self.peek() {
                Some(Token::Eq) => BinOp::Eq,
                Some(Token::Ne) => BinOp::Ne,
                Some(Token::Lt) => BinOp::Lt,
                Some(Token::Le) => BinOp::Le,
                Some(Token::Gt) => BinOp::Gt,
                Some(Token::Ge) => BinOp::Ge,
                _ => return Ok(left),
            };
            self.pos += 1;
            left = Expr::Binary(op, Box::new(left), Box::new(self.concat()?));
        }
    }

    fn concat(&mut self) -> Result<Expr, ErrorKind> {
        let mut left = self.additive()?;
        while self.eat(&Token::Amp) {
            left = Expr::Binary(BinOp::Concat, Box::new(left), Box::new(self.additive()?));
        }
        Ok(left)
    }

    fn additive(&mut self) -> Result<Expr, ErrorKind> {
        let mut left = self.multiplicative()?;
        loop {
            let op = match self.peek() {
                Some(Token::Plus) => BinOp::Add,
                Some(Token::Minus) => BinOp::Sub,
                _ => return Ok(left),
            };
            self.pos += 1;
            left = Expr::Binary(op, Box::new(left), Box::new(self.multiplicative()?));
        }
    }

    fn multiplicative(&mut self) -> Result<Expr, ErrorKind> {
        let mut left = self.power()?;
        loop {
            let op = match self.peek() {
                Some(Token::Star) => BinOp::Mul,
                Some(Token::Slash) => BinOp::Div,
                _ => return Ok(left),
            };
            self.pos += 1;
            left = Expr::Binary(op, Box::new(left), Box::new(self.power()?));
        }
    }

    fn power(&mut self) -> Result<Expr, ErrorKind> {
        let mut left = self.percent()?;
        while self.eat(&Token::Caret) {
            left = Expr::Binary(BinOp::Pow, Box::new(left), Box::new(self.percent()?));
        }
        Ok(left)
    }

    fn percent(&mut self) -> Result<Expr, ErrorKind> {
        let mut e = self.unary()?;
        while self.eat(&Token::Percent) {
            e = Expr::Unary(UnaryOp::Percent, Box::new(e));
        }
        Ok(e)
    }

    fn unary(&mut self) -> Result<Expr, ErrorKind> {
        if self.eat(&Token::Minus) {
            return Ok(Expr::Unary(UnaryOp::Neg, Box::new(self.unary()?)));
        }
        if self.eat(&Token::Plus) {
            return self.unary();
        }
        self.primary()
    }

    fn primary(&mut self) -> Result<Expr, ErrorKind> {
        match self.next().ok_or(ErrorKind::Parse)? {
            Token::Number(n) => match self.row_range(&format_row(n)) {
                Some(r) => Ok(Expr::Range(r?)),
                None => Ok(Expr::Number(n)),
            },
            Token::Sheet(sheet) => match self.next() {
                Some(Token::Ident(name)) => match self.ident(name)? {
                    Expr::Ref(c) => Ok(Expr::SheetRef(sheet, c)),
                    Expr::Range(r) => Ok(Expr::SheetRange(sheet, r)),
                    _ => Err(ErrorKind::Parse),
                },
                Some(Token::Number(n)) => match self.row_range(&format_row(n)) {
                    Some(r) => Ok(Expr::SheetRange(sheet, r?)),
                    None => Err(ErrorKind::Parse),
                },
                // Excel writes `Sheet2!#REF!` for a deleted reference.
                Some(Token::Error(e)) => Ok(Expr::Error(e)),
                _ => Err(ErrorKind::Parse),
            },
            Token::Text(s) => Ok(Expr::Text(s)),
            Token::Error(e) => Ok(Expr::Error(e)),
            Token::LParen => {
                let e = self.comparison()?;
                if !self.eat(&Token::RParen) {
                    return Err(ErrorKind::Parse);
                }
                Ok(e)
            }
            Token::Ident(name) => self.ident(name),
            _ => Err(ErrorKind::Parse),
        }
    }

    fn ident(&mut self, name: String) -> Result<Expr, ErrorKind> {
        if self.eat(&Token::LParen) {
            let mut args = Vec::new();
            if !self.eat(&Token::RParen) {
                loop {
                    args.push(self.comparison()?);
                    if self.eat(&Token::RParen) {
                        break;
                    }
                    if !self.eat(&Token::Comma) {
                        return Err(ErrorKind::Parse);
                    }
                }
            }
            return Ok(Expr::Call(name.to_ascii_uppercase(), args));
        }
        if name.eq_ignore_ascii_case("TRUE") {
            return Ok(Expr::Bool(true));
        }
        if name.eq_ignore_ascii_case("FALSE") {
            return Ok(Expr::Bool(false));
        }
        if let Some(r) = self.col_range(&name).or_else(|| self.row_range(&name)) {
            return Ok(Expr::Range(r?));
        }
        let Some(start) = CellRef::parse(&name) else {
            return Ok(Expr::Name(name));
        };
        if self.eat(&Token::Colon) {
            let end = match self.next() {
                Some(Token::Ident(s)) => CellRef::parse(&s).ok_or(ErrorKind::Parse)?,
                _ => return Err(ErrorKind::Parse),
            };
            return Ok(Expr::Range(Range::new(start, end)));
        }
        Ok(Expr::Ref(start))
    }

    /// `A:C` (whole columns) when `first` is a column and a `:column` follows.
    fn col_range(&mut self, first: &str) -> Option<Result<Range, ErrorKind>> {
        let a = col_only(first)?;
        if self.peek() != Some(&Token::Colon) {
            return None;
        }
        self.pos += 1;
        let b = match self.next() {
            Some(Token::Ident(s)) => col_only(&s),
            _ => None,
        };
        Some(b.map(|b| Range::new(CellRef::new(0, a), CellRef::new(MAX_ROWS - 1, b))).ok_or(ErrorKind::Parse))
    }

    /// `1:3` (whole rows) when `first` is a row number and a `:row` follows.
    fn row_range(&mut self, first: &str) -> Option<Result<Range, ErrorKind>> {
        let a = row_only(first)?;
        if self.peek() != Some(&Token::Colon) {
            return None;
        }
        self.pos += 1;
        let b = match self.next() {
            Some(Token::Ident(s)) => row_only(&s),
            Some(Token::Number(n)) => row_only(&format_row(n)),
            _ => None,
        };
        Some(b.map(|b| Range::new(CellRef::new(a, 0), CellRef::new(b, MAX_COLS - 1))).ok_or(ErrorKind::Parse))
    }
}

/// A whole number token as row text ("" for anything else, which `row_only` rejects).
fn format_row(n: f64) -> String {
    if n.fract() == 0.0 && n >= 1.0 && n <= MAX_ROWS as f64 {
        format!("{}", n as u32)
    } else {
        String::new()
    }
}

/// `A`, `$XFD` -> zero-based column.
fn col_only(s: &str) -> Option<u32> {
    let letters = s.strip_prefix('$').unwrap_or(s);
    if letters.is_empty() || !letters.bytes().all(|b| b.is_ascii_alphabetic()) {
        return None;
    }
    parse_col(letters)
}

/// `1`, `$20` -> zero-based row.
fn row_only(s: &str) -> Option<u32> {
    let digits = s.strip_prefix('$').unwrap_or(s);
    if digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let n: u32 = digits.parse().ok()?;
    (1..=MAX_ROWS).contains(&n).then(|| n - 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn num(n: f64) -> Box<Expr> {
        Box::new(Expr::Number(n))
    }

    #[test]
    fn respects_precedence() {
        // 1 + 2 * 3
        assert_eq!(
            parse("1+2*3").unwrap(),
            Expr::Binary(BinOp::Add, num(1.0), Box::new(Expr::Binary(BinOp::Mul, num(2.0), num(3.0))))
        );
        // -2^2 is (-2)^2 in Excel
        assert_eq!(
            parse("-2^2").unwrap(),
            Expr::Binary(BinOp::Pow, Box::new(Expr::Unary(UnaryOp::Neg, num(2.0))), num(2.0))
        );
    }

    #[test]
    fn parses_calls_and_ranges() {
        let e = parse("sum(A1:B2, 3)").unwrap();
        let Expr::Call(name, args) = e else { panic!() };
        assert_eq!(name, "SUM");
        assert_eq!(args.len(), 2);
        assert!(matches!(args[0], Expr::Range(_)));
        assert_eq!(parse("NOW()").unwrap(), Expr::Call("NOW".into(), vec![]));
        assert_eq!(
            parse("#REF!+1").unwrap(),
            Expr::Binary(BinOp::Add, Box::new(Expr::Error(ErrorKind::Ref)), num(1.0))
        );
    }

    #[test]
    fn parses_whole_columns_rows_and_sheet_refs() {
        let col = |a, b| Range::new(CellRef::new(0, a), CellRef::new(MAX_ROWS - 1, b));
        let row = |a, b| Range::new(CellRef::new(a, 0), CellRef::new(b, MAX_COLS - 1));
        assert_eq!(parse("A:A").unwrap(), Expr::Range(col(0, 0)));
        assert_eq!(parse("$B:$D").unwrap(), Expr::Range(col(1, 3)));
        assert_eq!(parse("1:1").unwrap(), Expr::Range(row(0, 0)));
        assert_eq!(parse("$3:5").unwrap(), Expr::Range(row(2, 4)));
        assert_eq!(parse("Sheet2!B3").unwrap(), Expr::SheetRef("Sheet2".into(), CellRef::new(2, 1)));
        assert_eq!(
            parse("'My Sheet'!A1:B5").unwrap(),
            Expr::SheetRange("My Sheet".into(), Range::new(CellRef::new(0, 0), CellRef::new(4, 1)))
        );
        assert_eq!(parse("Data!C:C").unwrap(), Expr::SheetRange("Data".into(), col(2, 2)));
        assert_eq!(parse("Data!2:2").unwrap(), Expr::SheetRange("Data".into(), row(1, 1)));
        assert_eq!(parse("Data!#REF!").unwrap(), Expr::Error(ErrorKind::Ref));
        assert_eq!(parse("1+2").unwrap(), Expr::Binary(BinOp::Add, num(1.0), num(2.0)));
        for bad in ["A:", "A:1", "1:A", "Sheet2!SUM(1)", "Sheet2!", "0:1"] {
            assert!(parse(bad).is_err(), "{bad} should fail");
        }
    }

    #[test]
    fn rejects_bad_syntax() {
        for bad in ["", "1+", "(1", "SUM(1,", "1 2", "A1:"] {
            assert!(parse(bad).is_err(), "{bad} should fail");
        }
    }
}
