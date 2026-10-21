//! Recursive-descent parser. Precedence follows Excel, lowest first:
//!
//! 1. comparison `= <> < <= > >=`
//! 2. concatenation `&`
//! 3. `+ -`
//! 4. `* /`
//! 5. `^` (left-associative, as in Excel: 2^3^2 = 64)
//! 6. postfix `%`
//! 7. prefix `-` / `+` (binds tighter than `^`, as in Excel: -2^2 = 4)
//! 8. primary: number, "text", TRUE/FALSE, A1, A1:B2, FUNC(args), (expr)

use super::ast::{BinOp, Expr, UnaryOp};
use super::lexer::{tokenize, Token};
use crate::cell::{CellRef, Range};
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
            Token::Number(n) => Ok(Expr::Number(n)),
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
    fn rejects_bad_syntax() {
        for bad in ["", "1+", "(1", "SUM(1,", "1 2", "A1:"] {
            assert!(parse(bad).is_err(), "{bad} should fail");
        }
    }
}
