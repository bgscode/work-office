use crate::value::ErrorKind;

#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    Number(f64),
    Text(String),
    /// Cell refs, function names, TRUE/FALSE: anything starting with a letter, `_` or `$`.
    Ident(String),
    LParen,
    RParen,
    Comma,
    Colon,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Amp,
    Percent,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    /// An error literal such as `#REF!` (written into formulas when references are deleted).
    Error(ErrorKind),
}

pub fn tokenize(src: &str) -> Result<Vec<Token>, ErrorKind> {
    let chars: Vec<char> = src.chars().collect();
    let mut tokens = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        let simple = match c {
            '(' => Some(Token::LParen),
            ')' => Some(Token::RParen),
            ',' => Some(Token::Comma),
            ':' => Some(Token::Colon),
            '+' => Some(Token::Plus),
            '-' => Some(Token::Minus),
            '*' => Some(Token::Star),
            '/' => Some(Token::Slash),
            '^' => Some(Token::Caret),
            '&' => Some(Token::Amp),
            '%' => Some(Token::Percent),
            '=' => Some(Token::Eq),
            _ => None,
        };
        if let Some(t) = simple {
            tokens.push(t);
            i += 1;
            continue;
        }
        match c {
            c if c.is_whitespace() => i += 1,
            '#' => {
                let rest: String = chars[i..].iter().take(8).collect::<String>().to_ascii_uppercase();
                let kind = ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"]
                    .iter()
                    .find(|lit| rest.starts_with(*lit))
                    .ok_or(ErrorKind::Parse)?;
                tokens.push(Token::Error(ErrorKind::parse(kind).unwrap()));
                i += kind.len();
            }
            '<' | '>' => {
                let next = chars.get(i + 1).copied();
                let (t, len) = match (c, next) {
                    ('<', Some('=')) => (Token::Le, 2),
                    ('<', Some('>')) => (Token::Ne, 2),
                    ('>', Some('=')) => (Token::Ge, 2),
                    ('<', _) => (Token::Lt, 1),
                    _ => (Token::Gt, 1),
                };
                tokens.push(t);
                i += len;
            }
            '"' => {
                // "" inside a string is an escaped quote.
                let mut s = String::new();
                i += 1;
                loop {
                    match chars.get(i) {
                        None => return Err(ErrorKind::Parse),
                        Some('"') if chars.get(i + 1) == Some(&'"') => {
                            s.push('"');
                            i += 2;
                        }
                        Some('"') => {
                            i += 1;
                            break;
                        }
                        Some(&ch) => {
                            s.push(ch);
                            i += 1;
                        }
                    }
                }
                tokens.push(Token::Text(s));
            }
            c if c.is_ascii_digit() || c == '.' => {
                let start = i;
                while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                    i += 1;
                }
                // Exponent: 1e5, 2.5E-3
                if i < chars.len() && (chars[i] == 'e' || chars[i] == 'E') {
                    let mut j = i + 1;
                    if j < chars.len() && (chars[j] == '+' || chars[j] == '-') {
                        j += 1;
                    }
                    if j < chars.len() && chars[j].is_ascii_digit() {
                        i = j;
                        while i < chars.len() && chars[i].is_ascii_digit() {
                            i += 1;
                        }
                    }
                }
                let text: String = chars[start..i].iter().collect();
                tokens.push(Token::Number(text.parse().map_err(|_| ErrorKind::Parse)?));
            }
            c if c.is_alphabetic() || c == '_' || c == '$' => {
                let start = i;
                while i < chars.len() && (chars[i].is_alphanumeric() || matches!(chars[i], '_' | '$' | '.')) {
                    i += 1;
                }
                tokens.push(Token::Ident(chars[start..i].iter().collect()));
            }
            _ => return Err(ErrorKind::Parse),
        }
    }
    Ok(tokens)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tokenizes() {
        assert_eq!(
            tokenize(r#"SUM(A1:B2) >= 1.5e2 & "a""b""#).unwrap(),
            vec![
                Token::Ident("SUM".into()),
                Token::LParen,
                Token::Ident("A1".into()),
                Token::Colon,
                Token::Ident("B2".into()),
                Token::RParen,
                Token::Ge,
                Token::Number(150.0),
                Token::Amp,
                Token::Text("a\"b".into()),
            ]
        );
        assert!(tokenize("\"open").is_err());
    }
}
