use crate::cell::{CellRef, Range};
use crate::value::ErrorKind;

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    Ref(CellRef),
    Range(Range),
    /// An identifier that is not a cell, function or boolean (e.g. a future named range).
    Name(String),
    Unary(UnaryOp, Box<Expr>),
    Binary(BinOp, Box<Expr>, Box<Expr>),
    /// Function name is stored upper-case.
    Call(String, Vec<Expr>),
    /// The formula failed to parse; evaluates to this error.
    Error(ErrorKind),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOp {
    Neg,
    Percent,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

impl Expr {
    /// Every cell/range this expression reads. Used to build the dependency graph.
    pub fn precedents(&self) -> Vec<Range> {
        let mut out = Vec::new();
        self.collect(&mut out);
        out
    }

    fn collect(&self, out: &mut Vec<Range>) {
        match self {
            Expr::Ref(c) => out.push(Range::single(*c)),
            Expr::Range(r) => out.push(*r),
            Expr::Unary(_, e) => e.collect(out),
            Expr::Binary(_, a, b) => {
                a.collect(out);
                b.collect(out);
            }
            Expr::Call(_, args) => args.iter().for_each(|a| a.collect(out)),
            Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Name(_) | Expr::Error(_) => {}
        }
    }
}
