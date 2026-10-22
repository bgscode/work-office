use crate::cell::{CellRef, Range};
use crate::value::ErrorKind;

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    Ref(CellRef),
    /// `A1:B2`; also whole columns `A:A` (rows 1..=1048576) and whole rows `1:1`.
    Range(Range),
    /// A cell on another sheet: `Sheet2!A1`. The sheet name is stored as written, unquoted.
    SheetRef(String, CellRef),
    /// A range on another sheet: `'My Sheet'!A1:B5`, `Sheet2!A:A`.
    SheetRange(String, Range),
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

/// Functions whose result can change without any of their inputs changing.
/// Like Excel, cells using them are recalculated on every recalculation.
pub const VOLATILE_FUNCTIONS: &[&str] = &["INDIRECT", "NOW", "OFFSET", "RAND", "RANDBETWEEN", "TODAY"];

impl Expr {
    /// Every cell/range on the formula's own sheet this expression reads.
    /// Used to build the dependency graph.
    pub fn precedents(&self) -> Vec<Range> {
        let mut out = Vec::new();
        self.visit(&mut |e| match e {
            Expr::Ref(c) => out.push(Range::single(*c)),
            Expr::Range(r) => out.push(*r),
            _ => {}
        });
        out
    }

    /// Every cell/range on a named sheet (`Sheet2!A1`) this expression reads.
    pub fn sheet_precedents(&self) -> Vec<(String, Range)> {
        let mut out = Vec::new();
        self.visit(&mut |e| match e {
            Expr::SheetRef(s, c) => out.push((s.clone(), Range::single(*c))),
            Expr::SheetRange(s, r) => out.push((s.clone(), *r)),
            _ => {}
        });
        out
    }

    /// True if the expression calls a volatile function (`NOW`, `RAND`, `OFFSET`...).
    pub fn is_volatile(&self) -> bool {
        let mut found = false;
        self.visit(&mut |e| {
            if let Expr::Call(name, _) = e {
                found |= VOLATILE_FUNCTIONS.contains(&name.as_str());
            }
        });
        found
    }

    /// True if the expression calls `name` (upper-case) anywhere.
    pub fn calls(&self, name: &str) -> bool {
        let mut found = false;
        self.visit(&mut |e| {
            if let Expr::Call(n, _) = e {
                found |= n == name;
            }
        });
        found
    }

    /// Call `f` on this node and every node below it.
    pub fn visit(&self, f: &mut dyn FnMut(&Expr)) {
        f(self);
        match self {
            Expr::Unary(_, e) => e.visit(f),
            Expr::Binary(_, a, b) => {
                a.visit(f);
                b.visit(f);
            }
            Expr::Call(_, args) => args.iter().for_each(|a| a.visit(f)),
            _ => {}
        }
    }
}
