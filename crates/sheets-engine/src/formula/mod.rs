//! Formula pipeline: text -> tokens (`lexer`) -> AST (`parser`) -> value (`eval`).

pub mod ast;
pub mod eval;
pub mod lexer;
pub mod parser;
pub mod refs;

pub use ast::{BinOp, Expr, UnaryOp};
pub use eval::{eval, CellSource};
pub use parser::parse;
