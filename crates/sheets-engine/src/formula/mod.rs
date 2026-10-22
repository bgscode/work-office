//! Formula pipeline: text -> tokens (`lexer`) -> AST (`parser`) -> value (`eval`).
//! `special` holds the functions that need references or the current cell
//! (`ROW`, `OFFSET`, `INDIRECT`, `SUBTOTAL`...); `refs` rewrites formula text.

pub mod ast;
pub mod eval;
pub mod lexer;
pub mod parser;
pub mod refs;
mod special;

pub use ast::{BinOp, Expr, UnaryOp, VOLATILE_FUNCTIONS};
pub use eval::{eval, eval_at, CellSource, Reference};
pub use parser::parse;
