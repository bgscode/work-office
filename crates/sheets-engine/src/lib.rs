//! Spreadsheet engine for Office Sheets.
//!
//! Data flow for an edit:
//!
//! ```text
//! user input "=SUM(A1:A3)"
//!   -> formula::parse        (lexer -> parser -> Expr AST)
//!   -> Sheet::set_inputs     (stores input, updates dependency graph)
//!   -> Sheet::recalc         (re-evaluates only the dirty cells, in topological order)
//!   -> formula::eval         (Expr + cell values -> Value; functions come from `functions`)
//! ```
//!
//! See `crates/sheets-engine/README.md` for a guided tour.

pub mod borders;
pub mod cell;
pub mod date;
pub mod fill;
pub mod format;
pub mod formula;
pub mod functions;
pub mod layout;
pub mod ops;
pub mod painter;
pub mod sheet;
pub mod style;
pub mod value;
pub mod workbook;

pub use cell::{CellRef, Range, MAX_COLS, MAX_ROWS};
pub use sheet::{FindOptions, Sheet};
pub use style::{Align, NumberFormat, Style, StylePatch};
pub use style::{BorderLine, BorderStyle, Borders, Side, VAlign};
pub use value::{ErrorKind, Value};
pub use workbook::{Workbook, Worksheet};
