//! One grid of cells plus its dependency graph.
//!
//! Every formula cell records the ranges it reads (`precedents`). The reverse
//! index (`dependents`) maps each cell to the formulas that read it. When cells
//! change, we walk `dependents` to find everything that is now stale and
//! re-evaluate just those cells in topological order (Kahn's algorithm). Cells
//! left over after the sort are part of a cycle and get `#CIRC!`.
//!
//! The sheet also stores per-cell `Style`s and column widths. Structural edits
//! (insert/delete rows and columns, sort) build a new `Sheet` with formulas
//! rewritten; `ops::ReplaceSheet` makes them undoable.

use std::borrow::Cow;
use std::cmp::Ordering;
use std::collections::BTreeMap;

// FxHash: much faster than the default SipHash for small keys like CellRef;
// hash-flooding resistance doesn't matter for a document's own cells.
use rustc_hash::{FxHashMap as HashMap, FxHashSet as HashSet};

use crate::cell::{CellRef, Range, MAX_COLS, MAX_ROWS};
use crate::fill::{self, FillDirection};
use crate::format;
use crate::formula::refs::{self, Axis};
use crate::formula::{self, CellSource, Expr};
use crate::layout::SheetLayout;
use crate::style::{NumberFormat, Style};
use crate::value::{compare, Value};

/// Inputs and styles to write for a fill-handle drag.
pub type FillEdits = (Vec<(CellRef, String)>, Vec<(CellRef, Style)>);

/// Options for find and replace.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct FindOptions {
    pub match_case: bool,
    /// The whole cell must equal the query, not just contain it.
    pub whole_cell: bool,
}

impl FindOptions {
    fn matches(&self, text: &str, query: &str) -> bool {
        match (self.whole_cell, self.match_case) {
            (true, true) => text == query,
            (true, false) => text.to_lowercase() == query.to_lowercase(),
            (false, true) => text.contains(query),
            (false, false) => text.to_lowercase().contains(&query.to_lowercase()),
        }
    }
}

/// Replace every occurrence of `query` in `text`, optionally ignoring case.
/// Works on characters, so it is safe for any script (Hindi, accents...).
fn replace_text(text: &str, query: &str, replacement: &str, match_case: bool) -> String {
    if match_case {
        return text.replace(query, replacement);
    }
    let fold = |c: char| c.to_lowercase().next().unwrap_or(c);
    let hay: Vec<char> = text.chars().collect();
    let needle: Vec<char> = query.chars().map(fold).collect();
    let mut out = String::with_capacity(text.len());
    let mut i = 0;
    while i < hay.len() {
        if i + needle.len() <= hay.len() && hay[i..i + needle.len()].iter().map(|c| fold(*c)).eq(needle.iter().copied())
        {
            out.push_str(replacement);
            i += needle.len();
        } else {
            out.push(hay[i]);
            i += 1;
        }
    }
    out
}

/// Column width in CSS pixels when none is set: Excel's default of 8.43
/// characters of Calibri 11, which is 64px.
pub const DEFAULT_COL_WIDTH: f64 = 64.0;

#[derive(Debug, Clone)]
struct Cell {
    /// Exactly what the user typed, e.g. "=SUM(A1:A3)" or "12%". `None` when
    /// it is simply the value written back (`Value::to_input`), which is true
    /// for most literal cells: storing it would double the memory of big sheets.
    input: Option<Box<str>>,
    content: Content,
}

impl Cell {
    fn input(&self) -> Cow<'_, str> {
        match (&self.input, &self.content) {
            (Some(s), _) => Cow::Borrowed(s),
            (None, Content::Literal(v)) => Cow::Owned(v.to_input()),
            (None, Content::Formula { .. }) => Cow::Borrowed(""),
        }
    }

    fn has_line_break(&self) -> bool {
        match (&self.input, &self.content) {
            (Some(s), _) => s.contains('\n'),
            (None, Content::Literal(Value::Text(t))) => t.contains('\n'),
            _ => false,
        }
    }
}

#[derive(Debug, Clone)]
enum Content {
    Literal(Value),
    Formula {
        expr: Expr,
        /// Ranges on this sheet the formula reads.
        precedents: Vec<Range>,
        /// Ranges on other sheets (`Sheet2!A1`): (sheet name in lower case, range).
        external: Vec<(String, Range)>,
        value: Value,
    },
}

/// The number format Excel gives a formula's result by default: dates for
/// date functions (also `=TODAY()+30`), time for `TIME`, currency for loan
/// functions.
fn implied_format(expr: &Expr) -> Option<NumberFormat> {
    use crate::formula::BinOp;
    match expr {
        Expr::Call(name, _) => match name.as_str() {
            "TODAY" | "DATE" | "EDATE" | "EOMONTH" | "WORKDAY" | "DATEVALUE" => Some(NumberFormat::Date),
            "NOW" => Some(NumberFormat::DateTime),
            "TIME" => Some(NumberFormat::Time),
            "PMT" | "PV" | "FV" | "IPMT" | "PPMT" => Some(NumberFormat::Currency { symbol: "₹".into(), decimals: 2 }),
            _ => None,
        },
        Expr::Binary(BinOp::Add | BinOp::Sub, a, b) => {
            let date = |e: &Expr| matches!(implied_format(e), Some(NumberFormat::Date | NumberFormat::DateTime));
            (date(a) || date(b)).then_some(NumberFormat::Date).filter(|_| !(date(a) && date(b)))
        }
        _ => None,
    }
}

/// Ranges bigger than this (whole columns like `A:A`, large blocks) are kept
/// as ranges in the dependency graph instead of one entry per cell.
const BIG_RANGE: u64 = 4096;

#[derive(Debug, Default, Clone)]
pub struct Sheet {
    cells: HashMap<CellRef, Cell>,
    dependents: HashMap<CellRef, HashSet<CellRef>>,
    /// Formulas that read a big range, with that range (see `BIG_RANGE`).
    range_dependents: Vec<(Range, CellRef)>,
    /// Formulas that read other sheets; the workbook recalculates them.
    links: HashSet<CellRef>,
    /// Formulas with volatile functions.
    volatiles: HashSet<CellRef>,
    /// Smallest (rows, cols) that has ever held data. Grows only.
    extent: (u32, u32),
    styles: HashMap<CellRef, Style>,
    /// Column index -> width in pixels, for columns that aren't the default width.
    col_widths: BTreeMap<u32, f64>,
    /// Row heights, hidden rows/columns, merged cells, frozen panes (see `layout.rs`).
    layout: SheetLayout,
    /// AutoFilter, data validation and conditional formatting (see `data_tools.rs`).
    pub(crate) tools: crate::data_tools::DataTools,
}

static EMPTY: Value = Value::Empty;
static DEFAULT_STYLE: Style = Style {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color: None,
    fill: None,
    align: None,
    number_format: NumberFormat::General,
    font: None,
    font_size: None,
    wrap: false,
    valign: crate::style::VAlign::Bottom,
    borders: crate::style::Borders::NONE,
};

impl Sheet {
    pub fn new() -> Self {
        Self::default()
    }

    /// Set one cell and recalculate everything that depends on it.
    pub fn set_input(&mut self, at: CellRef, input: &str) {
        self.set_inputs([(at, input.to_string())]);
    }

    /// Set many cells, then recalculate once. Use this for paste and file import.
    /// An empty string clears the cell.
    pub fn set_inputs(&mut self, edits: impl IntoIterator<Item = (CellRef, String)>) {
        let mut changed = Vec::new();
        for (at, input) in edits {
            self.write(at, input);
            changed.push(at);
        }
        self.recalc_from(&changed);
    }

    /// The raw input of a cell ("" if empty).
    pub fn input(&self, at: CellRef) -> Cow<'_, str> {
        self.cells.get(&at).map_or(Cow::Borrowed(""), Cell::input)
    }

    /// The formula of a cell including its `=`, or `None` for literal and empty cells.
    pub fn formula_input(&self, at: CellRef) -> Option<&str> {
        match self.cells.get(&at) {
            Some(Cell { input: Some(s), content: Content::Formula { .. } }) => Some(s),
            _ => None,
        }
    }

    /// Every non-empty cell's position, sorted row by row.
    pub fn cell_positions(&self) -> Vec<CellRef> {
        let mut out: Vec<CellRef> = self.cells.keys().copied().collect();
        out.sort_unstable();
        out
    }

    /// The current value of a cell.
    pub fn value(&self, at: CellRef) -> &Value {
        match self.cells.get(&at).map(|c| &c.content) {
            None => &EMPTY,
            Some(Content::Literal(v)) => v,
            Some(Content::Formula { value, .. }) => value,
        }
    }

    pub fn is_formula(&self, at: CellRef) -> bool {
        matches!(self.cells.get(&at), Some(Cell { content: Content::Formula { .. }, .. }))
    }

    /// (rows, cols) covering every cell that has held data.
    pub fn extent(&self) -> (u32, u32) {
        self.extent
    }

    /// Value formatted with the cell's number format: what the grid shows.
    pub fn display(&self, at: CellRef) -> String {
        format::format_value(self.value(at), &self.style(at).number_format)
    }

    pub fn style(&self, at: CellRef) -> &Style {
        self.styles.get(&at).unwrap_or(&DEFAULT_STYLE)
    }

    pub fn set_style(&mut self, at: CellRef, style: Style) {
        if style.is_default() {
            self.styles.remove(&at);
        } else {
            self.styles.insert(at, style);
        }
    }

    /// Every cell with non-default formatting, sorted row by row.
    pub fn styled_cells(&self) -> Vec<(CellRef, &Style)> {
        let mut out: Vec<_> = self.styles.iter().map(|(at, s)| (*at, s)).collect();
        out.sort_by_key(|(at, _)| *at);
        out
    }

    pub fn col_width(&self, col: u32) -> f64 {
        self.col_widths.get(&col).copied().unwrap_or(DEFAULT_COL_WIDTH)
    }

    /// `None` resets the column to the default width.
    pub fn set_col_width(&mut self, col: u32, width: Option<f64>) {
        match width {
            Some(w) if w != DEFAULT_COL_WIDTH => self.col_widths.insert(col, w.clamp(8.0, 2000.0)),
            _ => self.col_widths.remove(&col),
        };
    }

    /// Columns with a custom width, ascending.
    pub fn col_widths(&self) -> impl Iterator<Item = (u32, f64)> + '_ {
        self.col_widths.iter().map(|(c, w)| (*c, *w))
    }

    /// Rows whose automatic height may exceed the default: they hold a cell
    /// with Wrap Text and content, a font bigger than the default, or a line
    /// break. The UI measures only these rows (row autofit, like Excel).
    pub fn rows_to_measure(&self) -> Vec<u32> {
        let mut rows: Vec<u32> = self
            .styles
            .iter()
            .filter(|(at, st)| {
                st.font_size.is_some_and(|sz| sz > crate::style::DEFAULT_FONT_SIZE)
                    || (st.wrap && !self.input(**at).is_empty())
            })
            .map(|(at, _)| at.row)
            .chain(self.cells.iter().filter(|(_, c)| c.has_line_break()).map(|(at, _)| at.row))
            .collect();
        rows.sort_unstable();
        rows.dedup();
        rows
    }

    pub fn layout(&self) -> &SheetLayout {
        &self.layout
    }

    pub fn layout_mut(&mut self) -> &mut SheetLayout {
        &mut self.layout
    }

    /// The one notion of "row hidden": hidden by the user or by a filter.
    pub fn is_row_hidden(&self, row: u32) -> bool {
        self.layout.hidden_rows.contains(row) || self.is_filter_hidden(row)
    }

    pub fn is_col_hidden(&self, col: u32) -> bool {
        self.layout.hidden_cols.contains(col)
    }

    /// All non-empty cells, sorted row by row.
    pub fn cells(&self) -> Vec<(CellRef, Cow<'_, str>, &Value)> {
        let mut out: Vec<_> = self.cells.keys().map(|&at| (at, self.input(at), self.value(at))).collect();
        out.sort_by_key(|(at, _, _)| *at);
        out
    }

    /// Bulk-load cells from a file into an empty sheet, then calculate once.
    /// Literal values arrive already typed (`Some(value)`), so they skip the
    /// typed-input parsing that `set_inputs` does; formulas (`None`) are parsed.
    /// This is the fast path for opening large workbooks.
    /// `input` may be `None` when it is just the value written back.
    pub fn load_cells(&mut self, cells: Vec<(CellRef, Option<String>, Option<Value>)>) {
        self.cells.reserve(cells.len());
        let mut formulas = Vec::new();
        for (at, input, value) in cells {
            match value {
                Some(Value::Empty) => {}
                Some(v) => {
                    self.extent = (self.extent.0.max(at.row + 1), self.extent.1.max(at.col + 1));
                    let input = input.map(String::into_boxed_str);
                    self.cells.insert(at, Cell { input, content: Content::Literal(v) });
                }
                None => {
                    let Some(input) = input.filter(|i| !i.is_empty()) else { continue };
                    self.write(at, input);
                    if self.is_formula(at) {
                        formulas.push(at);
                    }
                }
            }
        }
        self.recalc(formulas.into_iter().collect());
    }

    /// Store an input without recalculating (the caller recalculates: the sheet
    /// itself in `set_inputs`, or the workbook for cross-sheet formulas).
    pub(crate) fn write(&mut self, at: CellRef, input: String) {
        if let Some(Cell { content: Content::Formula { precedents, .. }, .. }) = self.cells.remove(&at) {
            for r in &precedents {
                if r.area() > BIG_RANGE {
                    continue;
                }
                for c in r.cells() {
                    if let Some(set) = self.dependents.get_mut(&c) {
                        set.remove(&at);
                        if set.is_empty() {
                            self.dependents.remove(&c);
                        }
                    }
                }
            }
            if precedents.iter().any(|r| r.area() > BIG_RANGE) {
                self.range_dependents.retain(|(_, f)| *f != at);
            }
            self.links.remove(&at);
            self.volatiles.remove(&at);
        }
        if input.is_empty() {
            return;
        }
        let content = match input.strip_prefix('=') {
            Some(src) => {
                let expr = formula::parse(src).unwrap_or_else(Expr::Error);
                let precedents = expr.precedents();
                for r in &precedents {
                    if r.area() > BIG_RANGE {
                        self.range_dependents.push((*r, at));
                    } else {
                        for c in r.cells() {
                            self.dependents.entry(c).or_default().insert(at);
                        }
                    }
                }
                let external: Vec<(String, Range)> =
                    expr.sheet_precedents().into_iter().map(|(s, r)| (s.to_lowercase(), r)).collect();
                if !external.is_empty() {
                    self.links.insert(at);
                }
                // TODAY, NOW, RAND, OFFSET, INDIRECT...: recalculated every time.
                if expr.is_volatile() {
                    self.volatiles.insert(at);
                }
                // Like Excel, =TODAY() shows a date and =PMT(...) money unless the cell has a format.
                if self.style(at).number_format == NumberFormat::General {
                    if let Some(nf) = implied_format(&expr) {
                        let style = Style { number_format: nf, ..self.style(at).clone() };
                        self.set_style(at, style);
                    }
                }
                Content::Formula { expr, precedents, external, value: Value::Empty }
            }
            None if self.style(at).number_format == NumberFormat::Text => Content::Literal(Value::Text(input.clone())),
            None => {
                // Typing "12%" or "₹500" also sets the matching number format, as in Excel.
                if let Some((_, nf)) = format::parse_formatted(&input) {
                    if self.style(at).number_format == NumberFormat::General {
                        let style = Style { number_format: nf, ..self.style(at).clone() };
                        self.set_style(at, style);
                    }
                }
                Content::Literal(Value::from_input(&input))
            }
        };
        // Keep the typed text only when the value alone wouldn't give it back.
        let input = match &content {
            Content::Literal(v) if v.to_input() == input => None,
            _ => Some(input.into_boxed_str()),
        };
        self.cells.insert(at, Cell { input, content });
        self.extent = (self.extent.0.max(at.row + 1), self.extent.1.max(at.col + 1));
    }

    /// Next cell after `from` (row by row; `forward = false` searches backwards,
    /// both wrap around) whose input or displayed text matches `query`.
    pub fn find(&self, query: &str, from: CellRef, forward: bool, opts: FindOptions) -> Option<CellRef> {
        if query.is_empty() {
            return None;
        }
        let mut hits: Vec<CellRef> = self
            .cells
            .iter()
            .filter(|(at, cell)| opts.matches(&cell.input(), query) || opts.matches(&self.display(**at), query))
            .map(|(at, _)| *at)
            .collect();
        hits.sort();
        if forward {
            hits.iter().find(|at| **at > from).or(hits.first()).copied()
        } else {
            hits.iter().rev().find(|at| **at < from).or(hits.last()).copied()
        }
    }

    /// New input for one cell after replacing `query` in it, or `None` if it doesn't match.
    pub fn replace_in(&self, at: CellRef, query: &str, replacement: &str, opts: FindOptions) -> Option<String> {
        let input = self.input(at);
        if query.is_empty() || !opts.matches(&input, query) {
            return None;
        }
        Some(if opts.whole_cell {
            replacement.to_string()
        } else {
            replace_text(&input, query, replacement, opts.match_case)
        })
    }

    /// Edits that replace `query` in every matching cell (in inputs, so formulas too).
    pub fn replace_all(&self, query: &str, replacement: &str, opts: FindOptions) -> Vec<(CellRef, String)> {
        let mut edits: Vec<_> =
            self.cells.keys().filter_map(|at| Some((*at, self.replace_in(*at, query, replacement, opts)?))).collect();
        edits.sort();
        edits
    }

    /// Ctrl+arrow: from `from`, move in direction (d_row, d_col) to the edge of
    /// the current block of data, or to the next filled cell, or to the sheet edge.
    pub fn jump(&self, from: CellRef, d_row: i32, d_col: i32) -> CellRef {
        let (rows, cols) = (MAX_ROWS as i64, MAX_COLS as i64);
        // One cell in the direction of travel, skipping hidden rows and columns.
        let step = |c: CellRef| -> Option<CellRef> {
            let (mut r, mut k) = (c.row as i64, c.col as i64);
            loop {
                (r, k) = (r + d_row as i64, k + d_col as i64);
                if r < 0 || k < 0 || r >= rows || k >= cols {
                    return None;
                }
                let hidden =
                    (d_row != 0 && self.is_row_hidden(r as u32)) || (d_col != 0 && self.is_col_hidden(k as u32));
                if !hidden {
                    return Some(CellRef::new(r as u32, k as u32));
                }
            }
        };
        let filled = |c: CellRef| self.cells.contains_key(&c);
        let Some(next) = step(from) else { return from };
        let mut cur = from;
        if filled(from) && filled(next) {
            // Run to the last filled cell of this block.
            while let Some(n) = step(cur).filter(|n| filled(*n)) {
                cur = n;
            }
            return cur;
        }
        // Skip empty cells to the next filled one; past the data, go to the sheet edge.
        let (extent_rows, extent_cols) = self.extent;
        cur = next;
        loop {
            if filled(cur) {
                return cur;
            }
            let beyond = (d_row > 0 && cur.row >= extent_rows) || (d_col > 0 && cur.col >= extent_cols);
            match step(cur) {
                Some(n) if !beyond => cur = n,
                _ => {
                    return CellRef::new(
                        if d_row > 0 {
                            MAX_ROWS - 1
                        } else if d_row < 0 {
                            0
                        } else {
                            cur.row
                        },
                        if d_col > 0 {
                            MAX_COLS - 1
                        } else if d_col < 0 {
                            0
                        } else {
                            cur.col
                        },
                    )
                }
            }
        }
    }

    // ------------------------------------------------------------ structure

    /// A copy with `count` rows or columns inserted before index `at`.
    pub fn with_inserted(&self, axis: Axis, at: u32, count: u32) -> Sheet {
        let mut out = self.remapped(axis, &refs::insert_map(at, count));
        let limit = if axis == Axis::Row { MAX_ROWS } else { MAX_COLS };
        out.layout = self.layout.inserted(axis == Axis::Row, at, count, limit);
        out
    }

    /// A copy with `count` rows or columns deleted starting at index `at`.
    pub fn with_deleted(&self, axis: Axis, at: u32, count: u32) -> Sheet {
        let mut out = self.remapped(axis, &refs::delete_map(at, count));
        out.layout = self.layout.deleted(axis == Axis::Row, at, count);
        out
    }

    fn remapped(&self, axis: Axis, map: &dyn Fn(u32) -> Option<u32>) -> Sheet {
        let limit = if axis == Axis::Row { MAX_ROWS } else { MAX_COLS };
        let move_cell = |at: CellRef| -> Option<CellRef> {
            match axis {
                Axis::Row => map(at.row).filter(|r| *r < limit).map(|r| CellRef::new(r, at.col)),
                Axis::Col => map(at.col).filter(|c| *c < limit).map(|c| CellRef::new(at.row, c)),
            }
        };
        let mut out = Sheet { tools: self.tools.remapped(axis, map), ..Sheet::new() };
        for (at, style) in &self.styles {
            if let Some(to) = move_cell(*at) {
                out.styles.insert(to, style.clone());
            }
        }
        for (col, width) in &self.col_widths {
            let new_col = if axis == Axis::Col { map(*col).filter(|c| *c < limit) } else { Some(*col) };
            if let Some(c) = new_col {
                out.col_widths.insert(c, *width);
            }
        }
        let edits: Vec<_> = self
            .cells
            .iter()
            .filter_map(|(at, cell)| Some((move_cell(*at)?, refs::remap_input(&cell.input(), axis, map))))
            .collect();
        out.set_inputs(edits);
        out
    }

    /// A copy with the rows of `range` sorted by the values in column `key_col`.
    /// Empty keys go last; formulas move with their rows like in Excel.
    pub fn with_sorted_rows(&self, range: Range, key_col: u32, ascending: bool) -> Sheet {
        let mut rows: Vec<u32> = (range.start.row..=range.end.row).collect();
        rows.sort_by(|a, b| {
            let (va, vb) = (self.value(CellRef::new(*a, key_col)), self.value(CellRef::new(*b, key_col)));
            match (va.is_empty(), vb.is_empty()) {
                (true, true) => Ordering::Equal,
                (true, false) => Ordering::Greater,
                (false, true) => Ordering::Less,
                _ if ascending => compare(va, vb),
                _ => compare(vb, va),
            }
        });
        // rows[new_index] = old row
        let mut new_row_of: HashMap<u32, u32> = HashMap::default();
        for (i, old) in rows.iter().enumerate() {
            new_row_of.insert(*old, range.start.row + i as u32);
        }
        let in_range =
            |at: &CellRef| at.col >= range.start.col && at.col <= range.end.col && new_row_of.contains_key(&at.row);
        let mut out = Sheet {
            col_widths: self.col_widths.clone(),
            layout: self.layout.clone(),
            tools: self.tools.clone(),
            ..Sheet::new()
        };
        for (at, style) in &self.styles {
            let to = if in_range(at) { CellRef::new(new_row_of[&at.row], at.col) } else { *at };
            out.styles.insert(to, style.clone());
        }
        let edits: Vec<_> = self
            .cells
            .iter()
            .map(|(at, cell)| {
                if in_range(at) {
                    let to = CellRef::new(new_row_of[&at.row], at.col);
                    (to, refs::shift_input(&cell.input(), to.row as i64 - at.row as i64, 0))
                } else {
                    (*at, cell.input().into_owned())
                }
            })
            .collect();
        out.set_inputs(edits);
        out
    }

    /// Inputs and styles produced by dragging the fill handle from `source`
    /// for `count` cells in `dir`. See `fill::series` for the pattern rules.
    pub fn fill(&self, source: Range, dir: FillDirection, count: u32, increment_single: bool) -> FillEdits {
        let (mut edits, mut styles) = (Vec::new(), Vec::new());
        let vertical = matches!(dir, FillDirection::Down | FillDirection::Up);
        let backward = matches!(dir, FillDirection::Up | FillDirection::Left);
        let (lanes, len) = if vertical {
            (source.start.col..=source.end.col, source.rows())
        } else {
            (source.start.row..=source.end.row, source.cols())
        };
        for lane in lanes {
            // Source cells in natural order (top->bottom or left->right).
            let cells: Vec<CellRef> = (0..len)
                .map(|k| {
                    if vertical {
                        CellRef::new(source.start.row + k, lane)
                    } else {
                        CellRef::new(lane, source.start.col + k)
                    }
                })
                .collect();
            let inputs: Vec<String> = cells.iter().map(|c| self.input(*c).to_string()).collect();
            let is_date =
                len == 1 && matches!(self.style(cells[0]).number_format, NumberFormat::Date | NumberFormat::DateTime);
            let shift = |input: &str, delta: i64| {
                if vertical {
                    refs::shift_input(input, delta, 0)
                } else {
                    refs::shift_input(input, 0, delta)
                }
            };
            let values = fill::series(&inputs, count as usize, backward, increment_single || is_date, &shift);
            let ordered: Vec<CellRef> = if backward { cells.iter().rev().copied().collect() } else { cells.clone() };
            for (i, input) in values.into_iter().enumerate() {
                let distance = i as i64 + 1;
                // Distance is measured from the source edge nearest the fill.
                let anchor = if backward { cells[0] } else { *cells.last().unwrap() };
                let (r, c) = match dir {
                    FillDirection::Down => (anchor.row as i64 + distance, anchor.col as i64),
                    FillDirection::Up => (anchor.row as i64 - distance, anchor.col as i64),
                    FillDirection::Right => (anchor.row as i64, anchor.col as i64 + distance),
                    FillDirection::Left => (anchor.row as i64, anchor.col as i64 - distance),
                };
                if r < 0 || c < 0 || r >= MAX_ROWS as i64 || c >= MAX_COLS as i64 {
                    break;
                }
                let target = CellRef::new(r as u32, c as u32);
                edits.push((target, input));
                styles.push((target, self.style(ordered[i % ordered.len()]).clone()));
            }
        }
        (edits, styles)
    }

    /// Call `f` for every formula that reads cell `c` on this sheet. (Allocation-free:
    /// recalculation calls this once per formula.)
    fn for_each_dependent(&self, c: CellRef, mut f: impl FnMut(CellRef)) {
        if let Some(set) = self.dependents.get(&c) {
            for d in set {
                f(*d);
            }
        }
        for (r, d) in &self.range_dependents {
            if r.contains(c) {
                f(*d);
            }
        }
    }

    /// Formulas that read cell `c` on this sheet.
    pub(crate) fn dependents_of(&self, c: CellRef) -> impl Iterator<Item = CellRef> + '_ {
        let big = self.range_dependents.iter().filter(move |(r, _)| r.contains(c)).map(|(_, f)| *f);
        self.dependents.get(&c).into_iter().flatten().copied().chain(big)
    }

    /// Formulas that read other sheets.
    pub(crate) fn links(&self) -> impl Iterator<Item = CellRef> + '_ {
        self.links.iter().copied()
    }

    /// Formulas with volatile functions.
    pub(crate) fn volatiles(&self) -> impl Iterator<Item = CellRef> + '_ {
        self.volatiles.iter().copied()
    }

    /// Ranges on other sheets a formula reads: (sheet name in lower case, range).
    pub(crate) fn external_precedents(&self, at: CellRef) -> &[(String, Range)] {
        match self.cells.get(&at) {
            Some(Cell { content: Content::Formula { external, .. }, .. }) => external,
            _ => &[],
        }
    }

    pub(crate) fn formula_expr(&self, at: CellRef) -> Option<&Expr> {
        match self.cells.get(&at) {
            Some(Cell { content: Content::Formula { expr, .. }, .. }) => Some(expr),
            _ => None,
        }
    }

    fn recalc_from(&mut self, changed: &[CellRef]) {
        let mut dirty: HashSet<CellRef> = changed.iter().copied().filter(|&c| self.is_formula(c)).collect();
        dirty.extend(self.volatiles.iter().copied());
        let mut stack: Vec<CellRef> = dirty.iter().copied().chain(changed.iter().copied()).collect();
        while let Some(c) = stack.pop() {
            let next: Vec<CellRef> = self.dependents_of(c).filter(|d| !dirty.contains(d)).collect();
            for d in next {
                if dirty.insert(d) {
                    stack.push(d);
                }
            }
        }
        self.recalc(dirty);
    }

    fn recalc(&mut self, dirty: HashSet<CellRef>) {
        // In-degree = how many dirty formulas this formula still waits on.
        let mut waiting: HashMap<CellRef, usize> = dirty.iter().map(|&c| (c, 0)).collect();
        for p in &dirty {
            self.for_each_dependent(*p, |f| {
                if let Some(n) = waiting.get_mut(&f) {
                    *n += 1;
                }
            });
        }
        let mut ready: Vec<CellRef> = waiting.iter().filter(|(_, &n)| n == 0).map(|(&c, _)| c).collect();
        while let Some(c) = ready.pop() {
            waiting.remove(&c);
            self.evaluate(c);
            self.for_each_dependent(c, |f| {
                if let Some(n) = waiting.get_mut(&f) {
                    *n -= 1;
                    if *n == 0 {
                        ready.push(f);
                    }
                }
            });
        }
        // Whatever is still waiting is in (or downstream of) a cycle.
        for c in waiting.into_keys() {
            self.store(c, Value::Error(crate::ErrorKind::Circ));
        }
    }

    fn evaluate(&mut self, at: CellRef) {
        let value = match self.cells.get(&at) {
            Some(Cell { content: Content::Formula { expr, .. }, .. }) => formula::eval_at(expr, self, at).0,
            _ => return,
        };
        self.store(at, value);
    }

    pub(crate) fn store(&mut self, at: CellRef, v: Value) {
        if let Some(Cell { content: Content::Formula { value, .. }, .. }) = self.cells.get_mut(&at) {
            *value = v;
        }
    }
}

impl CellSource for Sheet {
    fn value(&self, at: CellRef) -> Value {
        Sheet::value(self, at).clone()
    }

    fn used_extent(&self) -> (u32, u32) {
        self.extent
    }

    fn is_subtotal_cell(&self, sheet: Option<&str>, at: CellRef) -> bool {
        sheet.is_none() && self.formula_expr(at).is_some_and(|e| e.calls("SUBTOTAL") || e.calls("AGGREGATE"))
    }

    fn is_row_hidden(&self, sheet: Option<&str>, row: u32) -> bool {
        sheet.is_none() && Sheet::is_row_hidden(self, row)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ErrorKind;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn sheet(cells: &[(&str, &str)]) -> Sheet {
        let mut s = Sheet::new();
        s.set_inputs(cells.iter().map(|(a, i)| (at(a), i.to_string())));
        s
    }

    #[test]
    fn evaluates_and_recalculates() {
        let mut s = sheet(&[("A1", "10"), ("A2", "20"), ("A3", "=A1+A2*2"), ("A4", "=SUM(A1:A3)")]);
        assert_eq!(s.value(at("A3")), &Value::Number(50.0));
        assert_eq!(s.value(at("A4")), &Value::Number(80.0));

        s.set_input(at("A1"), "0");
        assert_eq!(s.value(at("A3")), &Value::Number(40.0));
        assert_eq!(s.value(at("A4")), &Value::Number(60.0));
    }

    #[test]
    fn formulas_can_be_entered_before_their_inputs() {
        let mut s = sheet(&[("B1", "=A1*2")]);
        assert_eq!(s.value(at("B1")), &Value::Number(0.0));
        s.set_input(at("A1"), "21");
        assert_eq!(s.value(at("B1")), &Value::Number(42.0));
    }

    #[test]
    fn replacing_a_formula_drops_old_dependencies() {
        let mut s = sheet(&[("A1", "1"), ("B1", "=A1"), ("C1", "5")]);
        s.set_input(at("B1"), "=C1");
        s.set_input(at("A1"), "999");
        assert_eq!(s.value(at("B1")), &Value::Number(5.0));
        assert!(!s.dependents.contains_key(&at("A1")));
    }

    #[test]
    fn detects_cycles_and_recovers() {
        let mut s = sheet(&[("A1", "=B1"), ("B1", "=A1+1"), ("C1", "=A1")]);
        assert_eq!(s.value(at("A1")), &Value::Error(ErrorKind::Circ));
        assert_eq!(s.value(at("C1")), &Value::Error(ErrorKind::Circ));

        s.set_input(at("B1"), "7");
        assert_eq!(s.value(at("A1")), &Value::Number(7.0));
        assert_eq!(s.value(at("C1")), &Value::Number(7.0));
    }

    #[test]
    fn long_chains_do_not_overflow_the_stack() {
        let mut s = Sheet::new();
        let n = 50_000;
        let mut edits = vec![(CellRef::new(0, 0), "1".to_string())];
        for row in 1..n {
            edits.push((CellRef::new(row, 0), format!("=A{}+1", row)));
        }
        s.set_inputs(edits);
        assert_eq!(s.value(CellRef::new(n - 1, 0)), &Value::Number(n as f64));
    }

    #[test]
    fn bad_formulas_and_clearing() {
        let mut s = sheet(&[("A1", "=1+"), ("A2", "hello")]);
        assert_eq!(s.value(at("A1")), &Value::Error(ErrorKind::Parse));
        assert_eq!(s.input(at("A1")), "=1+");
        s.set_input(at("A2"), "");
        assert_eq!(s.value(at("A2")), &Value::Empty);
        assert_eq!(s.extent(), (2, 1));
    }

    #[test]
    fn typed_formats_set_the_number_format() {
        let mut s = sheet(&[("A1", "12%"), ("A2", "₹1,500"), ("A3", "=A1*A2")]);
        assert_eq!(s.value(at("A3")), &Value::Number(180.0));
        assert_eq!(s.display(at("A1")), "12%");
        assert_eq!(s.display(at("A2")), "₹1,500");
        s.set_style(at("B1"), Style { number_format: NumberFormat::Text, ..Style::default() });
        s.set_input(at("B1"), "00123");
        assert_eq!(s.value(at("B1")), &Value::Text("00123".into()));
    }

    #[test]
    fn inserting_and_deleting_rows_rewrites_formulas() {
        let s = sheet(&[("A1", "1"), ("A2", "2"), ("A3", "=SUM(A1:A2)"), ("B3", "=A2*10")]);
        let ins = s.with_inserted(Axis::Row, 1, 1);
        assert_eq!(ins.input(at("A4")), "=SUM(A1:A3)");
        assert_eq!(ins.value(at("A4")), &Value::Number(3.0));
        assert_eq!(ins.input(at("B4")), "=A3*10");

        let del = s.with_deleted(Axis::Row, 1, 1);
        assert_eq!(del.input(at("A2")), "=SUM(A1:A1)");
        assert_eq!(del.input(at("B2")), "=#REF!*10");
        assert_eq!(del.value(at("B2")), &Value::Error(ErrorKind::Ref));

        let cols = s.with_inserted(Axis::Col, 0, 2);
        assert_eq!(cols.input(at("D3")), "=C2*10");
    }

    #[test]
    fn sorts_rows_and_moves_formulas() {
        let s = sheet(&[("A1", "pear"), ("B1", "3"), ("A2", "apple"), ("B2", "5"), ("A3", ""), ("C1", "=B1*2")]);
        let sorted = s.with_sorted_rows(Range::new(at("A1"), at("C3")), 0, true);
        assert_eq!(sorted.input(at("A1")), "apple");
        assert_eq!(sorted.input(at("A2")), "pear");
        assert_eq!(sorted.input(at("C2")), "=B2*2");
        assert_eq!(sorted.value(at("C2")), &Value::Number(6.0));
        let desc = s.with_sorted_rows(Range::new(at("A1"), at("B2")), 1, false);
        assert_eq!(desc.input(at("A1")), "apple");
    }

    #[test]
    fn lists_rows_to_measure() {
        let mut s = sheet(&[("A1", "plain"), ("A3", "line\nbreak"), ("B5", "wrapped")]);
        s.set_style(at("B5"), Style { wrap: true, ..Style::default() });
        s.set_style(at("C7"), Style { font_size: Some(240), ..Style::default() });
        s.set_style(at("D9"), Style { wrap: true, ..Style::default() }); // wrap but empty
        assert_eq!(s.rows_to_measure(), vec![2, 4, 6]);
    }

    #[test]
    fn formula_results_get_excels_default_formats() {
        let s = sheet(&[("A1", "=TODAY()"), ("A2", "=DATE(2025,1,1)+30"), ("A3", "=PMT(1%,12,-1000)"), ("A4", "=1+1")]);
        assert_eq!(s.style(at("A1")).number_format, NumberFormat::Date);
        assert_eq!(s.display(at("A2")), "31-01-2025");
        assert_eq!(s.display(at("A3")), "₹88.85");
        assert_eq!(s.style(at("A4")).number_format, NumberFormat::General);
    }

    #[test]
    fn finds_cells() {
        let any = FindOptions::default();
        let s = sheet(&[("A1", "Apple"), ("B2", "pineapple"), ("C3", "=1+1"), ("A4", "banana")]);
        assert_eq!(s.find("apple", at("A1"), true, any), Some(at("B2")));
        assert_eq!(s.find("apple", at("B2"), true, any), Some(at("A1")));
        assert_eq!(s.find("APPLE", at("A1"), false, any), Some(at("B2")));
        assert_eq!(s.find("2", at("A1"), true, any), Some(at("C3")));
        assert_eq!(s.find("zzz", at("A1"), true, any), None);
        let case = FindOptions { match_case: true, whole_cell: false };
        assert_eq!(s.find("Apple", at("A1"), true, case), Some(at("A1")));
        let whole = FindOptions { match_case: false, whole_cell: true };
        assert_eq!(s.find("apple", at("A1"), true, whole), Some(at("A1")));
    }

    #[test]
    fn replaces_text() {
        let s = sheet(&[("A1", "Apple pie"), ("A2", "APPLE"), ("A3", "=SUM(A1)"), ("A4", "सेब और सेब")]);
        let edits = s.replace_all("apple", "Mango", FindOptions::default());
        assert_eq!(edits, vec![(at("A1"), "Mango pie".to_string()), (at("A2"), "Mango".to_string())]);
        let whole = FindOptions { match_case: false, whole_cell: true };
        assert_eq!(s.replace_all("apple", "X", whole), vec![(at("A2"), "X".to_string())]);
        assert_eq!(s.replace_in(at("A3"), "A1", "B1", FindOptions::default()), Some("=SUM(B1)".to_string()));
        assert_eq!(s.replace_in(at("A4"), "सेब", "आम", FindOptions::default()), Some("आम और आम".to_string()));
    }

    #[test]
    fn jumps_like_ctrl_arrow() {
        // A1:A3 filled, A4:A5 empty, A6:A7 filled.
        let s = sheet(&[("A1", "1"), ("A2", "2"), ("A3", "3"), ("A6", "6"), ("A7", "7")]);
        assert_eq!(s.jump(at("A1"), 1, 0), at("A3"));
        assert_eq!(s.jump(at("A3"), 1, 0), at("A6"));
        assert_eq!(s.jump(at("A6"), 1, 0), at("A7"));
        assert_eq!(s.jump(at("A7"), 1, 0), CellRef::new(MAX_ROWS - 1, 0));
        assert_eq!(s.jump(at("A7"), -1, 0), at("A6"));
        assert_eq!(s.jump(at("A5"), -1, 0), at("A3"));
        assert_eq!(s.jump(at("A1"), -1, 0), at("A1"));
        assert_eq!(s.jump(at("C2"), 0, -1), at("A2"));
        assert_eq!(s.jump(at("B1"), 0, 1), CellRef::new(0, MAX_COLS - 1));
    }

    #[test]
    fn fill_handle_extends_series() {
        let s = sheet(&[("A1", "1"), ("A2", "2"), ("A3", "3"), ("B1", "=A1*2"), ("C1", "Item 1")]);
        let (edits, _) = s.fill(Range::new(at("A1"), at("A3")), FillDirection::Down, 2, false);
        assert_eq!(edits, vec![(at("A4"), "4".to_string()), (at("A5"), "5".to_string())]);
        let (edits, _) = s.fill(Range::single(at("B1")), FillDirection::Down, 2, false);
        assert_eq!(edits[1], (at("B3"), "=A3*2".to_string()));
        let (edits, _) = s.fill(Range::single(at("C1")), FillDirection::Right, 1, false);
        assert_eq!(edits, vec![(at("D1"), "Item 2".to_string())]);
        let (edits, _) = s.fill(Range::new(at("A1"), at("A2")), FillDirection::Up, 1, false);
        assert!(edits.is_empty(), "nothing above row 1");
        let (edits, _) = s.fill(Range::new(at("A2"), at("A3")), FillDirection::Up, 1, false);
        assert_eq!(edits, vec![(at("A1"), "1".to_string())]);
    }
}
