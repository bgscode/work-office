//! A workbook: named sheets plus the calculation that spans them.
//!
//! Each `Sheet` knows the formulas that read its own cells. Formulas that read
//! other sheets (`Sheet2!A1`, `'My Data'!B:B`) are "links"; the workbook keeps
//! them consistent: after an edit it finds every formula that depends on the
//! changed cells — on the same sheet or through links — and evaluates them in
//! dependency order with the whole workbook visible. Volatile formulas (TODAY,
//! RAND, OFFSET, INDIRECT...) are evaluated on every recalculation, like Excel.

use rustc_hash::{FxHashMap, FxHashSet};

use crate::cell::{CellRef, Range};
use crate::formula::refs::same_sheet;
use crate::formula::{self, CellSource};
use crate::sheet::Sheet;
use crate::value::{ErrorKind, Value};

/// A formula cell anywhere in the workbook: (sheet index, cell).
type Node = (usize, CellRef);

#[derive(Debug, Clone)]
pub struct Worksheet {
    pub name: String,
    pub sheet: Sheet,
}

/// A file: an ordered list of named sheets.
#[derive(Debug, Clone)]
pub struct Workbook {
    pub sheets: Vec<Worksheet>,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

impl Workbook {
    /// A workbook with one empty "Sheet1".
    pub fn new() -> Self {
        Workbook { sheets: vec![Worksheet { name: "Sheet1".into(), sheet: Sheet::new() }] }
    }

    /// An empty workbook, for importers that add their own sheets.
    pub fn empty() -> Self {
        Workbook { sheets: Vec::new() }
    }

    /// Append a sheet named "SheetN" (first free N) and return its index.
    pub fn add_sheet(&mut self) -> usize {
        let name = (1..)
            .map(|n| format!("Sheet{n}"))
            .find(|name| !self.sheets.iter().any(|s| s.name.eq_ignore_ascii_case(name)))
            .unwrap();
        self.push_sheet(name, Sheet::new())
    }

    pub fn push_sheet(&mut self, name: String, sheet: Sheet) -> usize {
        self.sheets.push(Worksheet { name, sheet });
        self.sheets.len() - 1
    }

    /// Inputs of link formulas (on every sheet but `skip`) that `rewrite` changes,
    /// per sheet: what a rename, a deleted sheet, or rows moving on another
    /// sheet does to formulas that point there.
    pub fn link_edits(
        &self,
        skip: Option<usize>,
        rewrite: impl Fn(&str) -> String,
    ) -> Vec<(usize, Vec<(CellRef, String)>)> {
        let mut out = Vec::new();
        for (s, ws) in self.sheets.iter().enumerate() {
            if Some(s) == skip {
                continue;
            }
            let edits: Vec<(CellRef, String)> = ws
                .sheet
                .links()
                .filter_map(|at| {
                    let input = ws.sheet.formula_input(at)?;
                    let new = rewrite(input);
                    (new != input).then_some((at, new))
                })
                .collect();
            if !edits.is_empty() {
                out.push((s, edits));
            }
        }
        out
    }

    /// The current inputs of the cells that `edits` would change (to undo them).
    pub fn inputs_before(&self, edits: &[(usize, Vec<(CellRef, String)>)]) -> Vec<(usize, Vec<(CellRef, String)>)> {
        edits
            .iter()
            .map(|(s, e)| {
                let sheet = &self.sheets[*s].sheet;
                (*s, e.iter().map(|(at, _)| (*at, sheet.input(*at).into_owned())).collect())
            })
            .collect()
    }

    /// Index of the sheet with this name (case-insensitive, like Excel).
    pub fn sheet_index(&self, name: &str) -> Option<usize> {
        self.sheets.iter().position(|ws| same_sheet(&ws.name, name))
    }

    /// Set inputs on one sheet and recalculate everything that depends on them,
    /// on any sheet.
    pub fn set_inputs(&mut self, sheet: usize, edits: impl IntoIterator<Item = (CellRef, String)>) {
        let mut changed = Vec::new();
        {
            let s = &mut self.sheets[sheet].sheet;
            for (at, input) in edits {
                s.write(at, input);
                changed.push((sheet, at));
            }
        }
        self.recalc(changed);
    }

    /// Recalculate links and volatile formulas everywhere (after opening a file,
    /// or after a sheet was replaced, renamed or removed).
    pub fn recalc_links(&mut self) {
        self.recalc(Vec::new());
    }

    /// Recalculate formulas that depend on `changed` (plus every link and volatile
    /// formula when `changed` is empty).
    fn recalc(&mut self, changed: Vec<Node>) {
        let links = self.link_index();
        let everything = changed.is_empty();
        if everything && links.is_empty() && self.sheets.iter().all(|ws| ws.sheet.volatiles().next().is_none()) {
            return;
        }
        let deps = |book: &Workbook, (s, c): Node| -> Vec<Node> {
            let mut out: Vec<Node> = book.sheets[s].sheet.dependents_of(c).map(|f| (s, f)).collect();
            if let Some(list) = links.get(&s) {
                out.extend(list.iter().filter(|(r, _)| r.contains(c)).map(|(_, f)| *f));
            }
            out
        };

        // Formulas to recalculate: the changed formulas, links (when recalculating
        // everything) and volatiles, plus everything downstream of them.
        let mut dirty: FxHashSet<Node> = FxHashSet::default();
        let mut stack: Vec<Node> = Vec::new();
        for &(s, c) in &changed {
            if self.sheets[s].sheet.is_formula(c) {
                dirty.insert((s, c));
            }
            stack.push((s, c));
        }
        for (s, ws) in self.sheets.iter().enumerate() {
            for f in ws.sheet.volatiles() {
                dirty.insert((s, f));
                stack.push((s, f));
            }
            if everything {
                for f in ws.sheet.links() {
                    dirty.insert((s, f));
                    stack.push((s, f));
                }
            }
        }
        while let Some(n) = stack.pop() {
            for d in deps(self, n) {
                if dirty.insert(d) {
                    stack.push(d);
                }
            }
        }

        // Kahn's algorithm over the dirty formulas; leftovers are in a cycle.
        let mut waiting: FxHashMap<Node, usize> = dirty.iter().map(|n| (*n, 0)).collect();
        for n in &dirty {
            for d in deps(self, *n) {
                if let Some(count) = waiting.get_mut(&d) {
                    *count += 1;
                }
            }
        }
        let mut ready: Vec<Node> = waiting.iter().filter(|(_, n)| **n == 0).map(|(k, _)| *k).collect();
        while let Some(n) = ready.pop() {
            waiting.remove(&n);
            self.evaluate(n);
            for d in deps(self, n) {
                if let Some(count) = waiting.get_mut(&d) {
                    *count -= 1;
                    if *count == 0 {
                        ready.push(d);
                    }
                }
            }
        }
        for (s, c) in waiting.into_keys() {
            self.sheets[s].sheet.store(c, Value::Error(ErrorKind::Circ));
        }
    }

    /// For each sheet, the link formulas that read it: sheet index -> [(range read, formula)].
    fn link_index(&self) -> FxHashMap<usize, Vec<(Range, Node)>> {
        let mut out: FxHashMap<usize, Vec<(Range, Node)>> = FxHashMap::default();
        for (s, ws) in self.sheets.iter().enumerate() {
            for f in ws.sheet.links() {
                for (name, range) in ws.sheet.external_precedents(f) {
                    if let Some(target) = self.sheet_index(name) {
                        out.entry(target).or_default().push((*range, (s, f)));
                    }
                }
            }
        }
        out
    }

    fn evaluate(&mut self, (s, at): Node) {
        let value = {
            let Some(expr) = self.sheets[s].sheet.formula_expr(at) else { return };
            formula::eval_at(expr, &BookSource { book: self, sheet: s }, at).0
        };
        self.sheets[s].sheet.store(at, value);
    }
}

/// Lets a formula on one sheet read every sheet of the workbook.
struct BookSource<'a> {
    book: &'a Workbook,
    sheet: usize,
}

impl BookSource<'_> {
    fn named(&self, name: Option<&str>) -> Option<&Sheet> {
        match name {
            None => Some(&self.book.sheets[self.sheet].sheet),
            Some(n) => self.book.sheet_index(n).map(|i| &self.book.sheets[i].sheet),
        }
    }
}

impl CellSource for BookSource<'_> {
    fn value(&self, at: CellRef) -> Value {
        self.book.sheets[self.sheet].sheet.value(at).clone()
    }

    fn has_sheet(&self, sheet: &str) -> bool {
        self.book.sheet_index(sheet).is_some()
    }

    fn sheet_value(&self, sheet: &str, at: CellRef) -> Value {
        self.named(Some(sheet)).map_or(Value::Error(ErrorKind::Ref), |s| s.value(at).clone())
    }

    fn used_extent(&self) -> (u32, u32) {
        self.book.sheets.iter().fold((0, 0), |(r, c), ws| {
            let (er, ec) = ws.sheet.extent();
            (r.max(er), c.max(ec))
        })
    }

    fn is_subtotal_cell(&self, sheet: Option<&str>, at: CellRef) -> bool {
        self.named(sheet).and_then(|s| s.formula_expr(at)).is_some_and(|e| e.calls("SUBTOTAL") || e.calls("AGGREGATE"))
    }

    fn is_row_hidden(&self, sheet: Option<&str>, row: u32) -> bool {
        self.named(sheet).is_some_and(|s| s.is_row_hidden(row))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn book(sheets: &[(&str, &[(&str, &str)])]) -> Workbook {
        let mut wb = Workbook::empty();
        for (name, cells) in sheets {
            let i = wb.push_sheet(name.to_string(), Sheet::new());
            wb.set_inputs(i, cells.iter().map(|(a, v)| (CellRef::parse(a).unwrap(), v.to_string())));
        }
        wb.recalc_links();
        wb
    }

    fn val(wb: &Workbook, sheet: usize, a: &str) -> Value {
        wb.sheets[sheet].sheet.value(CellRef::parse(a).unwrap()).clone()
    }

    #[test]
    fn formulas_read_other_sheets_and_update() {
        let mut wb = book(&[
            ("Summary", &[("A1", "=Data!A1*2"), ("A2", "=SUM('Sales 2025'!B:B)"), ("A3", "=A1+A2")]),
            ("Data", &[("A1", "21")]),
            ("Sales 2025", &[("B1", "5"), ("B2", "7")]),
        ]);
        assert_eq!(val(&wb, 0, "A1"), Value::Number(42.0));
        assert_eq!(val(&wb, 0, "A2"), Value::Number(12.0));
        assert_eq!(val(&wb, 0, "A3"), Value::Number(54.0));

        // A change on another sheet flows through links and same-sheet formulas.
        wb.set_inputs(1, [(CellRef::parse("A1").unwrap(), "100".to_string())]);
        assert_eq!(val(&wb, 0, "A1"), Value::Number(200.0));
        assert_eq!(val(&wb, 0, "A3"), Value::Number(212.0));
        wb.set_inputs(2, [(CellRef::parse("B500").unwrap(), "88".to_string())]);
        assert_eq!(val(&wb, 0, "A2"), Value::Number(100.0));
    }

    #[test]
    fn missing_sheets_and_cycles() {
        let wb = book(&[("A", &[("A1", "=Nope!A1"), ("B1", "=B!A1")]), ("B", &[("A1", "=A!B1")])]);
        assert_eq!(val(&wb, 0, "A1"), Value::Error(ErrorKind::Ref));
        assert_eq!(val(&wb, 0, "B1"), Value::Error(ErrorKind::Circ));
    }

    #[test]
    fn row_and_whole_column_references() {
        let mut wb = book(&[("S", &[("A1", "1"), ("A2", "2"), ("B5", "=ROW()"), ("C1", "=SUM(A:A)")])]);
        assert_eq!(val(&wb, 0, "B5"), Value::Number(5.0));
        assert_eq!(val(&wb, 0, "C1"), Value::Number(3.0));
        wb.set_inputs(0, [(CellRef::parse("A1000").unwrap(), "10".to_string())]);
        assert_eq!(val(&wb, 0, "C1"), Value::Number(13.0));
    }

    #[test]
    fn names_new_sheets() {
        let mut wb = Workbook::new();
        assert_eq!(wb.add_sheet(), 1);
        assert_eq!(wb.sheets[1].name, "Sheet2");
    }
}
