//! Undoable edits. Each implements `office_core::Op` so it works with
//! `office_core::history::History`.

use office_core::Op;

use crate::cell::CellRef;
use crate::formula::refs;
use crate::sheet::Sheet;
use crate::style::Style;
use crate::workbook::{Workbook, Worksheet};

/// Set the input of many cells on one sheet at once (typing, paste, delete).
#[derive(Debug, Clone)]
pub struct SetCells {
    pub sheet: usize,
    pub edits: Vec<(CellRef, String)>,
}

impl Op<Workbook> for SetCells {
    fn apply(&self, doc: &mut Workbook) {
        if self.sheet < doc.sheets.len() {
            doc.set_inputs(self.sheet, self.edits.iter().cloned());
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let sheet = &doc.sheets[self.sheet].sheet;
        let edits = self.edits.iter().map(|(at, _)| (*at, sheet.input(*at).to_string())).collect();
        Box::new(SetCells { sheet: self.sheet, edits })
    }
}

/// Set the formatting of many cells on one sheet.
#[derive(Debug, Clone)]
pub struct SetStyles {
    pub sheet: usize,
    pub styles: Vec<(CellRef, Style)>,
}

impl Op<Workbook> for SetStyles {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            for (at, style) in &self.styles {
                ws.sheet.set_style(*at, style.clone());
            }
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let sheet = &doc.sheets[self.sheet].sheet;
        let styles = self.styles.iter().map(|(at, _)| (*at, sheet.style(*at).clone())).collect();
        Box::new(SetStyles { sheet: self.sheet, styles })
    }
}

/// Set column widths; `None` restores the default width.
#[derive(Debug, Clone)]
pub struct SetColWidths {
    pub sheet: usize,
    pub widths: Vec<(u32, Option<f64>)>,
}

impl Op<Workbook> for SetColWidths {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            for (col, w) in &self.widths {
                ws.sheet.set_col_width(*col, *w);
            }
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let sheet = &doc.sheets[self.sheet].sheet;
        let widths = self
            .widths
            .iter()
            .map(|(col, _)| {
                let w = sheet.col_width(*col);
                (*col, (w != crate::sheet::DEFAULT_COL_WIDTH).then_some(w))
            })
            .collect();
        Box::new(SetColWidths { sheet: self.sheet, widths })
    }
}

/// Swap in a whole new version of a sheet. Used for structural edits (insert
/// or delete rows/columns, sort) where computing a precise inverse is harder
/// than keeping the old sheet.
#[derive(Debug, Clone)]
pub struct ReplaceSheet {
    pub sheet: usize,
    pub with: Box<Sheet>,
    /// New inputs for formulas on other sheets that point at this one (rows or
    /// columns moved), applied in the same undo step.
    pub links: Vec<(usize, Vec<(CellRef, String)>)>,
}

impl ReplaceSheet {
    pub fn new(sheet: usize, with: Sheet) -> Self {
        ReplaceSheet { sheet, with: Box::new(with), links: Vec::new() }
    }
}

impl Op<Workbook> for ReplaceSheet {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            ws.sheet = (*self.with).clone();
        }
        for (s, edits) in &self.links {
            doc.set_inputs(*s, edits.iter().cloned());
        }
        doc.recalc_links();
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(ReplaceSheet {
            sheet: self.sheet,
            with: Box::new(doc.sheets[self.sheet].sheet.clone()),
            links: doc.inputs_before(&self.links),
        })
    }
}

/// Insert a sheet at `index`.
#[derive(Debug, Clone)]
pub struct InsertSheet {
    pub index: usize,
    pub sheet: Worksheet,
}

impl Op<Workbook> for InsertSheet {
    fn apply(&self, doc: &mut Workbook) {
        let index = self.index.min(doc.sheets.len());
        doc.sheets.insert(index, self.sheet.clone());
        // Sheet-scoped names keep pointing at their sheets.
        for d in &mut doc.names {
            if let Some(s) = d.scope.as_mut().filter(|s| **s >= index) {
                *s += 1;
            }
        }
        doc.recalc_links();
    }

    fn invert(&self, _doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(RemoveSheet { index: self.index })
    }
}

/// Remove the sheet at `index`.
#[derive(Debug, Clone)]
pub struct RemoveSheet {
    pub index: usize,
}

impl Op<Workbook> for RemoveSheet {
    /// Removes the sheet; references to it elsewhere become `#REF!`, as in Excel.
    fn apply(&self, doc: &mut Workbook) {
        if self.index >= doc.sheets.len() {
            return;
        }
        let name = doc.sheets.remove(self.index).name;
        for (s, edits) in doc.link_edits(None, |input| refs::delete_sheet_input(input, &name)) {
            doc.set_inputs(s, edits);
        }
        // Names scoped to the sheet go; names pointing at it become #REF!.
        let index = self.index;
        let mut names = doc.names_rewritten(|t| refs::delete_sheet_input(t, &name));
        names.retain(|d| d.scope != Some(index));
        for d in &mut names {
            if let Some(s) = d.scope.as_mut().filter(|s| **s > index) {
                *s -= 1;
            }
        }
        doc.names = names;
        doc.recalc_links();
    }

    /// Undo puts the sheet back and restores the formulas that became `#REF!`.
    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let name = doc.sheets[self.index].name.clone();
        let changed = doc.link_edits(Some(self.index), |input| refs::delete_sheet_input(input, &name));
        let restore: Vec<Box<dyn Op<Workbook>>> = doc
            .inputs_before(&changed)
            .into_iter()
            .map(|(s, edits)| Box::new(SetCells { sheet: s, edits }) as Box<dyn Op<Workbook>>)
            .collect();
        let mut ops: Vec<Box<dyn Op<Workbook>>> =
            vec![Box::new(InsertSheet { index: self.index, sheet: doc.sheets[self.index].clone() })];
        ops.extend(restore);
        ops.push(Box::new(crate::names::SetNames { names: doc.names.clone() }));
        Box::new(Sequence(ops))
    }
}

#[derive(Debug, Clone)]
pub struct RenameSheet {
    pub index: usize,
    pub name: String,
}

impl Op<Workbook> for RenameSheet {
    /// Renames the sheet and rewrites formulas that point at it, as Excel does.
    fn apply(&self, doc: &mut Workbook) {
        let Some(ws) = doc.sheets.get_mut(self.index) else { return };
        let old = std::mem::replace(&mut ws.name, self.name.clone());
        for (s, edits) in doc.link_edits(None, |input| refs::rename_sheet_input(input, &old, &self.name)) {
            doc.set_inputs(s, edits);
        }
        doc.names = doc.names_rewritten(|t| refs::rename_sheet_input(t, &old, &self.name));
        doc.recalc_links();
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(RenameSheet { index: self.index, name: doc.sheets[self.index].name.clone() })
    }
}

/// Several ops as one undo step. Each op's inverse is computed against the
/// state before the batch, so the ops must touch different things (for
/// example cell inputs and cell styles).
pub struct Batch(pub Vec<Box<dyn Op<Workbook>>>);

impl Op<Workbook> for Batch {
    fn apply(&self, doc: &mut Workbook) {
        for op in &self.0 {
            op.apply(doc);
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(Batch(self.0.iter().rev().map(|op| op.invert(doc)).collect()))
    }
}

/// Ops applied in order, undone in reverse. Unlike `Batch`, each inverse is
/// computed when needed, so later ops may depend on earlier ones (an inserted
/// sheet, then its restored formulas).
pub struct Sequence(pub Vec<Box<dyn Op<Workbook>>>);

impl Op<Workbook> for Sequence {
    fn apply(&self, doc: &mut Workbook) {
        for op in &self.0 {
            op.apply(doc);
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        // Inverting step by step needs the intermediate states.
        let mut scratch = doc.clone();
        let mut inverses = Vec::new();
        for op in &self.0 {
            inverses.push(op.invert(&scratch));
            op.apply(&mut scratch);
        }
        inverses.reverse();
        Box::new(Sequence(inverses))
    }
}

#[cfg(test)]
mod tests {
    use office_core::history::History;

    use super::*;
    use crate::Value;
    use crate::{Sheet, Style};

    #[test]
    fn undo_and_redo() {
        let mut wb = Workbook::new();
        let mut history = History::default();
        let a1 = CellRef::new(0, 0);
        let b1 = CellRef::new(0, 1);
        history.apply(&mut wb, Box::new(SetCells { sheet: 0, edits: vec![(a1, "5".into()), (b1, "=A1*2".into())] }));
        history.apply(&mut wb, Box::new(SetCells { sheet: 0, edits: vec![(a1, "7".into())] }));
        assert_eq!(wb.sheets[0].sheet.value(b1), &Value::Number(14.0));

        assert!(history.undo(&mut wb));
        assert_eq!(wb.sheets[0].sheet.value(b1), &Value::Number(10.0));
        assert!(history.undo(&mut wb));
        assert_eq!(wb.sheets[0].sheet.input(b1), "");
        assert!(!history.undo(&mut wb));

        assert!(history.redo(&mut wb));
        assert!(history.redo(&mut wb));
        assert_eq!(wb.sheets[0].sheet.value(b1), &Value::Number(14.0));
    }

    #[test]
    fn renaming_and_removing_sheets_rewrites_formulas() {
        let a1 = CellRef::new(0, 0);
        let mut wb = Workbook::new();
        let data = wb.add_sheet();
        wb.sheets[data].name = "Data".into();
        wb.set_inputs(data, [(a1, "5".to_string())]);
        wb.set_inputs(0, [(a1, "=Data!A1*2".to_string())]);
        let mut history = History::default();

        history.apply(&mut wb, Box::new(RenameSheet { index: data, name: "My Data".into() }));
        assert_eq!(wb.sheets[0].sheet.input(a1), "='My Data'!A1*2");
        assert_eq!(wb.sheets[0].sheet.value(a1), &Value::Number(10.0));
        history.undo(&mut wb);
        assert_eq!(wb.sheets[0].sheet.input(a1), "=Data!A1*2");

        history.apply(&mut wb, Box::new(RemoveSheet { index: data }));
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].sheet.input(a1), "=#REF!*2");
        assert_eq!(wb.sheets[0].sheet.value(a1), &Value::Error(crate::ErrorKind::Ref));
        history.undo(&mut wb);
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[0].sheet.input(a1), "=Data!A1*2");
        assert_eq!(wb.sheets[0].sheet.value(a1), &Value::Number(10.0));
    }

    #[test]
    fn structural_ops_undo() {
        let mut wb = Workbook::new();
        let mut history = History::default();
        let a1 = CellRef::new(0, 0);
        history.apply(&mut wb, Box::new(SetCells { sheet: 0, edits: vec![(a1, "x".into())] }));
        let bold = Style { bold: true, ..Style::default() };
        history.apply(
            &mut wb,
            Box::new(Batch(vec![
                Box::new(SetStyles { sheet: 0, styles: vec![(a1, bold.clone())] }),
                Box::new(SetColWidths { sheet: 0, widths: vec![(0, Some(150.0))] }),
            ])),
        );
        let inserted = wb.sheets[0].sheet.with_inserted(crate::formula::refs::Axis::Row, 0, 1);
        history.apply(&mut wb, Box::new(ReplaceSheet::new(0, inserted)));
        assert_eq!(wb.sheets[0].sheet.input(CellRef::new(1, 0)), "x");
        assert_eq!(wb.sheets[0].sheet.style(CellRef::new(1, 0)), &bold);

        history.undo(&mut wb);
        assert_eq!(wb.sheets[0].sheet.input(a1), "x");
        history.undo(&mut wb);
        assert!(wb.sheets[0].sheet.style(a1).is_default());
        assert_eq!(wb.sheets[0].sheet.col_width(0), crate::sheet::DEFAULT_COL_WIDTH);

        history.apply(&mut wb, Box::new(InsertSheet { index: 1, sheet: Worksheet::new("New".into(), Sheet::new()) }));
        history.apply(&mut wb, Box::new(RenameSheet { index: 1, name: "Renamed".into() }));
        history.apply(&mut wb, Box::new(RemoveSheet { index: 0 }));
        assert_eq!(wb.sheets.len(), 1);
        history.undo(&mut wb);
        history.undo(&mut wb);
        assert_eq!(wb.sheets[1].name, "New");
        assert_eq!(wb.sheets[0].sheet.input(a1), "x");
    }
}
