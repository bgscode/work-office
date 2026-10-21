//! Undoable edits. Each implements `office_core::Op` so it works with
//! `office_core::history::History`.

use office_core::Op;

use crate::cell::CellRef;
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
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            ws.sheet.set_inputs(self.edits.iter().cloned());
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
}

impl Op<Workbook> for ReplaceSheet {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.sheet) {
            ws.sheet = (*self.with).clone();
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(ReplaceSheet { sheet: self.sheet, with: Box::new(doc.sheets[self.sheet].sheet.clone()) })
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
        doc.sheets.insert(self.index.min(doc.sheets.len()), self.sheet.clone());
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
    fn apply(&self, doc: &mut Workbook) {
        if self.index < doc.sheets.len() {
            doc.sheets.remove(self.index);
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(InsertSheet { index: self.index, sheet: doc.sheets[self.index].clone() })
    }
}

#[derive(Debug, Clone)]
pub struct RenameSheet {
    pub index: usize,
    pub name: String,
}

impl Op<Workbook> for RenameSheet {
    fn apply(&self, doc: &mut Workbook) {
        if let Some(ws) = doc.sheets.get_mut(self.index) {
            ws.name = self.name.clone();
        }
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
        history.apply(&mut wb, Box::new(ReplaceSheet { sheet: 0, with: Box::new(inserted) }));
        assert_eq!(wb.sheets[0].sheet.input(CellRef::new(1, 0)), "x");
        assert_eq!(wb.sheets[0].sheet.style(CellRef::new(1, 0)), &bold);

        history.undo(&mut wb);
        assert_eq!(wb.sheets[0].sheet.input(a1), "x");
        history.undo(&mut wb);
        assert!(wb.sheets[0].sheet.style(a1).is_default());
        assert_eq!(wb.sheets[0].sheet.col_width(0), crate::sheet::DEFAULT_COL_WIDTH);

        history.apply(
            &mut wb,
            Box::new(InsertSheet { index: 1, sheet: Worksheet { name: "New".into(), sheet: Sheet::new() } }),
        );
        history.apply(&mut wb, Box::new(RenameSheet { index: 1, name: "Renamed".into() }));
        history.apply(&mut wb, Box::new(RemoveSheet { index: 0 }));
        assert_eq!(wb.sheets.len(), 1);
        history.undo(&mut wb);
        history.undo(&mut wb);
        assert_eq!(wb.sheets[1].name, "New");
        assert_eq!(wb.sheets[0].sheet.input(a1), "x");
    }
}
