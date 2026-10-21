//! Generic undo/redo stack used by all three editors.

use crate::Op;

pub struct History<Doc> {
    undo: Vec<Box<dyn Op<Doc>>>,
    redo: Vec<Box<dyn Op<Doc>>>,
}

impl<Doc> Default for History<Doc> {
    fn default() -> Self {
        Self { undo: Vec::new(), redo: Vec::new() }
    }
}

impl<Doc> History<Doc> {
    pub fn apply(&mut self, doc: &mut Doc, op: Box<dyn Op<Doc>>) {
        let inverse = op.invert(doc);
        op.apply(doc);
        self.undo.push(inverse);
        self.redo.clear();
    }

    pub fn undo(&mut self, doc: &mut Doc) -> bool {
        let Some(op) = self.undo.pop() else { return false };
        let inverse = op.invert(doc);
        op.apply(doc);
        self.redo.push(inverse);
        true
    }

    pub fn redo(&mut self, doc: &mut Doc) -> bool {
        let Some(op) = self.redo.pop() else { return false };
        let inverse = op.invert(doc);
        op.apply(doc);
        self.undo.push(inverse);
        true
    }
}
