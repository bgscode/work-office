//! Things attached to a cell besides its value and format: notes (Review →
//! New Note; a red triangle in the cell's corner) and hyperlinks (Insert →
//! Link, Ctrl+K). They move with their cells when rows or columns are
//! inserted, deleted or sorted, like in Excel.

use std::collections::BTreeMap;

use office_core::Op;

use crate::cell::CellRef;
use crate::formula::refs::Axis;
use crate::workbook::Workbook;

/// A note: plain text and who wrote it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Note {
    pub author: String,
    pub text: String,
}

/// Where a link goes: a web address, `mailto:`, a file, or a place in the
/// workbook written as a reference with a leading `#` (`#Sheet2!A1`, `#Sales`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    pub target: String,
    /// ScreenTip shown on hover; empty for the default.
    pub tooltip: String,
}

impl Hyperlink {
    /// The reference inside the workbook (`Sheet2!A1`), for links that go there.
    pub fn location(&self) -> Option<&str> {
        self.target.strip_prefix('#')
    }
}

/// Notes and hyperlinks of one sheet, by cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Annotations {
    pub notes: BTreeMap<CellRef, Note>,
    pub links: BTreeMap<CellRef, Hyperlink>,
}

impl Annotations {
    pub fn is_empty(&self) -> bool {
        self.notes.is_empty() && self.links.is_empty()
    }

    /// After rows or columns moved: `map` gives each line's new index, or None if it was deleted.
    pub fn remapped(&self, axis: Axis, map: &dyn Fn(u32) -> Option<u32>, limit: u32) -> Annotations {
        let move_cell = |at: &CellRef| match axis {
            Axis::Row => map(at.row).filter(|r| *r < limit).map(|r| CellRef::new(r, at.col)),
            Axis::Col => map(at.col).filter(|c| *c < limit).map(|c| CellRef::new(at.row, c)),
        };
        Annotations {
            notes: self.notes.iter().filter_map(|(at, n)| Some((move_cell(at)?, n.clone()))).collect(),
            links: self.links.iter().filter_map(|(at, l)| Some((move_cell(at)?, l.clone()))).collect(),
        }
    }

    /// Every cell moved to `to(cell)` (after sorting rows).
    pub fn moved(&self, to: impl Fn(&CellRef) -> CellRef) -> Annotations {
        Annotations {
            notes: self.notes.iter().map(|(at, n)| (to(at), n.clone())).collect(),
            links: self.links.iter().map(|(at, l)| (to(at), l.clone())).collect(),
        }
    }
}

/// Add, change or remove notes (None removes).
#[derive(Debug, Clone)]
pub struct SetNotes {
    pub sheet: usize,
    pub edits: Vec<(CellRef, Option<Note>)>,
}

impl Op<Workbook> for SetNotes {
    fn apply(&self, doc: &mut Workbook) {
        let Some(ws) = doc.sheets.get_mut(self.sheet) else { return };
        let notes = &mut ws.sheet.annotations_mut().notes;
        for (at, note) in &self.edits {
            match note {
                Some(n) => notes.insert(*at, n.clone()),
                None => notes.remove(at),
            };
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let notes = &doc.sheets[self.sheet].sheet.annotations().notes;
        let edits = self.edits.iter().map(|(at, _)| (*at, notes.get(at).cloned())).collect();
        Box::new(SetNotes { sheet: self.sheet, edits })
    }
}

/// Add, change or remove hyperlinks (None removes).
#[derive(Debug, Clone)]
pub struct SetHyperlinks {
    pub sheet: usize,
    pub edits: Vec<(CellRef, Option<Hyperlink>)>,
}

impl Op<Workbook> for SetHyperlinks {
    fn apply(&self, doc: &mut Workbook) {
        let Some(ws) = doc.sheets.get_mut(self.sheet) else { return };
        let links = &mut ws.sheet.annotations_mut().links;
        for (at, link) in &self.edits {
            match link {
                Some(l) => links.insert(*at, l.clone()),
                None => links.remove(at),
            };
        }
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        let links = &doc.sheets[self.sheet].sheet.annotations().links;
        let edits = self.edits.iter().map(|(at, _)| (*at, links.get(at).cloned())).collect();
        Box::new(SetHyperlinks { sheet: self.sheet, edits })
    }
}

#[cfg(test)]
mod tests {
    use office_core::history::History;

    use super::*;
    use crate::ops::ReplaceSheet;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    fn note(text: &str) -> Note {
        Note { author: "Asha".into(), text: text.into() }
    }

    #[test]
    fn notes_undo() {
        let mut wb = Workbook::new();
        let mut history = History::default();
        history.apply(&mut wb, Box::new(SetNotes { sheet: 0, edits: vec![(at("B2"), Some(note("Check")))] }));
        assert_eq!(wb.sheets[0].sheet.annotations().notes[&at("B2")].text, "Check");
        history.apply(&mut wb, Box::new(SetNotes { sheet: 0, edits: vec![(at("B2"), None)] }));
        assert!(wb.sheets[0].sheet.annotations().notes.is_empty());
        history.undo(&mut wb);
        history.undo(&mut wb);
        assert!(wb.sheets[0].sheet.annotations().is_empty());
    }

    #[test]
    fn notes_and_links_move_with_their_cells() {
        let mut wb = Workbook::new();
        SetNotes { sheet: 0, edits: vec![(at("B3"), Some(note("x"))), (at("B5"), Some(note("y")))] }.apply(&mut wb);
        let link = Hyperlink { target: "https://example.com".into(), tooltip: String::new() };
        SetHyperlinks { sheet: 0, edits: vec![(at("C4"), Some(link.clone()))] }.apply(&mut wb);
        // Insert two rows above row 4, delete row 3.
        let s = wb.sheets[0].sheet.with_inserted(Axis::Row, 3, 2).with_deleted(Axis::Row, 2, 1);
        ReplaceSheet::new(0, s).apply(&mut wb);
        let a = wb.sheets[0].sheet.annotations();
        assert_eq!(a.notes.keys().copied().collect::<Vec<_>>(), [at("B6")]);
        assert_eq!(a.links.get(&at("C5")), Some(&link));
    }

    #[test]
    fn notes_sort_with_their_rows() {
        let mut wb = Workbook::new();
        wb.set_inputs(0, [(at("A1"), "3".into()), (at("A2"), "1".into())]);
        SetNotes { sheet: 0, edits: vec![(at("A1"), Some(note("three")))] }.apply(&mut wb);
        let range = crate::Range::new(at("A1"), at("A2"));
        let s = wb.sheets[0].sheet.with_sorted_rows(range, 0, true);
        assert_eq!(s.annotations().notes[&at("A2")].text, "three");
    }
}
