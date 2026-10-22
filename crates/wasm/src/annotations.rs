//! Bindings for notes (Review → New Note) and hyperlinks (Insert → Link, Ctrl+K).

use sheets_engine::annotations::{Hyperlink, Note, SetHyperlinks, SetNotes};
use sheets_engine::ops::{Batch, SetCells, SetStyles};
use sheets_engine::{CellRef, StylePatch};
use wasm_bindgen::prelude::*;

use crate::{json_string, rect, SheetsDocument};

/// Excel's Hyperlink cell style: blue and underlined.
const LINK_COLOR: &str = "#0563C1";

#[wasm_bindgen]
impl SheetsDocument {
    /// The note on a cell as `{"author","text"}`, or `""` when there is none.
    #[wasm_bindgen(js_name = noteAt)]
    pub fn note_at(&self, sheet: usize, row: u32, col: u32) -> String {
        match self.sheet(sheet).annotations().notes.get(&CellRef::new(row, col)) {
            Some(n) => format!(r#"{{"author":{},"text":{}}}"#, json_string(&n.author), json_string(&n.text)),
            None => String::new(),
        }
    }

    /// Cells with notes in a block, as `[row, col, row, col, …]` (for the red triangles).
    #[wasm_bindgen(js_name = notesBlock)]
    pub fn notes_block(&self, sheet: usize, row: u32, col: u32, rows: u32, cols: u32) -> Vec<u32> {
        let (end_row, end_col) = (row.saturating_add(rows), col.saturating_add(cols));
        self.sheet(sheet)
            .annotations()
            .notes
            .range(CellRef::new(row, 0)..CellRef::new(end_row, 0))
            .filter(|(at, _)| at.col >= col && at.col < end_col)
            .flat_map(|(at, _)| [at.row, at.col])
            .collect()
    }

    /// Every note on the sheet in order (Review → Next/Previous Note), as `[row, col, …]`.
    #[wasm_bindgen(js_name = noteCells)]
    pub fn note_cells(&self, sheet: usize) -> Vec<u32> {
        self.sheet(sheet).annotations().notes.keys().flat_map(|at| [at.row, at.col]).collect()
    }

    /// Add or change a note; empty text removes it.
    #[wasm_bindgen(js_name = setNote)]
    pub fn set_note(&mut self, sheet: usize, row: u32, col: u32, text: &str, author: &str) {
        let at = CellRef::new(row, col);
        let note = (!text.is_empty()).then(|| Note { author: author.to_string(), text: text.to_string() });
        if self.sheet(sheet).annotations().notes.get(&at) != note.as_ref() {
            self.apply(SetNotes { sheet, edits: vec![(at, note)] });
        }
    }

    /// Remove the notes in a rectangle (Clear → Clear Notes, Review → Delete).
    #[wasm_bindgen(js_name = deleteNotes)]
    pub fn delete_notes(&mut self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32) {
        let r = rect(top, left, bottom, right);
        let notes = &self.sheet(sheet).annotations().notes;
        let edits: Vec<_> = notes.keys().filter(|at| r.contains(**at)).map(|at| (*at, None)).collect();
        if !edits.is_empty() {
            self.apply(SetNotes { sheet, edits });
        }
    }

    /// The link on a cell as `{"target","tooltip"}`, or `""`.
    #[wasm_bindgen(js_name = hyperlinkAt)]
    pub fn hyperlink_at(&self, sheet: usize, row: u32, col: u32) -> String {
        match self.sheet(sheet).annotations().links.get(&CellRef::new(row, col)) {
            Some(l) => format!(r#"{{"target":{},"tooltip":{}}}"#, json_string(&l.target), json_string(&l.tooltip)),
            None => String::new(),
        }
    }

    /// Insert → Link on a rectangle, one undo step: the link, the text to
    /// display in the cells (`display`; empty keeps what is there, or shows the
    /// address in empty cells) and Excel's blue underlined link style.
    #[allow(clippy::too_many_arguments)]
    #[wasm_bindgen(js_name = setHyperlink)]
    pub fn set_hyperlink(
        &mut self,
        sheet: usize,
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
        target: &str,
        display: &str,
        tooltip: &str,
    ) {
        let s = self.sheet(sheet);
        let cells: Vec<CellRef> = rect(top, left, bottom, right).cells().collect();
        let link = Hyperlink { target: target.to_string(), tooltip: tooltip.to_string() };
        let shown = if display.is_empty() { target.trim_start_matches('#') } else { display };
        let texts: Vec<(CellRef, String)> = cells
            .iter()
            .filter(|at| !display.is_empty() || s.input(**at).is_empty())
            .map(|at| (*at, shown.to_string()))
            .collect();
        let patches = [StylePatch::Color(Some(LINK_COLOR.into())), StylePatch::Underline(true)];
        let styles =
            cells.iter().map(|at| (*at, patches.iter().fold(s.style(*at).clone(), |st, p| st.patched(p)))).collect();
        let mut ops: Vec<Box<dyn office_core::Op<sheets_engine::Workbook>>> = vec![
            Box::new(SetHyperlinks { sheet, edits: cells.iter().map(|at| (*at, Some(link.clone()))).collect() }),
            Box::new(SetStyles { sheet, styles }),
        ];
        if !texts.is_empty() {
            ops.push(Box::new(SetCells { sheet, edits: texts }));
        }
        self.apply(Batch(ops));
    }

    /// Remove the links in a rectangle. `keep_format` false also clears the
    /// link colour and underline (Excel's Clear Hyperlinks vs. Remove Hyperlinks).
    #[wasm_bindgen(js_name = removeHyperlinks)]
    pub fn remove_hyperlinks(&mut self, sheet: usize, top: u32, left: u32, bottom: u32, right: u32, keep_format: bool) {
        let r = rect(top, left, bottom, right);
        let s = self.sheet(sheet);
        let cells: Vec<CellRef> = s.annotations().links.keys().filter(|at| r.contains(**at)).copied().collect();
        if cells.is_empty() {
            return;
        }
        let mut ops: Vec<Box<dyn office_core::Op<sheets_engine::Workbook>>> =
            vec![Box::new(SetHyperlinks { sheet, edits: cells.iter().map(|at| (*at, None)).collect() })];
        if !keep_format {
            let patches = [StylePatch::Color(None), StylePatch::Underline(false)];
            let styles = cells
                .iter()
                .map(|at| (*at, patches.iter().fold(s.style(*at).clone(), |st, p| st.patched(p))))
                .collect();
            ops.push(Box::new(SetStyles { sheet, styles }));
        }
        self.apply(Batch(ops));
    }
}
