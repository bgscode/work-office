//! Bindings for defined names (Formulas → Name Manager, the Name Box).
//! Scopes cross as a sheet index, or -1 for the whole workbook.

use sheets_engine::formula::Expr;
use sheets_engine::names::{DefinedName, SetNames};
use wasm_bindgen::prelude::*;

use crate::{json_string, SheetsDocument};

fn scope_of(i: i32) -> Option<usize> {
    (i >= 0).then_some(i as usize)
}

#[wasm_bindgen]
impl SheetsDocument {
    /// The defined names as JSON: `[{"name","refersTo","scope","comment"}]`, scope -1 = workbook.
    #[wasm_bindgen(js_name = namesJson)]
    pub fn names_json(&self) -> String {
        let items: Vec<String> = self
            .book
            .names
            .iter()
            .map(|d| {
                format!(
                    r#"{{"name":{},"refersTo":{},"scope":{},"comment":{}}}"#,
                    json_string(&d.name),
                    json_string(&d.refers_to),
                    d.scope.map_or(-1, |s| s as i64),
                    json_string(&d.comment)
                )
            })
            .collect();
        format!("[{}]", items.join(","))
    }

    /// Add a name, or change the one called `old_name` in `old_scope` (when
    /// `old_name` isn't empty). Returns "" on success or the reason it was refused.
    #[wasm_bindgen(js_name = defineName)]
    pub fn define_name(
        &mut self,
        name: &str,
        refers_to: &str,
        scope: i32,
        comment: &str,
        old_name: &str,
        old_scope: i32,
    ) -> String {
        if scope_of(scope).is_some_and(|s| s >= self.book.sheets.len()) {
            return "That sheet doesn't exist.".into();
        }
        let mut d = DefinedName::new(name, refers_to, scope_of(scope));
        d.comment = comment.to_string();
        let replacing = (!old_name.is_empty()).then_some((old_name, scope_of(old_scope)));
        match self.book.with_name(d, replacing) {
            Ok(names) => {
                self.apply(SetNames { names });
                String::new()
            }
            Err(e) => e,
        }
    }

    #[wasm_bindgen(js_name = deleteName)]
    pub fn delete_name(&mut self, name: &str, scope: i32) -> bool {
        let names: Vec<DefinedName> = self
            .book
            .names
            .iter()
            .filter(|d| !(d.name.eq_ignore_ascii_case(name) && d.scope == scope_of(scope)))
            .cloned()
            .collect();
        if names.len() == self.book.names.len() {
            return false;
        }
        self.apply(SetNames { names });
        true
    }

    /// Where a name points, as seen from `sheet`: [sheet index, top, left, bottom, right],
    /// or empty when it isn't a range (a constant or formula) or doesn't exist.
    #[wasm_bindgen(js_name = nameTarget)]
    pub fn name_target(&self, name: &str, sheet: usize) -> Vec<u32> {
        let Some(d) = self.book.defined_name(name, sheet) else { return Vec::new() };
        let (target, range) = match d.expr() {
            Expr::Ref(c) => (sheet, sheets_engine::Range::single(*c)),
            Expr::Range(r) => (sheet, *r),
            Expr::SheetRef(s, c) => match self.book.sheet_index(s) {
                Some(i) => (i, sheets_engine::Range::single(*c)),
                None => return Vec::new(),
            },
            Expr::SheetRange(s, r) => match self.book.sheet_index(s) {
                Some(i) => (i, *r),
                None => return Vec::new(),
            },
            _ => return Vec::new(),
        };
        vec![target as u32, range.start.row, range.start.col, range.end.row, range.end.col]
    }
}
