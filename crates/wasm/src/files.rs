//! Bindings for files and printing: page setup and CSV options.

use sheets_engine::page_setup::SetPageSetup;
use sheets_engine::Workbook;
use wasm_bindgen::prelude::*;

use crate::SheetsDocument;

/// "," / ";" / "\t" / "|" -> char; anything else (e.g. "auto") -> None.
fn delimiter(d: &str) -> Option<char> {
    let mut chars = d.chars();
    match (chars.next(), chars.next()) {
        (Some(c), None) => Some(c),
        _ => None,
    }
}

#[wasm_bindgen]
impl SheetsDocument {
    /// Page setup of a sheet as JSON (see `PageSetup` in `@office/engine`).
    #[wasm_bindgen(js_name = pageSetup)]
    pub fn page_setup(&self, sheet: usize) -> String {
        self.book.sheets[sheet].page_setup.to_json()
    }

    /// Change page setup settings as one undo step. `keys` and `values` are
    /// parallel; see `PageSetup::set` for the keys. Nothing changes on error.
    #[wasm_bindgen(js_name = setPageSetup)]
    pub fn set_page_setup(&mut self, sheet: usize, keys: Vec<String>, values: Vec<String>) -> Result<(), JsError> {
        let ws = self.book.sheets.get(sheet).ok_or_else(|| JsError::new("no such sheet"))?;
        let mut setup = ws.page_setup.clone();
        for (k, v) in keys.iter().zip(&values) {
            setup.set(k, v).map_err(|e| JsError::new(&e.to_string()))?;
        }
        if setup != ws.page_setup {
            self.apply(SetPageSetup { sheet, setup });
        }
        Ok(())
    }

    /// Delimited text of a sheet's displayed values; `delimiter` is one character.
    #[wasm_bindgen(js_name = toCsvWith)]
    pub fn to_csv_with(&self, sheet: usize, delimiter_char: &str) -> String {
        office_formats::csv::write_with(self.sheet(sheet), delimiter(delimiter_char).unwrap_or(','))
    }

    /// Open delimited text. `delimiter_char` is one character, or "auto" to detect it.
    #[wasm_bindgen(js_name = fromCsvWith)]
    pub fn from_csv_with(text: &str, delimiter_char: &str, sheet_name: &str) -> SheetsDocument {
        let d = delimiter(delimiter_char).unwrap_or_else(|| office_formats::csv::detect_delimiter(text));
        let mut book = Workbook::empty();
        let name = sheet_name.trim();
        book.push_sheet(
            if name.is_empty() { "Sheet1".into() } else { name.chars().take(31).collect() },
            office_formats::csv::read_with(text, d),
        );
        Self::from_book(book)
    }

    /// The delimiter `fromCsvWith(text, "auto")` would use.
    #[wasm_bindgen(js_name = detectCsvDelimiter)]
    pub fn detect_csv_delimiter(text: &str) -> String {
        office_formats::csv::detect_delimiter(text).to_string()
    }

    /// The first rows of delimited text split into fields, for an import preview.
    /// Each row's fields are joined with the unit separator U+001F.
    #[wasm_bindgen(js_name = previewCsv)]
    pub fn preview_csv(text: &str, delimiter_char: &str, max_rows: usize) -> Vec<String> {
        let d = delimiter(delimiter_char).unwrap_or_else(|| office_formats::csv::detect_delimiter(text));
        // Only parse the start of big files.
        let head: String = text.chars().take(256 * 1024).collect();
        office_formats::csv::parse_with(&head, d).into_iter().take(max_rows).map(|row| row.join("\u{1f}")).collect()
    }
}
