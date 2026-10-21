use crate::page_setup::PageSetup;
use crate::sheet::Sheet;

#[derive(Debug, Clone)]
pub struct Worksheet {
    pub name: String,
    pub sheet: Sheet,
    /// How the sheet prints (Page Layout settings).
    pub page_setup: PageSetup,
}

impl Worksheet {
    pub fn new(name: String, sheet: Sheet) -> Self {
        Worksheet { name, sheet, page_setup: PageSetup::default() }
    }
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
        Workbook { sheets: vec![Worksheet::new("Sheet1".into(), Sheet::new())] }
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
        self.sheets.push(Worksheet::new(name, sheet));
        self.sheets.len() - 1
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn names_new_sheets() {
        let mut wb = Workbook::new();
        assert_eq!(wb.add_sheet(), 1);
        assert_eq!(wb.sheets[1].name, "Sheet2");
    }
}
