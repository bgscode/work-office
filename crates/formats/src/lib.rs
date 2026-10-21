//! File import/export.
//!
//! | Module | Format | Status |
//! |---|---|---|
//! | [`csv`] | CSV (RFC 4180) | read + write |
//! | [`xlsx`] | Excel workbook (OOXML SpreadsheetML) | read + write: values and formulas |
//! | `docx`, `pptx`, `pdf` | | planned |

pub mod csv;
pub mod xlsx;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Format {
    Docx,
    Xlsx,
    Pptx,
    Pdf,
    Csv,
}

impl Format {
    pub fn from_extension(ext: &str) -> Option<Self> {
        match ext.to_ascii_lowercase().as_str() {
            "docx" => Some(Format::Docx),
            "xlsx" => Some(Format::Xlsx),
            "pptx" => Some(Format::Pptx),
            "pdf" => Some(Format::Pdf),
            "csv" => Some(Format::Csv),
            _ => None,
        }
    }
}
