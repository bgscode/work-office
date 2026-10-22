//! Benchmark: open/save a big workbook (default 3 sheets × 100,000 rows × 10 columns).
//!
//! cargo run --release -p office-formats --example bench_open -- [rows] [sheets]

use std::time::Instant;

use sheets_engine::{CellRef, Workbook};

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let rows: u32 = args.get(1).and_then(|s| s.parse().ok()).unwrap_or(100_000);
    let sheets: usize = args.get(2).and_then(|s| s.parse().ok()).unwrap_or(3);

    let t = Instant::now();
    let mut book = Workbook::empty();
    for s in 0..sheets {
        let mut sheet = sheets_engine::Sheet::new();
        let mut edits = Vec::with_capacity(rows as usize * 10);
        for r in 0..rows {
            let n = r + 2;
            edits.push((CellRef::new(r, 0), format!("SKU-{:06}", r)));
            edits.push((CellRef::new(r, 1), format!("Item name {}", r % 977)));
            edits.push((CellRef::new(r, 2), format!("{}", (r * 7) % 500 + 1)));
            edits.push((CellRef::new(r, 3), format!("{}.{:02}", (r * 13) % 2000, r % 100)));
            edits.push((CellRef::new(r, 4), format!("=C{n}*D{n}")));
            edits.push((CellRef::new(r, 5), format!("{}", 45000 + r % 365)));
            edits.push((CellRef::new(r, 6), ["North", "South", "East", "West"][(r % 4) as usize].to_string()));
            edits.push((CellRef::new(r, 7), format!("{}", r % 2 == 0).to_uppercase()));
            edits.push((CellRef::new(r, 8), format!("{}", (r * 31) % 10_000)));
            edits.push((CellRef::new(r, 9), format!("Note {}", r % 101)));
        }
        sheet.set_inputs(edits);
        book.push_sheet(format!("Sheet{}", s + 1), sheet);
    }
    println!("build in memory: {:?}", t.elapsed());

    let t = Instant::now();
    let bytes = office_formats::xlsx::write(&book).unwrap();
    println!("xlsx write: {:?} ({:.1} MB)", t.elapsed(), bytes.len() as f64 / 1e6);
    std::fs::write("/tmp/bench_open.xlsx", &bytes).unwrap();

    let t = Instant::now();
    let back = office_formats::xlsx::read(&bytes).unwrap();
    println!("xlsx read (unzip + parse + load + recalc): {:?}", t.elapsed());
    assert_eq!(back.sheets.len(), sheets);
    println!("cells sheet1: {}", back.sheets[0].sheet.cells().len());
}
