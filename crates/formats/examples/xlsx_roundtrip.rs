//! Debug helper: print every cell of an .xlsx file, optionally re-save it.
//!
//! cargo run -p office-formats --example xlsx_roundtrip -- input.xlsx [output.xlsx]

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let Some(input) = args.get(1) else {
        eprintln!("usage: xlsx_roundtrip <input.xlsx> [output.xlsx]");
        std::process::exit(2);
    };
    let bytes = std::fs::read(input).expect("read input");
    let book = office_formats::xlsx::read(&bytes).expect("parse xlsx");
    for ws in &book.sheets {
        println!("== {}", ws.name);
        for (at, input, value) in ws.sheet.cells() {
            println!("{at}\t{input}\t{}", value.display());
        }
    }
    if let Some(output) = args.get(2) {
        std::fs::write(output, office_formats::xlsx::write(&book).expect("write xlsx")).expect("write output");
    }
}
