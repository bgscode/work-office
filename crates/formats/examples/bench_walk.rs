//! Micro-benchmark for the XML reader on a big worksheet part.
use std::io::Read;
use std::time::Instant;

fn main() {
    let bytes = std::fs::read("/tmp/bench_open.xlsx").expect("run bench_open first");
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap();
    println!("sheet xml: {:.1} MB", xml.len() as f64 / 1e6);
    let t = Instant::now();
    let mut n = 0usize;
    office_formats::xml::walk(&xml, |_| n += 1);
    println!("walk only: {:?} ({n} events)", t.elapsed());
    let t = Instant::now();
    let mut m = 0usize;
    office_formats::xml::walk(&xml, |ev| {
        if let office_formats::xml::XmlEvent::Start { name: "c", attrs } = ev {
            m += office_formats::xml::attr(attrs, "r").len() + office_formats::xml::attr(attrs, "t").len();
        }
    });
    println!("walk + 2 attrs per cell: {:?} ({m})", t.elapsed());
}
