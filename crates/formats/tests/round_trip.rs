//! Every feature that `.xlsx` can hold survives a write and a read: values and
//! formulas, cross-sheet formulas, formatting (fonts, borders, alignment, wrap,
//! number formats), column widths, row heights, hidden lines, merges, frozen
//! panes, AutoFilter, data validation, conditional formatting, charts,
//! pictures and page setup.

use sheets_engine::borders::{BorderLine, BorderStyle};
use sheets_engine::chart::{self, Anchor, AnchorPoint, ChartType, DrawingContent, Image};
use sheets_engine::conditional_format::rule_from_spec;
use sheets_engine::filter::{AutoFilter, ColumnFilter};
use sheets_engine::layout::Freeze;
use sheets_engine::validation::{Validation, ValidationKind, ValidationOp};
use sheets_engine::{Align, CellRef, NumberFormat, Range, Style, VAlign, Value, Workbook};

fn at(s: &str) -> CellRef {
    CellRef::parse(s).unwrap()
}

fn range(s: &str) -> Range {
    let (a, b) = s.split_once(':').unwrap();
    Range::new(at(a), at(b))
}

fn build() -> Workbook {
    let mut book = Workbook::new();
    book.add_sheet();
    book.sheets[1].name = "Data 2025".into();
    let rows = [
        ["Month", "Sales", "Cost"],
        ["Jan", "120", "80"],
        ["Feb", "150", "90"],
        ["Mar", "90", "70"],
        ["Apr", "200", "110"],
    ];
    let mut edits = Vec::new();
    for (r, row) in rows.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            edits.push((CellRef::new(r as u32, c as u32), v.to_string()));
        }
    }
    edits.push((at("D2"), "=B2-C2".into()));
    edits.push((at("E1"), "='Data 2025'!A1*2".into()));
    book.set_inputs(0, edits);
    book.set_inputs(1, vec![(at("A1"), "21".into())]);

    let s = &mut book.sheets[0].sheet;
    s.set_style(
        at("A1"),
        Style {
            bold: true,
            italic: true,
            color: Some("#C00000".into()),
            fill: Some("#FFF2CC".into()),
            font: Some("Arial".into()),
            font_size: Some(140),
            align: Some(Align::Center),
            valign: VAlign::Center,
            wrap: true,
            ..Style::default()
        },
    );
    let mut bordered = Style { number_format: NumberFormat::Percent { decimals: 1 }, ..Style::default() };
    bordered.borders.bottom = Some(BorderLine::new(BorderStyle::Double, Some("#FF0000".into())));
    bordered.borders.left = Some(BorderLine::new(BorderStyle::Thin, None));
    s.set_style(at("B2"), bordered);
    s.set_col_width(0, Some(120.0));

    let layout = s.layout_mut();
    layout.set_row_height(6, 6, Some(40.0));
    layout.hidden_cols.set(7, 7, Some(()));
    layout.set_merges([range("F1:G2")]);
    layout.freeze = Some(Freeze { rows: 1, cols: 1 });

    let mut filter = AutoFilter::new(range("A1:C5"));
    filter
        .columns
        .insert(0, ColumnFilter::Values { values: ["Jan".to_string(), "Mar".to_string()].into(), blanks: false });
    let mut tools = s.with_filter_applied(filter);
    tools.validations = vec![Validation {
        ranges: vec![range("C2:C5")],
        kind: ValidationKind::Whole,
        operator: ValidationOp::Between,
        formula1: "0".into(),
        formula2: "500".into(),
        allow_blank: true,
        in_cell_dropdown: true,
        show_error: true,
        error_title: "Cost".into(),
        error_message: "Between 0 and 500".into(),
        ..Validation::default()
    }];
    tools.cond_formats = vec![
        rule_from_spec(vec![range("B2:B5")], "cellIs", &["greaterThan".into(), "100".into()], "lightRedDarkRed")
            .unwrap(),
        rule_from_spec(vec![range("C2:C5")], "dataBar", &["#638EC6".into()], "").unwrap(),
    ];
    s.set_data_tools(tools);

    let mut drawings = s.drawings().clone();
    let anchor = Anchor { from: AnchorPoint::new(7, 0, 0.0, 0.0), to: AnchorPoint::new(20, 6, 0.0, 0.0) };
    let c = chart::chart_for_range(s, range("A1:C5"), ChartType::Line).unwrap();
    drawings.push("Chart 1".into(), anchor, DrawingContent::Chart(Box::new(c)));
    // A 1x1 transparent PNG.
    let png: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 0x0D, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1, 0, 0, 0, 1,
        8, 6, 0, 0, 0, 0x1F, 0x15, 0xC4, 0x89, 0, 0, 0, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0, 1, 0, 0, 5,
        0, 1, 0x0D, 0x0A, 0x2D, 0xB4, 0, 0, 0, 0, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];
    let pic = Anchor { from: AnchorPoint::new(22, 0, 0.0, 0.0), to: AnchorPoint::new(25, 2, 0.0, 0.0) };
    drawings.push(
        "Picture 1".into(),
        pic,
        DrawingContent::Picture(Image { mime: "image/png".into(), data: png.into() }),
    );
    s.set_drawings(drawings);

    let p = &mut book.sheets[0].page_setup;
    p.set("orientation", "landscape").unwrap();
    p.set("fit", "1,0").unwrap();
    p.set("printArea", "A1:G20").unwrap();
    p.set("titleRows", "1").unwrap();
    p.set("gridlines", "1").unwrap();
    p.set("footer", "&CPage &P of &N").unwrap();
    book
}

#[test]
fn everything_round_trips() {
    let book = build();
    let back = office_formats::xlsx::read(&office_formats::xlsx::write(&book).unwrap()).unwrap();
    let (a, b) = (&book.sheets[0], &back.sheets[0]);
    assert_eq!(back.sheets[1].name, "Data 2025");

    // Values and formulas, including the one reading another sheet.
    for cell in ["A1", "B2", "D2", "E1"] {
        assert_eq!(a.sheet.input(at(cell)), b.sheet.input(at(cell)), "{cell}");
        assert_eq!(a.sheet.value(at(cell)), b.sheet.value(at(cell)), "{cell}");
    }
    assert_eq!(b.sheet.value(at("E1")), &Value::Number(42.0));

    // Formatting and sizes.
    assert_eq!(a.sheet.style(at("A1")), b.sheet.style(at("A1")));
    assert_eq!(a.sheet.style(at("B2")), b.sheet.style(at("B2")));
    assert_eq!(b.sheet.col_width(0), 120.0);

    // Layout.
    assert_eq!(b.sheet.layout().row_height(6), 40.0);
    assert!(b.sheet.is_col_hidden(7));
    assert_eq!(b.sheet.layout().merges(), &[range("F1:G2")]);
    assert_eq!(b.sheet.layout().freeze, Some(Freeze { rows: 1, cols: 1 }));

    // Data tools: the filter keeps hiding Feb and Apr.
    assert_eq!(a.sheet.data_tools(), b.sheet.data_tools());
    assert!(b.sheet.is_row_hidden(2) && b.sheet.is_row_hidden(4) && !b.sheet.is_row_hidden(1));

    // Charts and pictures.
    let names: Vec<&str> = b.sheet.drawings().items.iter().map(|d| d.name.as_str()).collect();
    assert_eq!(names, ["Chart 1", "Picture 1"]);
    assert_eq!(a.sheet.drawings().items[0].chart(), b.sheet.drawings().items[0].chart());
    assert_eq!(a.sheet.drawings().items[1].anchor, b.sheet.drawings().items[1].anchor);

    // Page setup.
    assert_eq!(a.page_setup, b.page_setup);
}
