//! Notes and hyperlinks in .xlsx.
//!
//! Notes are `xl/commentsN.xml` (the text and author) plus a VML drawing
//! (`xl/drawings/vmlDrawingN.vml`) that Excel needs to show the note box; the
//! worksheet points at the VML with `<legacyDrawing>`. Hyperlinks are
//! `<hyperlinks>` in the worksheet: web links through a relationship with
//! `TargetMode="External"`, places in the workbook through `location`.

use std::io::Cursor;

use office_core::Result;
use sheets_engine::annotations::{Annotations, Hyperlink, Note};
use sheets_engine::{CellRef, Range, Sheet};
use zip::ZipArchive;

use super::drawings::{insert_sheet_element, part_text, rels_path, resolve, DrawingWriter};
use super::{attr, escape, relationships, walk, XmlEvent, NS_MAIN, NS_REL};

const CT_COMMENTS: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";

/// Cells a hyperlink over a range is copied to, at most.
const MAX_LINK_CELLS: u32 = 10_000;

// ---------------------------------------------------------------- writing

/// Write sheet `id`'s notes and links into `parts` and its worksheet XML.
pub(super) fn write_sheet(id: usize, sheet: &Sheet, worksheet_xml: String, parts: &mut DrawingWriter) -> String {
    let a = sheet.annotations();
    let mut xml = worksheet_xml;
    let mut rels = String::new();
    if !a.links.is_empty() {
        let mut links = String::from("<hyperlinks>");
        for (n, (at, link)) in a.links.iter().enumerate() {
            let tip = if link.tooltip.is_empty() {
                String::new()
            } else {
                format!(r#" tooltip="{}""#, escape(&link.tooltip))
            };
            match link.location() {
                Some(place) => {
                    links.push_str(&format!(r#"<hyperlink ref="{}" location="{}"{tip}/>"#, at, escape(place)))
                }
                None => {
                    let rid = format!("rIdLink{}", n + 1);
                    links.push_str(&format!(r#"<hyperlink ref="{at}" r:id="{rid}"{tip}/>"#));
                    rels.push_str(&format!(
                        r#"<Relationship Id="{rid}" Type="{NS_REL}/hyperlink" Target="{}" TargetMode="External"/>"#,
                        escape(&link.target)
                    ));
                }
            }
        }
        links.push_str("</hyperlinks>");
        xml = insert_sheet_element(
            xml,
            &links,
            &[
                "<printOptions",
                "<pageMargins",
                "<pageSetup",
                "<headerFooter",
                "<rowBreaks",
                "<colBreaks",
                "<customProperties",
                "<cellWatches",
                "<ignoredErrors",
                "<smartTags",
                "<drawing",
                "<legacyDrawing",
                "<picture",
                "<oleObjects",
                "<controls",
                "<webPublishItems",
                "<tableParts",
                "<extLst",
                "</worksheet>",
            ],
        );
    }
    if !a.notes.is_empty() {
        parts.add_part(format!("xl/comments{id}.xml"), comments_xml(a), Some(CT_COMMENTS));
        parts.add_part(format!("xl/drawings/vmlDrawing{id}.vml"), vml_xml(id, a), None);
        parts.add_extension("vml");
        rels.push_str(&format!(
            r#"<Relationship Id="rIdComments1" Type="{NS_REL}/comments" Target="../comments{id}.xml"/>"#
        ));
        rels.push_str(&format!(
            r#"<Relationship Id="rIdVml1" Type="{NS_REL}/vmlDrawing" Target="../drawings/vmlDrawing{id}.vml"/>"#
        ));
        xml = insert_sheet_element(
            xml,
            r#"<legacyDrawing r:id="rIdVml1"/>"#,
            &[
                "<legacyDrawingHF",
                "<picture",
                "<oleObjects",
                "<controls",
                "<webPublishItems",
                "<tableParts",
                "<extLst",
                "</worksheet>",
            ],
        );
    }
    if !rels.is_empty() {
        parts.add_part(
            format!("xl/worksheets/_rels/sheet{id}.xml.rels"),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{}">{rels}</Relationships>"#,
                super::NS_PKG_REL
            ),
            None,
        );
    }
    xml
}

/// Like Excel, the author's name starts the text in bold.
fn comments_xml(a: &Annotations) -> String {
    let mut authors: Vec<&str> = Vec::new();
    for note in a.notes.values() {
        if !authors.contains(&note.author.as_str()) {
            authors.push(&note.author);
        }
    }
    let mut xml =
        format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="{NS_MAIN}"><authors>"#);
    for author in &authors {
        xml.push_str(&format!("<author>{}</author>", escape(author)));
    }
    xml.push_str("</authors><commentList>");
    for (at, note) in &a.notes {
        let author_id = authors.iter().position(|x| *x == note.author).unwrap_or(0);
        xml.push_str(&format!(r#"<comment ref="{at}" authorId="{author_id}"><text>"#));
        if !note.author.is_empty() {
            xml.push_str(&format!(
                r#"<r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>{}:</t></r>"#,
                escape(&note.author)
            ));
        }
        let body = if note.author.is_empty() { note.text.clone() } else { format!("\n{}", note.text) };
        xml.push_str(&format!(
            r#"<r><rPr><sz val="9"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">{}</t></r></text></comment>"#,
            escape(&body)
        ));
    }
    xml.push_str("</commentList></comments>");
    xml
}

/// The note boxes (hidden until the cell is hovered), as Excel writes them.
fn vml_xml(id: usize, a: &Annotations) -> String {
    let mut xml = format!(
        concat!(
            r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"#,
            r#"<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="{id}"/></o:shapelayout>"#,
            r#"<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">"#,
            r#"<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>"#
        ),
        id = id
    );
    for (n, at) in a.notes.keys().enumerate() {
        let (row, col) = (at.row, at.col);
        xml.push_str(&format!(
            concat!(
                r##"<v:shape id="_x0000_s{shape}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:{z};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">"##,
                r##"<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>"##,
                r#"<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>"#,
                r#"<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>"#,
                r#"<x:Anchor>{c1}, 15, {r1}, 10, {c2}, 15, {r2}, 4</x:Anchor><x:AutoFill>False</x:AutoFill>"#,
                r#"<x:Row>{row}</x:Row><x:Column>{col}</x:Column></x:ClientData></v:shape>"#
            ),
            shape = id * 1024 + n + 1,
            z = n + 1,
            c1 = col + 1,
            r1 = row.saturating_sub(1),
            c2 = col + 3,
            r2 = row + 3,
            row = row,
            col = col
        ));
    }
    xml.push_str("</xml>");
    xml
}

// ---------------------------------------------------------------- reading

/// Read the notes (through the sheet's relationships) and the hyperlinks
/// (`tail`: the worksheet XML after `</sheetData>`).
pub(super) fn read_sheet(
    zip: &mut ZipArchive<Cursor<&[u8]>>,
    sheet_part: &str,
    tail: &str,
    sheet: &mut Sheet,
) -> Result<()> {
    let has_links = tail.contains("<hyperlink");
    let rels = match part_text(zip, &rels_path(sheet_part))? {
        Some(xml) => relationships(&xml),
        None if has_links => Vec::new(),
        None => return Ok(()),
    };
    let mut a = Annotations::default();
    for rel in rels.iter().filter(|r| r.kind.ends_with("/comments")) {
        if let Some(xml) = part_text(zip, &resolve(sheet_part, &rel.target))? {
            a.notes.extend(parse_comments(&xml));
        }
    }
    if has_links {
        walk(tail, |ev| {
            if let XmlEvent::Start { name: "hyperlink", attrs } = ev {
                let Some(range) = parse_range(&attr(attrs, "ref")) else { return };
                let rid = attr(attrs, "id");
                let target = match rels.iter().find(|r| !rid.is_empty() && r.id == *rid) {
                    Some(r) => {
                        let location = attr(attrs, "location");
                        if location.is_empty() {
                            r.target.clone()
                        } else {
                            format!("{}#{location}", r.target)
                        }
                    }
                    None => match &*attr(attrs, "location") {
                        "" => return,
                        place => format!("#{place}"),
                    },
                };
                let link = Hyperlink { target, tooltip: attr(attrs, "tooltip").to_string() };
                let mut count = 0;
                'cells: for row in range.start.row..=range.end.row {
                    for col in range.start.col..=range.end.col {
                        count += 1;
                        if count > MAX_LINK_CELLS {
                            break 'cells;
                        }
                        a.links.insert(CellRef::new(row, col), link.clone());
                    }
                }
            }
        });
    }
    if !a.is_empty() {
        *sheet.annotations_mut() = a;
    }
    Ok(())
}

/// `A1` or `A1:B5`.
fn parse_range(s: &str) -> Option<Range> {
    match s.split_once(':') {
        Some((a, b)) => Some(Range::new(CellRef::parse(a)?, CellRef::parse(b)?)),
        None => CellRef::parse(s).map(Range::single),
    }
}

fn parse_comments(xml: &str) -> Vec<(CellRef, Note)> {
    let mut authors: Vec<String> = Vec::new();
    let mut out = Vec::new();
    let (mut in_author, mut in_t, mut in_phonetic) = (false, false, false);
    let mut current: Option<(CellRef, usize)> = None;
    let mut text = String::new();
    walk(xml, |ev| match ev {
        XmlEvent::Start { name: "author", .. } => {
            in_author = true;
            authors.push(String::new());
        }
        XmlEvent::End { name: "author" } => in_author = false,
        XmlEvent::Text(t) if in_author => {
            if let Some(a) = authors.last_mut() {
                a.push_str(t);
            }
        }
        XmlEvent::Start { name: "comment", attrs } => {
            text.clear();
            current = CellRef::parse(&attr(attrs, "ref")).map(|at| (at, attr(attrs, "authorId").parse().unwrap_or(0)));
        }
        XmlEvent::End { name: "comment" } => {
            if let Some((at, author_id)) = current.take() {
                let author = authors.get(author_id).cloned().unwrap_or_default();
                // Excel starts the text with "Author:" and a line break.
                let body = text
                    .strip_prefix(&format!("{author}:"))
                    .filter(|_| !author.is_empty())
                    .map(|rest| rest.strip_prefix('\n').unwrap_or(rest))
                    .unwrap_or(&text)
                    .to_string();
                out.push((at, Note { author, text: body }));
            }
        }
        XmlEvent::Start { name: "rPh", .. } => in_phonetic = true,
        XmlEvent::End { name: "rPh" } => in_phonetic = false,
        XmlEvent::Start { name: "t", .. } => in_t = true,
        XmlEvent::End { name: "t" } => in_t = false,
        XmlEvent::Text(t) if in_t && !in_phonetic && current.is_some() => text.push_str(t),
        _ => {}
    });
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_excel_comments() {
        let xml = r#"<comments xmlns="x"><authors><author>Ravi Kumar</author></authors><commentList>
            <comment ref="B2" authorId="0"><text><r><rPr><b/></rPr><t>Ravi Kumar:</t></r><r><t xml:space="preserve">
Check this &amp; that</t></r></text></comment>
            <comment ref="C3" authorId="0"><text><t>No prefix</t></text></comment></commentList></comments>"#;
        let notes = parse_comments(xml);
        assert_eq!(
            notes[0],
            (CellRef::parse("B2").unwrap(), Note { author: "Ravi Kumar".into(), text: "Check this & that".into() })
        );
        assert_eq!(notes[1].1.text, "No prefix");
    }
}
