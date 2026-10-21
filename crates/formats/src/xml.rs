//! A small, fast XML reader for Office Open XML parts.
//!
//! Worksheets can hold millions of `<c>` elements, so this reader never
//! allocates per element: names, attribute values and text are slices of the
//! input, unescaped only when they contain `&`. It understands what OOXML
//! producers write — elements, attributes (single or double quotes), text,
//! CDATA, comments, processing instructions and a DOCTYPE — and ignores
//! namespace prefixes (`x:c` is `c`, `r:id` is `id`). It does not validate.

use std::borrow::Cow;

/// One event while walking a document.
pub enum XmlEvent<'a> {
    /// An opening (or self-closing) tag. A self-closing tag is followed by its `End`.
    Start {
        name: &'a str,
        attrs: Attrs<'a>,
    },
    End {
        name: &'a str,
    },
    /// Text content, already unescaped.
    Text(&'a str),
}

/// The raw attribute text of a tag, read lazily with [`attr`].
#[derive(Clone, Copy)]
pub struct Attrs<'a>(&'a str);

impl<'a> Attrs<'a> {
    /// The value of attribute `key` (compared without its namespace prefix).
    pub fn get(&self, key: &str) -> Option<Cow<'a, str>> {
        let s = self.0;
        let b = s.as_bytes();
        let mut i = 0;
        while i < b.len() {
            while i < b.len() && b[i].is_ascii_whitespace() {
                i += 1;
            }
            let name_start = i;
            while i < b.len() && b[i] != b'=' && !b[i].is_ascii_whitespace() {
                i += 1;
            }
            let name = &s[name_start..i];
            while i < b.len() && (b[i].is_ascii_whitespace() || b[i] == b'=') {
                i += 1;
            }
            if i >= b.len() {
                return None;
            }
            let quote = b[i];
            if quote != b'"' && quote != b'\'' {
                return None;
            }
            i += 1;
            let value_start = i;
            while i < b.len() && b[i] != quote {
                i += 1;
            }
            let value = &s[value_start..i.min(b.len())];
            i += 1;
            if local(name) == key {
                return Some(unescape(value));
            }
        }
        None
    }
}

/// The attribute `key`, or "" when missing.
pub fn attr<'a>(attrs: &Attrs<'a>, key: &str) -> Cow<'a, str> {
    attrs.get(key).unwrap_or(Cow::Borrowed(""))
}

fn local(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, l)| l)
}

/// Replace the five predefined entities and numeric character references.
pub fn unescape(s: &str) -> Cow<'_, str> {
    if !s.contains('&') {
        return Cow::Borrowed(s);
    }
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(amp) = rest.find('&') {
        out.push_str(&rest[..amp]);
        rest = &rest[amp..];
        let Some(semi) = rest[..rest.len().min(12)].find(';') else {
            out.push('&');
            rest = &rest[1..];
            continue;
        };
        let entity = &rest[1..semi];
        let ch = match entity {
            "amp" => Some('&'),
            "lt" => Some('<'),
            "gt" => Some('>'),
            "quot" => Some('"'),
            "apos" => Some('\''),
            _ if entity.starts_with("#x") || entity.starts_with("#X") => {
                u32::from_str_radix(&entity[2..], 16).ok().and_then(char::from_u32)
            }
            _ if entity.starts_with('#') => entity[1..].parse().ok().and_then(char::from_u32),
            _ => None,
        };
        match ch {
            Some(c) => {
                out.push(c);
                rest = &rest[semi + 1..];
            }
            None => {
                out.push('&');
                rest = &rest[1..];
            }
        }
    }
    out.push_str(rest);
    Cow::Owned(out)
}

/// Call `on` for every element start/end and every piece of text.
/// Malformed input ends the walk early; what was read so far is kept.
pub fn walk(xml: &str, mut on: impl FnMut(&XmlEvent)) {
    let b = xml.as_bytes();
    let mut i = 0;
    while i < b.len() {
        if b[i] != b'<' {
            let end = xml[i..].find('<').map_or(b.len(), |j| i + j);
            let text = unescape(&xml[i..end]);
            on(&XmlEvent::Text(&text));
            i = end;
            continue;
        }
        let rest = &xml[i..];
        if rest.starts_with("<?") {
            let Some(end) = rest.find("?>") else { return };
            i += end + 2;
        } else if rest.starts_with("<!--") {
            let Some(end) = rest.find("-->") else { return };
            i += end + 3;
        } else if let Some(cdata) = rest.strip_prefix("<![CDATA[") {
            let Some(end) = cdata.find("]]>") else { return };
            on(&XmlEvent::Text(&cdata[..end]));
            i += 9 + end + 3;
        } else if rest.starts_with("<!") {
            let Some(end) = rest.find('>') else { return };
            i += end + 1;
        } else if let Some(close) = rest.strip_prefix("</") {
            let Some(end) = close.find('>') else { return };
            on(&XmlEvent::End { name: local(close[..end].trim()) });
            i += 2 + end + 1;
        } else {
            // Find the closing '>' outside quoted attribute values.
            let mut j = i + 1;
            let mut quote = 0u8;
            while j < b.len() {
                let c = b[j];
                if quote != 0 {
                    if c == quote {
                        quote = 0;
                    }
                } else if c == b'"' || c == b'\'' {
                    quote = c;
                } else if c == b'>' {
                    break;
                }
                j += 1;
            }
            if j >= b.len() {
                return;
            }
            let inner = &xml[i + 1..j];
            let (inner, empty) = match inner.strip_suffix('/') {
                Some(x) => (x, true),
                None => (inner, false),
            };
            let name_end = inner.find(|c: char| c.is_ascii_whitespace()).unwrap_or(inner.len());
            let name = local(&inner[..name_end]);
            on(&XmlEvent::Start { name, attrs: Attrs(&inner[name_end..]) });
            if empty {
                on(&XmlEvent::End { name });
            }
            i = j + 1;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn events(xml: &str) -> Vec<String> {
        let mut out = Vec::new();
        walk(xml, |ev| {
            out.push(match ev {
                XmlEvent::Start { name, attrs } => format!("<{name} r={} t={}>", attr(attrs, "r"), attr(attrs, "id")),
                XmlEvent::End { name } => format!("</{name}>"),
                XmlEvent::Text(t) => format!("'{t}'"),
            })
        });
        out
    }

    #[test]
    fn reads_elements_attributes_and_text() {
        let xml =
            r#"<?xml version="1.0"?><!-- c --><x:c r="A1" r:id='rId&amp;1'><v>1 &lt; 2</v><e/><![CDATA[a<b]]></x:c>"#;
        assert_eq!(
            events(xml),
            ["<c r=A1 t=rId&1>", "<v r= t=>", "'1 < 2'", "</v>", "<e r= t=>", "</e>", "'a<b'", "</c>"]
        );
    }

    #[test]
    fn handles_gt_inside_attributes_and_char_refs() {
        let xml = r#"<f a="x>y" b='&#65;&#x42;'/>"#;
        let mut got = Vec::new();
        walk(xml, |ev| {
            if let XmlEvent::Start { attrs, .. } = ev {
                got.push((attr(attrs, "a").into_owned(), attr(attrs, "b").into_owned()));
            }
        });
        assert_eq!(got, [("x>y".to_string(), "AB".to_string())]);
        assert_eq!(unescape("a & b &bogus; &amp;"), "a & b &bogus; &");
    }
}
