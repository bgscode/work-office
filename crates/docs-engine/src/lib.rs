//! Word-processor engine: document model -> layout -> pages.

use office_core::NodeId;

#[derive(Debug, Clone, Default, PartialEq)]
pub struct RunStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub font_size_pt: Option<f32>,
    pub font_family: Option<String>,
}

#[derive(Debug, Clone)]
pub struct Run {
    pub text: String,
    pub style: RunStyle,
}

#[derive(Debug, Clone)]
pub struct Paragraph {
    pub id: NodeId,
    pub runs: Vec<Run>,
}

impl Paragraph {
    pub fn plain(text: &str) -> Self {
        Self { id: NodeId::new(), runs: vec![Run { text: text.into(), style: RunStyle::default() }] }
    }

    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }
}

#[derive(Debug, Default)]
pub struct Document {
    pub paragraphs: Vec<Paragraph>,
}

impl Document {
    pub fn push_paragraph(&mut self, text: &str) -> NodeId {
        let p = Paragraph::plain(text);
        let id = p.id;
        self.paragraphs.push(p);
        id
    }

    pub fn word_count(&self) -> usize {
        self.paragraphs.iter().map(|p| p.text().split_whitespace().count()).sum()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn counts_words() {
        let mut d = Document::default();
        d.push_paragraph("Hello office world");
        d.push_paragraph("second para");
        assert_eq!(d.word_count(), 5);
    }
}
