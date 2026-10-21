//! Presentation engine: slides made of positioned shapes.

use office_core::NodeId;

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

#[derive(Debug, Clone)]
pub enum ShapeKind {
    TextBox(String),
    Rectangle,
    Ellipse,
    Image { src: String },
}

#[derive(Debug, Clone)]
pub struct Shape {
    pub id: NodeId,
    pub frame: Rect,
    pub kind: ShapeKind,
}

#[derive(Debug, Default)]
pub struct Slide {
    pub id: NodeId,
    pub shapes: Vec<Shape>,
}

#[derive(Debug)]
pub struct Presentation {
    /// Slide size in points; default 16:9 (960 x 540).
    pub width: f32,
    pub height: f32,
    pub slides: Vec<Slide>,
}

impl Default for Presentation {
    fn default() -> Self {
        Self { width: 960.0, height: 540.0, slides: vec![Slide::default()] }
    }
}
