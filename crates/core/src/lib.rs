//! Shared foundation for every Office app: ids, operations, undo/redo.
//! Every edit in Docs/Sheets/Slides is expressed as an `Op` so that undo,
//! collaboration (CRDT) and file autosave all work the same way.

pub mod history;

use std::sync::atomic::{AtomicU64, Ordering};

/// Unique id for any node in a document (paragraph, cell, shape...).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct NodeId(pub u64);

static NEXT_ID: AtomicU64 = AtomicU64::new(1);

impl NodeId {
    pub fn new() -> Self {
        NodeId(NEXT_ID.fetch_add(1, Ordering::Relaxed))
    }
}

impl Default for NodeId {
    fn default() -> Self {
        Self::new()
    }
}

/// An edit that can be applied and reverted.
pub trait Op<Doc> {
    fn apply(&self, doc: &mut Doc);
    fn invert(&self, doc: &Doc) -> Box<dyn Op<Doc>>;
}

#[derive(Debug)]
pub enum OfficeError {
    Parse(String),
    InvalidRef(String),
    Io(String),
}

impl std::fmt::Display for OfficeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            OfficeError::Parse(m) => write!(f, "parse error: {m}"),
            OfficeError::InvalidRef(m) => write!(f, "invalid reference: {m}"),
            OfficeError::Io(m) => write!(f, "io error: {m}"),
        }
    }
}

impl std::error::Error for OfficeError {}

pub type Result<T> = std::result::Result<T, OfficeError>;
