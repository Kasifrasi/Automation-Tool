//! Excel writer for the financial report
//!
//! - `engine`: main write functions ([`write_report_with_options`], `create_protected_report`)
//! - `layout`: sheet setup (column widths, row heights, freeze panes)
//! - `structure`: report structure (merges, blanks, static strings, validations)

mod engine;
mod layout;
mod structure;

pub use layout::{setup_freeze_panes, setup_sheet, MergeRange};
pub use engine::{apply_sheet_options, write_report_with_options, BodyResult};
pub(crate) use engine::{create_protected_report, create_protected_report_precomputed};
