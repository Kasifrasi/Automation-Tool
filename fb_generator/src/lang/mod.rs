//! Language sheet generation.
//!
//! - [`data`]: [`TEXT_MATRIX`] — 66 UI terms x 5 languages + [`Currency`](crate::Currency) enum
//! - [`config`]: [`LanguageConfig`] + [`LANG_CONFIG`] — per-language sheet names and messages
//! - [`builder`]: [`build_sheet`] / [`build_sheet_with_visibility`] — writes the "Sprachversionen" sheet

pub mod builder;
pub mod config;
pub mod data;

pub use builder::{build_sheet, build_sheet_with_visibility};
pub use config::{LanguageConfig, LANG_CONFIG};
pub use data::TEXT_MATRIX;
