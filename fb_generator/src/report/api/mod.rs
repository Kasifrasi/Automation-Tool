//! Public API types for the financial report
//!
//! - `keys`: [`ApiKey`] and [`PositionField`] for cell addressing
//! - `types`: Type-safe types ([`Language`], [`Currency`], [`Category`], [`ReportDate`])
//! - `values`: [`ReportValues`](CellValue) for input data

mod keys;
mod types;
mod values;

pub use keys::{ApiKey, FooterField, PositionField};
pub use types::{Category, Currency, DateError, Language, ReportDate};
pub use values::{CellValue, ReportValues};
