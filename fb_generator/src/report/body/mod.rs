//! Dynamic body area (cost categories) — domain logic
//!
//! - [`BodyConfig`]: number of position rows per category (1–8)
//! - [`BodyLayout`]: computed row positions for all categories
//!
//! Excel writing logic lives in `writer/sections/body.rs` and `writer/structure.rs`.

pub(crate) mod config;
mod layout;

pub use config::BodyConfig;
pub use layout::{
    BodyLayout, CategoryLayout, CategoryMode, FooterLayout, PositionRange, TOTAL_LABEL_INDEX,
};
