//! IronCalc-based formula engine
//!
//! Single source of truth for all formulas — IronCalc evaluates them,
//! the writer copies formulas + cached values into the Excel output.
//!
//! - [`ModelTemplate`]: master model with static formulas + language sheet (created once)
//! - [`CalcBridge`]: per-report instance — dynamic formulas, input population, evaluation

mod bridge;
mod template;

pub(crate) use bridge::CalcBridge;
pub(crate) use template::ModelTemplate;
