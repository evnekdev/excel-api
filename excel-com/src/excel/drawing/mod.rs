//! Stable public facade for Excel charts, shapes, and sparklines.
//!
//! The former monolithic implementation is split by responsibility. Public
//! names are re-exported unchanged so existing users retain source compatibility.

mod types;
pub use types::*;
mod chart_objects;
mod helpers;
pub use chart_objects::*;
mod chart;
pub use chart::*;
mod chart_groups;
pub use chart_groups::*;
mod export;
mod labels;
pub use labels::*;
mod office_format;
pub use office_format::*;
mod series;
pub use series::*;
mod points;
pub use points::*;
mod trendlines;
pub use trendlines::*;
mod axes;
pub use axes::*;
mod chart_sheets;
pub use chart_sheets::*;
mod shapes;
pub use shapes::*;
mod sparklines;
pub use sparklines::*;

#[cfg(test)]
mod tests;
