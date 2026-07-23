//! Excel-owned text interchange, data transformation, what-if, and link APIs.
//!
//! This module deliberately delegates parsing, calculation, and link handling
//! to Excel.  Its public wrappers are apartment-bound like the rest of the
//! crate and do not expose COM pointers.

mod advanced_filter;
mod helpers;
mod links;
mod scenarios;
mod text_export;
mod text_import;
mod transformation;
mod types;
mod what_if;

#[allow(unused_imports)]
pub use links::AskToUpdateLinksGuard;
#[allow(unused_imports)]
pub use scenarios::{Scenario, Scenarios, ScenariosIter};
pub use types::*;
