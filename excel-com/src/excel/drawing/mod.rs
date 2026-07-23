//! Stable public facade for Excel charts, shapes, and sparklines.
//!
//! The retained implementation is isolated in `legacy` while the public
//! names remain re-exported unchanged. New drawing work belongs in focused
//! modules rather than extending this facade.

mod legacy;

pub use legacy::*;
