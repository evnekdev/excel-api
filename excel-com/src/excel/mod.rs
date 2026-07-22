mod application;
mod range;
mod workbook;
mod workbooks;
mod worksheet;
mod worksheets;

pub use application::Application;
pub use range::Range;
pub use workbook::Workbook;
pub use workbooks::Workbooks;
pub use worksheet::{Worksheet, XlSheetVisibility};
pub use worksheets::{Worksheets, WorksheetsAddOptions};

use crate::internal::{ComPtr, Dispatch};
use std::fmt::{Debug, Formatter};

pub(crate) struct DispatchObject {
    pub(crate) dispatch: ComPtr<Dispatch>,
    pub(crate) kind: &'static str,
}
impl Clone for DispatchObject {
    fn clone(&self) -> Self {
        Self {
            dispatch: self.dispatch.clone(),
            kind: self.kind,
        }
    }
}
impl Debug for DispatchObject {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("DispatchObject")
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}
