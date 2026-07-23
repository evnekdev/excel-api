//! PivotCaches, PivotTables, typed fields, filters, and slicer inspection.
//!
//! PivotCaches own source data while PivotTables own report layout. All
//! wrappers remain apartment-bound; provider, Excel-version, and host runtime
//! restrictions are returned as ordinary Excel errors rather than hidden.
//!
//! A PivotCache owns source data while a PivotTable owns report layout:
//!
//! ```no_run
//! # use excel_com::{ExcelComError, ListObject, PivotTableCreateOptions, Workbook, Worksheet};
//! # fn example(workbook: &Workbook, worksheet: &Worksheet, table: &ListObject) -> Result<(), ExcelComError> {
//! let cache = workbook.pivot_caches()?.create_from_table(table, None)?;
//! let pivot = cache.create_pivot_table(&PivotTableCreateOptions {
//!     destination: &worksheet.range("H2")?,
//!     name: "SalesPivot",
//!     version: None,
//!     read_data: None,
//!     default_version: None,
//! })?;
//! # let _ = pivot;
//! # Ok(())
//! # }
//! ```
//!
//! Field orientations are typed and layout temporarily uses `ManualUpdate`.
//! Excel validates incompatible orientations and its own visibility rules:
//!
//! ```no_run
//! # use excel_com::{AggregationFunction, ExcelComError, PivotDataField, PivotFieldOrientation, PivotFieldPlacement, PivotLayoutOptions, PivotTable};
//! # fn example(pivot: &PivotTable) -> Result<(), ExcelComError> {
//! pivot.apply_layout(&PivotLayoutOptions {
//!     fields: vec![PivotFieldPlacement {
//!         field_name: "Region",
//!         orientation: PivotFieldOrientation::ROW,
//!         position: Some(1),
//!     }],
//!     data_fields: vec![PivotDataField {
//!         field_name: "Amount",
//!         caption: Some("Total Amount"),
//!         function: AggregationFunction::SUM,
//!         number_format: Some("#,##0.00"),
//!     }],
//! })?;
//! # Ok(())
//! # }
//! ```

mod cache;
mod caches;
mod fields;
mod filters;
mod helpers;
mod items;
mod slicers;
mod table;
mod tables;
mod types;

pub use cache::PivotCache;
pub use caches::{PivotCaches, PivotCachesIter};
pub use fields::{PivotField, PivotFields, PivotFieldsIter};
pub use filters::{PivotFilter, PivotFilters, PivotFiltersIter};
pub use items::{PivotItem, PivotItems, PivotItemsIter};
pub use slicers::{Slicer, SlicerCache, SlicerCaches, SlicerCachesIter, Slicers, SlicersIter};
pub use table::PivotTable;
pub use tables::{PivotTables, PivotTablesIter};
pub use types::*;

#[cfg(test)]
mod tests;
