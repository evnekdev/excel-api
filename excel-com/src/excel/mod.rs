mod application;
mod areas;
mod border;
mod borders;
mod calculation;
mod collection;
mod file_lifecycle;
mod font;
mod formatting;
mod formula;
mod interior;
mod name;
mod names;
mod range;
mod reference;
mod search;
mod text;
mod workbook;
mod workbooks;
mod worksheet;
mod worksheets;

pub use application::{Application, CalculationModeGuard, DisplayAlertsGuard, ReferenceStyleGuard};
pub use areas::{Areas, AreasIter};
pub use border::Border;
pub use borders::{Borders, BordersIter};
pub use calculation::{CalculationMode, CalculationState};
pub use file_lifecycle::{
    SaveChanges, WorkbookCloseOptions, WorkbookOpenFormat, WorkbookOpenOptions,
    WorkbookSaveAsOptions, XlCorruptLoad, XlFileFormat, XlPlatform, XlSaveAsAccessMode,
    XlSaveConflictResolution, XlUpdateLinks,
};
pub use font::Font;
pub use formatting::{
    BorderIndex, BorderLineStyle, BorderWeight, ExcelColor, ExcelColorIndex, FillPattern,
    HorizontalAlignment, MixedValue, UnderlineStyle, VerticalAlignment,
};
pub use formula::FormulaValue;
pub use interior::Interior;
pub use name::Name;
pub use names::{NameAddOptions, NameRefersTo, Names, NamesIter};
pub use range::Range;
pub use reference::{
    FormulaConversionOptions, RangeAddressOptions, ReferenceAbsoluteMode, ReferenceStyle,
};
pub use search::{
    FindLookIn, FindMatchMode, FindOptions, RangeFindIter, ReplaceOptions, SearchDirection,
    SearchOrder, SpecialCellType, SpecialCellValueMask,
};
pub use workbook::Workbook;
pub use workbooks::{Workbooks, WorkbooksIter};
pub use worksheet::{Worksheet, XlSheetVisibility};
pub use worksheets::{Worksheets, WorksheetsAddOptions, WorksheetsIter};

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
impl DispatchObject {
    pub(crate) fn same_object(&self, other: &Self) -> Result<bool, crate::ExcelComError> {
        let left = self.dispatch.canonical_unknown()?;
        let right = other.dispatch.canonical_unknown()?;
        Ok(left.raw() == right.raw())
    }
}
