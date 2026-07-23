mod application;
mod areas;
mod autofilter;
mod border;
mod borders;
mod calculation;
mod collection;
mod data;
mod drawing;
mod file_lifecycle;
mod font;
mod formatting;
mod formula;
mod interior;
mod name;
mod names;
mod presentation;
mod range;
mod range_structure;
mod reference;
mod search;
mod sort;
mod table;
mod text;
mod validation;
mod workbook;
mod workbooks;
mod worksheet;
mod worksheets;

pub use application::{Application, CalculationModeGuard, DisplayAlertsGuard, ReferenceStyleGuard};
pub use areas::{Areas, AreasIter};
pub use autofilter::{
    AutoFilter, AutoFilterOperator, AutoFilterOptions, DynamicFilterCriteria, Filter,
    FilterCriterion, Filters, FiltersIter,
};
pub use border::Border;
pub use borders::{Borders, BordersIter};
pub use calculation::{CalculationMode, CalculationState};
#[allow(unused_imports)]
pub use data::*;
pub use drawing::*;
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
#[allow(unused_imports)]
pub use presentation::{
    AboveAverage, AboveAverageOptions, AboveAverageRuleOptions, AboveBelowMode,
    CellValueRuleOptions, ColorScale, ColorScaleCriteria, ColorScaleCriterion, Comment,
    CommentAuthor, Comments, CommentsIter, ConditionValue, ConditionValueType, ConditionalFormat,
    ConditionalFormatType, ConditionalOperator, DataBar, DataBarAxisPosition, DataBarDirection,
    DataBarFillType, DisplayFormat, DuplicateMode, ExpressionRuleOptions, FormatColor,
    FormatCondition, FormatConditions, FormatConditionsIter, Hyperlink, HyperlinkAddOptions,
    Hyperlinks, HyperlinksIter, IconCriteria, IconCriterion, IconKind, IconSetCondition,
    IconSetKind, Style, Styles, StylesIter, TextCondition, TextConditionOperator, TextRuleOptions,
    ThemeColor, ThemeFont, ThreadedComment, ThreadedComments, ThreadedCommentsIter, TimePeriod,
    Top10, Top10Options, TopBottomRuleOptions, UniqueValues, UniqueValuesOptions,
    UnsupportedConditionalFormat,
};
pub use presentation::{
    AutomationSecurity, AutomationSecurityGuard, FixedFormatOptions, FixedFormatQuality,
    FixedFormatType, HPageBreak, HPageBreaks, Outline, PageBreakType, PageFit, PageOrientation,
    PageSetup, PageZoom, PaperSize, PrintErrors, PrintLocation, PrintOrder, PrintOutOptions,
    ReadingOrder, SafeWorkbookOpenOptions, Sheet, SheetDestination, SheetObject, SheetView, Sheets,
    SheetsIter, SummaryColumn, SummaryRow, Tab, VPageBreak, VPageBreaks, Window, WindowView,
    Windows, WorkbookProtectOptions, WorksheetProtectOptions,
};
pub use range::Range;
pub use range_structure::{
    DeleteShiftDirection, InsertFormatOrigin, InsertShiftDirection, PasteOperation,
    PasteSpecialOptions, PasteType, RangeInsertOptions, RemoveDuplicatesOptions,
};
pub use reference::{
    FormulaConversionOptions, RangeAddressOptions, ReferenceAbsoluteMode, ReferenceStyle,
};
pub use search::{
    FindLookIn, FindMatchMode, FindOptions, RangeFindIter, ReplaceOptions, SearchDirection,
    SearchOrder, SpecialCellType, SpecialCellValueMask,
};
pub use sort::{
    RangeSortOptions, Sort, SortDataOption, SortField, SortFields, SortMethod, SortOrder,
    SortOrientation,
};
pub use table::{
    ListColumn, ListColumns, ListColumnsIter, ListObject, ListObjectAddOptions,
    ListObjectSourceType, ListObjects, ListObjectsIter, ListRow, ListRows, ListRowsIter,
    TableHeaderMode, TotalsCalculation,
};
pub use validation::{
    Validation, ValidationAddOptions, ValidationAlertStyle, ValidationOperator, ValidationType,
};
pub use workbook::Workbook;
pub use workbooks::{Workbooks, WorkbooksIter};
pub use worksheet::{SheetVisibility, Worksheet, XlSheetVisibility};
pub use worksheets::{
    SheetType, WorksheetAddOptions, Worksheets, WorksheetsAddOptions, WorksheetsIter,
};

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
