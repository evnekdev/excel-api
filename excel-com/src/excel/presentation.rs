//! Worksheet presentation, printing, protection, and macro-runtime wrappers.
//!
//! This module deliberately keeps every COM object behind a typed, apartment-
//! bound wrapper. It exposes no raw dispatch pointer or `VARIANT` ownership.

use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;
use std::path::Path;

use windows_sys::Win32::System::Variant::VT_DISPATCH;

use crate::automation::{
    AutomationArgument, AutomationValue, ConversionPolicy, OwnedVariant, PositionalArguments,
    decode_variant, invoke, property_get, property_put,
};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index,
};
use crate::excel::formatting::{map_mixed, mixed_bool, mixed_i32, property_mixed_get};
use crate::excel::text::text_bstr;
use crate::excel::{
    Application, DispatchObject, ExcelColor, MixedValue, Range, ReferenceStyle, Workbook,
    WorkbookOpenOptions, Workbooks, Worksheet,
};
use crate::internal::{ComPtr, Dispatch, path_bstr};
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};

macro_rules! excel_i32 {
    ($(#[$meta:meta])* $name:ident { $($(#[$constant_meta:meta])* $constant:ident = $value:expr;)* }) => {
        $(#[$meta])*
        #[repr(transparent)]
        #[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
        pub struct $name(i32);
        impl $name {
            $($(#[$constant_meta])* pub const $constant: Self = Self($value);)*
            /// Builds the value from Excel's raw integer without discarding an unknown value.
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            /// Returns the raw integer passed to or read from Excel.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

excel_i32! {
    /// A forward-compatible Excel window view.
    WindowView {
        /// Normal worksheet view.
        NORMAL = 1;
        /// Page-break preview.
        PAGE_BREAK_PREVIEW = 2;
        /// Page-layout view.
        PAGE_LAYOUT = 3;
    }
}

/// Compatibility name for a sheet-window view value.
pub type SheetView = WindowView;

excel_i32! {
    /// A forward-compatible page orientation.
    PageOrientation {
        /// Portrait orientation.
        PORTRAIT = 1;
        /// Landscape orientation.
        LANDSCAPE = 2;
    }
}

excel_i32! {
    /// A forward-compatible paper-size value.
    PaperSize {
        /// Letter paper.
        LETTER = 1;
        /// Legal paper.
        LEGAL = 5;
        /// A4 paper.
        A4 = 9;
        /// Excel's printer-defined paper size.
        USER = 256;
    }
}

excel_i32! {
    /// Page traversal order.
    PrintOrder {
        /// Print down rows before moving across columns.
        DOWN_THEN_OVER = 1;
        /// Print across columns before moving down rows.
        OVER_THEN_DOWN = 2;
    }
}

excel_i32! {
    /// Header/footer print location for comments.
    PrintLocation {
        /// Do not print comments.
        NO_COMMENTS = -4142;
        /// Print comments at the end of the sheet.
        SHEET_END = 1;
        /// Print comments in place.
        IN_PLACE = 16;
    }
}

excel_i32! {
    /// How Excel renders cell errors in printed output.
    PrintErrors {
        /// Display errors as shown in the worksheet.
        DISPLAYED = 0;
        /// Print errors as blank cells.
        BLANK = 1;
        /// Print errors as dashes.
        DASH = 2;
        /// Print errors as `#N/A`.
        NA = 3;
    }
}

excel_i32! {
    /// Excel's range reading-order setting.
    ReadingOrder {
        /// Let Excel choose from the current language context.
        CONTEXT = -5002;
        /// Left-to-right reading order.
        LEFT_TO_RIGHT = -5003;
        /// Right-to-left reading order.
        RIGHT_TO_LEFT = -5004;
    }
}

excel_i32! {
    /// Location of outline summary rows.
    SummaryRow {
        /// Summary rows appear above details.
        ABOVE = 0;
        /// Summary rows appear below details.
        BELOW = 1;
    }
}

excel_i32! {
    /// Location of outline summary columns.
    SummaryColumn {
        /// Summary columns appear to the left of details.
        LEFT = -4131;
        /// Summary columns appear to the right of details.
        RIGHT = -4152;
    }
}

excel_i32! {
    /// A manual or automatic page-break classification.
    PageBreakType {
        /// Excel chose the break automatically.
        AUTOMATIC = -4105;
        /// A user-controlled manual break.
        MANUAL = -4135;
        /// No break is present.
        NONE = -4142;
    }
}

excel_i32! {
    /// Output format for `ExportAsFixedFormat`.
    FixedFormatType {
        /// Portable Document Format.
        PDF = 0;
        /// XML Paper Specification.
        XPS = 1;
    }
}

excel_i32! {
    /// Output quality for `ExportAsFixedFormat`.
    FixedFormatQuality {
        /// Standard quality.
        STANDARD = 0;
        /// Smaller, minimum-quality output.
        MINIMUM = 1;
    }
}

excel_i32! {
    /// Excel's process-global macro Automation security setting.
    AutomationSecurity {
        /// Enable macros using Excel's low-security behavior.
        LOW = 1;
        /// Follow the user's Excel UI setting.
        BY_UI = 2;
        /// Disable macros while a file is opened through Automation.
        FORCE_DISABLE = 3;
    }
}

/// A page zoom representation that preserves Excel's `bool | number` contract.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum PageZoom {
    /// Excel's automatic zoom behavior (`false`).
    Automatic,
    /// A numeric zoom percentage.
    Percent(i32),
}

/// Explicit page-fitting dimensions.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct PageFit {
    /// Number of pages wide, or `None` to leave that dimension Excel-controlled.
    pub wide: Option<usize>,
    /// Number of pages tall, or `None` to leave that dimension Excel-controlled.
    pub tall: Option<usize>,
}

/// A typed destination for worksheet copy or move operations.
#[derive(Clone, Copy, Debug)]
pub enum SheetDestination<'a> {
    /// Place the resulting sheet before this worksheet.
    Before(&'a Worksheet),
    /// Place the resulting sheet after this worksheet.
    After(&'a Worksheet),
    /// Let Excel create a new workbook containing the resulting sheet.
    NewWorkbook,
}

/// A non-worksheet object contained by the heterogeneous Excel `Sheets` collection.
pub struct SheetObject {
    inner: DispatchObject,
}

impl Debug for SheetObject {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("SheetObject")
            .field(&self.inner)
            .finish()
    }
}

impl Clone for SheetObject {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl SheetObject {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "SheetObject",
            },
        }
    }

    /// Returns the sheet name supplied by Excel.
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_string(&self.inner, "excel.worksheet.name")
    }

    /// Returns the Excel sheet-type value.
    pub fn sheet_type(&self) -> Result<i32, ExcelComError> {
        get_i32(&self.inner, "excel.worksheet.type", "Sheet.Type")
    }
}

/// A safe heterogeneous member of Excel's `Sheets` collection.
#[derive(Clone, Debug)]
pub enum Sheet {
    /// A normal worksheet with the full typed worksheet API.
    Worksheet(Worksheet),
    /// A chart, macro sheet, dialog sheet, or other non-worksheet object.
    Other(SheetObject),
}

impl Sheet {
    /// Returns the sheet name supplied by Excel.
    pub fn name(&self) -> Result<String, ExcelComError> {
        match self {
            Self::Worksheet(value) => value.name(),
            Self::Other(value) => value.name(),
        }
    }

    /// Returns the numeric Excel sheet type.
    pub fn sheet_type(&self) -> Result<i32, ExcelComError> {
        match self {
            Self::Worksheet(_) => Ok(-4167),
            Self::Other(value) => value.sheet_type(),
        }
    }
}

const SHEETS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Sheets",
    count: MemberId::new("excel.sheets.count"),
    item: MemberId::new("excel.sheets.item"),
    new_enum: MemberId::new("excel.sheets.newenum"),
};

/// Safe wrapper for Excel's heterogeneous `Sheets` collection.
pub struct Sheets {
    inner: DispatchObject,
}

impl Debug for Sheets {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Sheets").field(&self.inner).finish()
    }
}
impl Clone for Sheets {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Sheets {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Sheets",
            },
        }
    }
    /// Returns the number of sheets, including non-worksheets.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        count_i32(&self.inner, SHEETS_DESCRIPTOR)
    }
    /// Returns a one-based heterogeneous sheet member.
    pub fn item_by_index(&self, index: usize) -> Result<Sheet, ExcelComError> {
        sheet_from_dispatch(item_by_index(&self.inner, SHEETS_DESCRIPTOR, index)?)
    }
    /// Iterates every safe heterogeneous sheet member in Excel enum order.
    pub fn iter(&self) -> Result<SheetsIter, ExcelComError> {
        Ok(SheetsIter {
            enumerator: enumerator(&self.inner, SHEETS_DESCRIPTOR)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Single-pass iterator over [`Sheets`].
pub struct SheetsIter {
    enumerator: crate::automation::EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for SheetsIter {
    type Item = Result<Sheet, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(
                    crate::automation::enumerated_dispatch(&mut value, "Sheets", index)
                        .and_then(sheet_from_dispatch),
                )
            }
            Ok(None) => {
                self.terminal = true;
                None
            }
            Err(error) => {
                self.terminal = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for SheetsIter {}

const WINDOWS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Windows",
    count: MemberId::new("excel.windows.count"),
    item: MemberId::new("excel.windows.item"),
    new_enum: MemberId::new("excel.windows.newenum"),
};

/// A workbook or application window controlled through Excel Automation.
pub struct Window {
    inner: DispatchObject,
}
impl Debug for Window {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Window").field(&self.inner).finish()
    }
}
impl Clone for Window {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Window {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Window",
            },
        }
    }
    /// Activates this Excel window.
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.window.activate", vec![])
    }
    /// Returns the window's one-based Excel index.
    pub fn index(&self) -> Result<i32, ExcelComError> {
        get_i32(&self.inner, "excel.window.index", "Window.Index")
    }
    /// Returns whether gridlines are displayed.
    pub fn display_gridlines(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.window.displaygridlines")
    }
    /// Changes whether gridlines are displayed.
    pub fn set_display_gridlines(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.displaygridlines",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether headings are displayed.
    pub fn display_headings(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.window.displayheadings")
    }
    /// Changes whether headings are displayed.
    pub fn set_display_headings(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.displayheadings",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether zero values are displayed.
    pub fn display_zeros(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.window.displayzeros")
    }
    /// Changes whether zero values are displayed.
    pub fn set_display_zeros(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.displayzeros",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether panes are frozen.
    pub fn freeze_panes(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.window.freezepanes")
    }
    /// Changes the window's frozen-panes state.
    pub fn set_freeze_panes(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.freezepanes",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the first visible column.
    pub fn scroll_column(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.window.scrollcolumn",
            "Window.ScrollColumn",
        )
    }
    /// Sets the first visible column.
    pub fn set_scroll_column(&self, value: usize) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.scrollcolumn",
            one_based(value, "Window.ScrollColumn")?,
        )
    }
    /// Returns the first visible row.
    pub fn scroll_row(&self) -> Result<i32, ExcelComError> {
        get_i32(&self.inner, "excel.window.scrollrow", "Window.ScrollRow")
    }
    /// Sets the first visible row.
    pub fn set_scroll_row(&self, value: usize) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.scrollrow",
            one_based(value, "Window.ScrollRow")?,
        )
    }
    /// Returns the horizontal split position in points.
    pub fn split_column(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.window.splitcolumn",
            "Window.SplitColumn",
        )
    }
    /// Sets the horizontal split position in points.
    pub fn set_split_column(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value)?;
        put(
            &self.inner,
            "excel.window.splitcolumn",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the vertical split position in points.
    pub fn split_row(&self) -> Result<f64, ExcelComError> {
        get_f64(&self.inner, "excel.window.splitrow", "Window.SplitRow")
    }
    /// Sets the vertical split position in points.
    pub fn set_split_row(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value)?;
        put(
            &self.inner,
            "excel.window.splitrow",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the split column count.
    pub fn split_horizontal(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.window.splithorizontal",
            "Window.SplitHorizontal",
        )
    }
    /// Returns the split row count.
    pub fn split_vertical(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.window.splitvertical",
            "Window.SplitVertical",
        )
    }
    /// Returns the zoom percentage or Excel's automatic setting.
    pub fn zoom(&self) -> Result<PageZoom, ExcelComError> {
        page_zoom_get(&self.inner, "excel.window.zoom")
    }
    /// Sets a numeric zoom percentage or automatic zoom.
    pub fn set_zoom(&self, value: PageZoom) -> Result<(), ExcelComError> {
        page_zoom_put(&self.inner, "excel.window.zoom", value)
    }
    /// Returns the window view.
    pub fn view(&self) -> Result<WindowView, ExcelComError> {
        Ok(WindowView::from_raw(get_i32(
            &self.inner,
            "excel.window.view",
            "Window.View",
        )?))
    }
    /// Changes the window view.
    pub fn set_view(&self, value: WindowView) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.window.view",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Selects `cell` and freezes panes at that selection.
    pub fn freeze_at(&self, cell: &Range) -> Result<(), ExcelComError> {
        cell.select()?;
        self.set_freeze_panes(false)?;
        self.set_freeze_panes(true)
    }
}

/// Safe collection of Excel windows.
pub struct Windows {
    inner: DispatchObject,
}
impl Debug for Windows {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Windows").field(&self.inner).finish()
    }
}
impl Clone for Windows {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Windows {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Windows",
            },
        }
    }
    /// Returns the number of Excel windows.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        count_i32(&self.inner, WINDOWS_DESCRIPTOR)
    }
    /// Returns a one-based window.
    pub fn item_by_index(&self, index: usize) -> Result<Window, ExcelComError> {
        Ok(Window::from_dispatch(item_by_index(
            &self.inner,
            WINDOWS_DESCRIPTOR,
            index,
        )?))
    }
}

/// The worksheet tab formatting object.
pub struct Tab {
    inner: DispatchObject,
}
impl Debug for Tab {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Tab").field(&self.inner).finish()
    }
}
impl Clone for Tab {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Tab {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Tab",
            },
        }
    }
    /// Returns the tab RGB color, or `Mixed`/`Empty` as returned by Excel.
    pub fn color(&self) -> Result<MixedValue<ExcelColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.tab.color", |value| {
            mixed_i32(value).map(|v| map_mixed(v, ExcelColor::from_raw))
        })
    }
    /// Sets the tab RGB color in Excel's low-byte-red COLORREF order.
    pub fn set_color(&self, value: ExcelColor) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.tab.color",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Clears the explicit tab color so Excel selects its normal tab color.
    pub fn clear_color(&self) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.tab.colorindex",
            OwnedVariant::i32(-4142),
        )
    }
}

/// Worksheet outline configuration.
pub struct Outline {
    inner: DispatchObject,
}
impl Debug for Outline {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Outline").field(&self.inner).finish()
    }
}
impl Clone for Outline {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Outline {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Outline",
            },
        }
    }
    /// Returns whether Excel applies automatic outline styles.
    pub fn automatic_styles(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.outline.automaticstyles")
    }
    /// Enables or disables Excel's automatic outline styles.
    pub fn set_automatic_styles(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.outline.automaticstyles",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the summary-row position.
    pub fn summary_row(&self) -> Result<SummaryRow, ExcelComError> {
        Ok(SummaryRow::from_raw(get_i32(
            &self.inner,
            "excel.outline.summaryrow",
            "Outline.SummaryRow",
        )?))
    }
    /// Changes the summary-row position.
    pub fn set_summary_row(&self, value: SummaryRow) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.outline.summaryrow",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the summary-column position.
    pub fn summary_column(&self) -> Result<SummaryColumn, ExcelComError> {
        Ok(SummaryColumn::from_raw(get_i32(
            &self.inner,
            "excel.outline.summarycolumn",
            "Outline.SummaryColumn",
        )?))
    }
    /// Changes the summary-column position.
    pub fn set_summary_column(&self, value: SummaryColumn) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.outline.summarycolumn",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Shows the requested row and column outline levels.
    pub fn show_levels(
        &self,
        row_level: Option<usize>,
        column_level: Option<usize>,
    ) -> Result<(), ExcelComError> {
        let mut arguments = PositionalArguments::new();
        arguments.push_optional(
            row_level
                .map(|v| one_based(v, "Outline.ShowLevels row"))
                .transpose()?,
        );
        arguments.push_optional(
            column_level
                .map(|v| one_based(v, "Outline.ShowLevels column"))
                .transpose()?,
        );
        call(
            &self.inner,
            "excel.outline.showlevels",
            arguments.into_inner(),
        )
    }
}

/// Page layout, print-title, and header/footer settings for a worksheet.
pub struct PageSetup {
    inner: DispatchObject,
}
impl Debug for PageSetup {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PageSetup").field(&self.inner).finish()
    }
}
impl Clone for PageSetup {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PageSetup {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PageSetup",
            },
        }
    }
    /// Returns the page orientation.
    pub fn orientation(&self) -> Result<PageOrientation, ExcelComError> {
        Ok(PageOrientation::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.orientation",
            "PageSetup.Orientation",
        )?))
    }
    /// Changes the page orientation.
    pub fn set_orientation(&self, value: PageOrientation) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.orientation",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the selected paper size.
    pub fn paper_size(&self) -> Result<PaperSize, ExcelComError> {
        Ok(PaperSize::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.papersize",
            "PageSetup.PaperSize",
        )?))
    }
    /// Changes the paper size.
    pub fn set_paper_size(&self, value: PaperSize) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.papersize",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the page traversal order.
    pub fn order(&self) -> Result<PrintOrder, ExcelComError> {
        Ok(PrintOrder::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.order",
            "PageSetup.Order",
        )?))
    }
    /// Changes the page traversal order.
    pub fn set_order(&self, value: PrintOrder) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.order",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the zoom setting.
    pub fn zoom(&self) -> Result<PageZoom, ExcelComError> {
        page_zoom_get(&self.inner, "excel.pagesetup.zoom")
    }
    /// Changes the zoom setting.
    pub fn set_zoom(&self, value: PageZoom) -> Result<(), ExcelComError> {
        page_zoom_put(&self.inner, "excel.pagesetup.zoom", value)
    }
    /// Sets both fit-to-pages dimensions, validating positive explicit values.
    pub fn set_fit_to_pages(&self, value: PageFit) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.fittopageswide",
            optional_positive(value.wide, "PageSetup.FitToPagesWide")?,
        )?;
        put(
            &self.inner,
            "excel.pagesetup.fittopagestall",
            optional_positive(value.tall, "PageSetup.FitToPagesTall")?,
        )
    }
    /// Returns the configured fit-to-pages dimensions.
    ///
    /// Excel represents an unconstrained dimension as a boolean value and an
    /// explicit dimension as a number; the public form makes that distinction
    /// `None` versus a positive page count.
    pub fn fit_to_pages(&self) -> Result<PageFit, ExcelComError> {
        Ok(PageFit {
            wide: page_fit_dimension_get(&self.inner, "excel.pagesetup.fittopageswide")?,
            tall: page_fit_dimension_get(&self.inner, "excel.pagesetup.fittopagestall")?,
        })
    }
    /// Returns the left page margin in points.
    pub fn left_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.leftmargin",
            "PageSetup.LeftMargin",
        )
    }
    /// Sets the left page margin in points.
    pub fn set_left_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.leftmargin",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the right page margin in points.
    pub fn right_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.rightmargin",
            "PageSetup.RightMargin",
        )
    }
    /// Sets the right page margin in points.
    pub fn set_right_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.rightmargin",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the top page margin in points.
    pub fn top_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.topmargin",
            "PageSetup.TopMargin",
        )
    }
    /// Sets the top page margin in points.
    pub fn set_top_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.topmargin",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the bottom page margin in points.
    pub fn bottom_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.bottommargin",
            "PageSetup.BottomMargin",
        )
    }
    /// Sets the bottom page margin in points.
    pub fn set_bottom_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.bottommargin",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the header margin in points.
    pub fn header_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.headermargin",
            "PageSetup.HeaderMargin",
        )
    }
    /// Sets the header margin in points.
    pub fn set_header_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.headermargin",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the footer margin in points.
    pub fn footer_margin(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.pagesetup.footermargin",
            "PageSetup.FooterMargin",
        )
    }
    /// Sets the footer margin in points.
    pub fn set_footer_margin(&self, value: f64) -> Result<(), ExcelComError> {
        nonnegative(value)?;
        put(
            &self.inner,
            "excel.pagesetup.footermargin",
            OwnedVariant::f64(value),
        )
    }
    /// Returns whether row and column headings print.
    pub fn print_headings(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.printheadings")
    }
    /// Changes whether row and column headings print.
    pub fn set_print_headings(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printheadings",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether gridlines print.
    pub fn print_gridlines(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.printgridlines")
    }
    /// Changes whether gridlines print.
    pub fn set_print_gridlines(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printgridlines",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether pages are centred horizontally.
    pub fn center_horizontally(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.centerhorizontally")
    }
    /// Changes horizontal page centring.
    pub fn set_center_horizontally(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.centerhorizontally",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether pages are centred vertically.
    pub fn center_vertically(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.centervertically")
    }
    /// Changes vertical page centring.
    pub fn set_center_vertically(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.centervertically",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether output is black and white.
    pub fn black_and_white(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.blackandwhite")
    }
    /// Changes black-and-white output.
    pub fn set_black_and_white(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.blackandwhite",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether output is draft quality.
    pub fn draft(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.pagesetup.draft")
    }
    /// Changes draft-quality output.
    pub fn set_draft(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.draft",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the first printed page number.
    pub fn first_page_number(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.pagesetup.firstpagenumber",
            "PageSetup.FirstPageNumber",
        )
    }
    /// Changes the first printed page number.
    pub fn set_first_page_number(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.firstpagenumber",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the print-comments location.
    pub fn print_comments(&self) -> Result<PrintLocation, ExcelComError> {
        Ok(PrintLocation::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.printcomments",
            "PageSetup.PrintComments",
        )?))
    }
    /// Changes the print-comments location.
    pub fn set_print_comments(&self, value: PrintLocation) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printcomments",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns Excel's current print-quality value.
    pub fn print_quality(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.pagesetup.printquality",
            "PageSetup.PrintQuality",
        )
    }
    /// Sets Excel's print-quality value for the selected printer.
    pub fn set_print_quality(&self, value: i32) -> Result<(), ExcelComError> {
        if value <= 0 {
            return Err(ExcelComError::Unsupported {
                detail: "PageSetup.PrintQuality must be positive",
            });
        }
        put(
            &self.inner,
            "excel.pagesetup.printquality",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the selected cell-error rendering behavior.
    pub fn print_errors(&self) -> Result<PrintErrors, ExcelComError> {
        Ok(PrintErrors::from_raw(get_i32(
            &self.inner,
            "excel.pagesetup.printerrors",
            "PageSetup.PrintErrors",
        )?))
    }
    /// Changes the cell-error rendering behavior.
    pub fn set_print_errors(&self, value: PrintErrors) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printerrors",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the print area, treating Excel's empty string as no explicit area.
    pub fn print_area(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.printarea")
    }
    /// Sets an external address as the print area.
    pub fn set_print_area(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printarea",
            text_bstr(&range.external_address(ReferenceStyle::A1)?)?,
        )
    }
    /// Clears Excel's explicit print area.
    pub fn clear_print_area(&self) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.pagesetup.printarea", text_bstr("")?)
    }
    /// Returns the repeating title rows, treating empty as absent.
    pub fn print_title_rows(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.printtitlerows")
    }
    /// Sets external-address title rows.
    pub fn set_print_title_rows(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printtitlerows",
            text_bstr(&range.external_address(ReferenceStyle::A1)?)?,
        )
    }
    /// Clears repeating title rows.
    pub fn clear_print_title_rows(&self) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printtitlerows",
            text_bstr("")?,
        )
    }
    /// Returns the repeating title columns, treating empty as absent.
    pub fn print_title_columns(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.printtitlecolumns")
    }
    /// Sets external-address title columns.
    pub fn set_print_title_columns(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printtitlecolumns",
            text_bstr(&range.external_address(ReferenceStyle::A1)?)?,
        )
    }
    /// Clears repeating title columns.
    pub fn clear_print_title_columns(&self) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.printtitlecolumns",
            text_bstr("")?,
        )
    }
    /// Returns the left header, treating an empty string as absent.
    pub fn left_header(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.leftheader")
    }
    /// Sets the left header text.
    pub fn set_left_header(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.pagesetup.leftheader", text_bstr(value)?)
    }
    /// Returns the centre header, treating an empty string as absent.
    pub fn center_header(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.centerheader")
    }
    /// Sets the centre header text.
    pub fn set_center_header(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.centerheader",
            text_bstr(value)?,
        )
    }
    /// Returns the right header, treating an empty string as absent.
    pub fn right_header(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.rightheader")
    }
    /// Sets the right header text.
    pub fn set_right_header(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.rightheader",
            text_bstr(value)?,
        )
    }
    /// Returns the left footer, treating an empty string as absent.
    pub fn left_footer(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.leftfooter")
    }
    /// Sets the left footer text.
    pub fn set_left_footer(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.pagesetup.leftfooter", text_bstr(value)?)
    }
    /// Returns the centre footer, treating an empty string as absent.
    pub fn center_footer(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.centerfooter")
    }
    /// Sets the centre footer text.
    pub fn set_center_footer(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.centerfooter",
            text_bstr(value)?,
        )
    }
    /// Returns the right footer, treating an empty string as absent.
    pub fn right_footer(&self) -> Result<Option<String>, ExcelComError> {
        optional_string(&self.inner, "excel.pagesetup.rightfooter")
    }
    /// Sets the right footer text.
    pub fn set_right_footer(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pagesetup.rightfooter",
            text_bstr(value)?,
        )
    }
}

/// Password and permission positions accepted by `Worksheet.Protect`.
#[derive(Clone, Default, PartialEq)]
pub struct WorksheetProtectOptions<'a> {
    /// Optional password, redacted by `Debug`.
    pub password: Option<&'a str>,
    /// Protect drawing objects.
    pub drawing_objects: Option<bool>,
    /// Protect cell contents.
    pub contents: Option<bool>,
    /// Protect scenarios.
    pub scenarios: Option<bool>,
    /// Let VBA macros change protected cells without changing the UI permission.
    pub user_interface_only: Option<bool>,
    /// Permit formatting cells.
    pub allow_formatting_cells: Option<bool>,
    /// Permit formatting columns.
    pub allow_formatting_columns: Option<bool>,
    /// Permit formatting rows.
    pub allow_formatting_rows: Option<bool>,
    /// Permit inserting columns.
    pub allow_inserting_columns: Option<bool>,
    /// Permit inserting rows.
    pub allow_inserting_rows: Option<bool>,
    /// Permit inserting hyperlinks.
    pub allow_inserting_hyperlinks: Option<bool>,
    /// Permit deleting columns.
    pub allow_deleting_columns: Option<bool>,
    /// Permit deleting rows.
    pub allow_deleting_rows: Option<bool>,
    /// Permit sorting.
    pub allow_sorting: Option<bool>,
    /// Permit filtering.
    pub allow_filtering: Option<bool>,
    /// Permit PivotTable use.
    pub allow_using_pivot_tables: Option<bool>,
}
impl Debug for WorksheetProtectOptions<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WorksheetProtectOptions")
            .field("password", &self.password.map(|_| "REDACTED"))
            .field("drawing_objects", &self.drawing_objects)
            .field("contents", &self.contents)
            .field("scenarios", &self.scenarios)
            .field("user_interface_only", &self.user_interface_only)
            .field("allow_formatting_cells", &self.allow_formatting_cells)
            .field("allow_formatting_columns", &self.allow_formatting_columns)
            .field("allow_formatting_rows", &self.allow_formatting_rows)
            .field("allow_inserting_columns", &self.allow_inserting_columns)
            .field("allow_inserting_rows", &self.allow_inserting_rows)
            .field(
                "allow_inserting_hyperlinks",
                &self.allow_inserting_hyperlinks,
            )
            .field("allow_deleting_columns", &self.allow_deleting_columns)
            .field("allow_deleting_rows", &self.allow_deleting_rows)
            .field("allow_sorting", &self.allow_sorting)
            .field("allow_filtering", &self.allow_filtering)
            .field("allow_using_pivot_tables", &self.allow_using_pivot_tables)
            .finish()
    }
}

/// Password and permission positions accepted by `Workbook.Protect`.
#[derive(Clone, Default, PartialEq)]
pub struct WorkbookProtectOptions<'a> {
    /// Optional password, redacted by `Debug`.
    pub password: Option<&'a str>,
    /// Protect the workbook structure.
    pub structure: Option<bool>,
    /// Protect workbook windows.
    pub windows: Option<bool>,
}
impl Debug for WorkbookProtectOptions<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WorkbookProtectOptions")
            .field("password", &self.password.map(|_| "REDACTED"))
            .field("structure", &self.structure)
            .field("windows", &self.windows)
            .finish()
    }
}

/// Typed options for Excel's `PrintOut` method.
#[derive(Clone, Default, PartialEq)]
pub struct PrintOutOptions<'a> {
    /// First page to print.
    pub from: Option<usize>,
    /// Last page to print.
    pub to: Option<usize>,
    /// Number of copies.
    pub copies: Option<usize>,
    /// Request Excel's print preview instead of printing.
    pub preview: Option<bool>,
    /// Explicit printer name; potentially sensitive and redacted by `Debug`.
    pub active_printer: Option<&'a str>,
    /// Write print data to a file.
    pub print_to_file: Option<bool>,
    /// Collate copies.
    pub collate: Option<bool>,
    /// Print-to-file destination; potentially sensitive and redacted by `Debug`.
    pub pr_to_file_name: Option<&'a str>,
    /// Ignore a worksheet's configured print area.
    pub ignore_print_areas: Option<bool>,
}
impl Debug for PrintOutOptions<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("PrintOutOptions")
            .field("from", &self.from)
            .field("to", &self.to)
            .field("copies", &self.copies)
            .field("preview", &self.preview)
            .field("active_printer", &self.active_printer.map(|_| "REDACTED"))
            .field("print_to_file", &self.print_to_file)
            .field("collate", &self.collate)
            .field("pr_to_file_name", &self.pr_to_file_name.map(|_| "REDACTED"))
            .field("ignore_print_areas", &self.ignore_print_areas)
            .finish()
    }
}

/// Typed optional positions for `ExportAsFixedFormat`.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct FixedFormatOptions<'a> {
    /// Optional target path. Excel chooses one when absent.
    pub output: Option<&'a Path>,
    /// Output quality.
    pub quality: Option<FixedFormatQuality>,
    /// Include document properties.
    pub include_doc_properties: Option<bool>,
    /// Ignore any configured print area.
    pub ignore_print_areas: Option<bool>,
    /// First page to export.
    pub from: Option<usize>,
    /// Last page to export.
    pub to: Option<usize>,
    /// Ask Excel to open the resulting file.
    pub open_after_publish: Option<bool>,
}

/// A horizontal page break.
pub struct HPageBreak {
    inner: DispatchObject,
}
impl Debug for HPageBreak {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("HPageBreak").field(&self.inner).finish()
    }
}
impl Clone for HPageBreak {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl HPageBreak {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "HPageBreak",
            },
        }
    }
    /// Returns the page-break location range.
    pub fn location(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.hpagebreak.location")
    }
    /// Moves the page break to `range`.
    pub fn set_location(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.hpagebreak.location",
            OwnedVariant::dispatch_borrowed(&range.dispatch_object().dispatch),
        )
    }
    /// Returns Excel's break classification.
    pub fn break_type(&self) -> Result<PageBreakType, ExcelComError> {
        Ok(PageBreakType::from_raw(get_i32(
            &self.inner,
            "excel.hpagebreak.type",
            "HPageBreak.Type",
        )?))
    }
    /// Removes this break from Excel.
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.hpagebreak.delete", vec![])
    }
}

/// A vertical page break.
pub struct VPageBreak {
    inner: DispatchObject,
}
impl Debug for VPageBreak {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("VPageBreak").field(&self.inner).finish()
    }
}
impl Clone for VPageBreak {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl VPageBreak {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "VPageBreak",
            },
        }
    }
    /// Returns the page-break location range.
    pub fn location(&self) -> Result<Range, ExcelComError> {
        get_range(&self.inner, "excel.vpagebreak.location")
    }
    /// Moves the page break to `range`.
    pub fn set_location(&self, range: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.vpagebreak.location",
            OwnedVariant::dispatch_borrowed(&range.dispatch_object().dispatch),
        )
    }
    /// Returns Excel's break classification.
    pub fn break_type(&self) -> Result<PageBreakType, ExcelComError> {
        Ok(PageBreakType::from_raw(get_i32(
            &self.inner,
            "excel.vpagebreak.type",
            "VPageBreak.Type",
        )?))
    }
    /// Removes this break from Excel.
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.vpagebreak.delete", vec![])
    }
}

const HPAGE_BREAKS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "HPageBreaks",
    count: MemberId::new("excel.hpagebreaks.count"),
    item: MemberId::new("excel.hpagebreaks.item"),
    new_enum: MemberId::new("excel.hpagebreaks.newenum"),
};
const VPAGE_BREAKS_DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "VPageBreaks",
    count: MemberId::new("excel.vpagebreaks.count"),
    item: MemberId::new("excel.vpagebreaks.item"),
    new_enum: MemberId::new("excel.vpagebreaks.newenum"),
};

/// Horizontal page-break collection for one worksheet.
pub struct HPageBreaks {
    inner: DispatchObject,
}
impl Debug for HPageBreaks {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("HPageBreaks").field(&self.inner).finish()
    }
}
impl Clone for HPageBreaks {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl HPageBreaks {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "HPageBreaks",
            },
        }
    }
    /// Returns the horizontal page-break count.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        count_i32(&self.inner, HPAGE_BREAKS_DESCRIPTOR)
    }
    /// Returns a one-based horizontal break.
    pub fn item_by_index(&self, index: usize) -> Result<HPageBreak, ExcelComError> {
        Ok(HPageBreak::from_dispatch(item_by_index(
            &self.inner,
            HPAGE_BREAKS_DESCRIPTOR,
            index,
        )?))
    }
    /// Adds a break above `before`.
    pub fn add(&self, before: &Range) -> Result<HPageBreak, ExcelComError> {
        add_page_break(
            &self.inner,
            "excel.hpagebreaks.add",
            before,
            HPageBreak::from_dispatch,
        )
    }
}

/// Vertical page-break collection for one worksheet.
pub struct VPageBreaks {
    inner: DispatchObject,
}
impl Debug for VPageBreaks {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("VPageBreaks").field(&self.inner).finish()
    }
}
impl Clone for VPageBreaks {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl VPageBreaks {
    fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "VPageBreaks",
            },
        }
    }
    /// Returns the vertical page-break count.
    pub fn count(&self) -> Result<i32, ExcelComError> {
        count_i32(&self.inner, VPAGE_BREAKS_DESCRIPTOR)
    }
    /// Returns a one-based vertical break.
    pub fn item_by_index(&self, index: usize) -> Result<VPageBreak, ExcelComError> {
        Ok(VPageBreak::from_dispatch(item_by_index(
            &self.inner,
            VPAGE_BREAKS_DESCRIPTOR,
            index,
        )?))
    }
    /// Adds a break to the left of `before`.
    pub fn add(&self, before: &Range) -> Result<VPageBreak, ExcelComError> {
        add_page_break(
            &self.inner,
            "excel.vpagebreaks.add",
            before,
            VPageBreak::from_dispatch,
        )
    }
}

/// Restores the previous process-global Automation macro-security setting on drop.
pub struct AutomationSecurityGuard<'a> {
    application: &'a Application,
    previous: AutomationSecurity,
    active: bool,
}
impl Debug for AutomationSecurityGuard<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("AutomationSecurityGuard")
            .field("previous", &self.previous)
            .field("active", &self.active)
            .finish()
    }
}
impl AutomationSecurityGuard<'_> {
    /// Restores the previous setting and disarms the guard.
    pub fn restore(mut self) -> Result<(), ExcelComError> {
        self.application.set_automation_security(self.previous)?;
        self.active = false;
        Ok(())
    }
}
impl Drop for AutomationSecurityGuard<'_> {
    fn drop(&mut self) {
        if self.active {
            let _ = self.application.set_automation_security(self.previous);
            self.active = false;
        }
    }
}

/// Options for opening a workbook with macros forcibly disabled.
#[derive(Clone, Default, PartialEq)]
pub struct SafeWorkbookOpenOptions<'a> {
    /// Underlying `Workbooks.Open` optional positions.
    pub open: WorkbookOpenOptions<'a>,
}
impl Debug for SafeWorkbookOpenOptions<'_> {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("SafeWorkbookOpenOptions")
            .field("open", &self.open)
            .finish()
    }
}

impl Application {
    /// Returns the active workbook.
    pub fn active_workbook(&self) -> Result<Workbook, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.activeworkbook",
            Workbook::from_dispatch,
        )
    }
    /// Returns the active sheet without exposing a raw heterogeneous dispatch object.
    pub fn active_sheet(&self) -> Result<Sheet, ExcelComError> {
        get_sheet(self.dispatch_object(), "excel.application.activesheet")
    }
    /// Returns the active cell.
    pub fn active_cell(&self) -> Result<Range, ExcelComError> {
        get_range(self.dispatch_object(), "excel.application.activecell")
    }
    /// Returns Excel's current selection when it is a Range.
    ///
    /// Excel can select non-range objects such as charts; in that case its COM
    /// conversion failure is retained rather than exposing an untyped dispatch.
    pub fn active_range(&self) -> Result<Range, ExcelComError> {
        get_range(self.dispatch_object(), "excel.application.selection")
    }
    /// Returns the active window.
    pub fn active_window(&self) -> Result<Window, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.activewindow",
            Window::from_dispatch,
        )
    }
    /// Returns every open Excel window.
    pub fn windows(&self) -> Result<Windows, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.windows",
            Windows::from_dispatch,
        )
    }
    /// Returns every sheet, including chart and macro sheets, through safe typed variants.
    pub fn sheets(&self) -> Result<Sheets, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.sheets",
            Sheets::from_dispatch,
        )
    }
    /// Returns the application worksheet collection.
    pub fn worksheets(&self) -> Result<crate::excel::Worksheets, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.application.worksheets",
            crate::excel::Worksheets::from_dispatch,
        )
    }
    /// Activates the requested range, optionally selecting it after navigation.
    pub fn go_to(&self, range: &Range, select: Option<bool>) -> Result<(), ExcelComError> {
        let mut arguments = PositionalArguments::new();
        arguments.push_object(range.dispatch_object());
        arguments.push_optional(select.map(OwnedVariant::bool));
        call(
            self.dispatch_object(),
            "excel.application.goto",
            arguments.into_inner(),
        )
    }
    /// Returns the process-global Automation macro-security setting.
    pub fn automation_security(&self) -> Result<AutomationSecurity, ExcelComError> {
        Ok(AutomationSecurity::from_raw(get_i32(
            self.dispatch_object(),
            "excel.application.automationsecurity",
            "Application.AutomationSecurity",
        )?))
    }
    /// Changes the process-global Automation macro-security setting.
    pub fn set_automation_security(&self, value: AutomationSecurity) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.application.automationsecurity",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Temporarily changes macro Automation security and restores it on drop.
    pub fn automation_security_guard(
        &self,
        value: AutomationSecurity,
    ) -> Result<AutomationSecurityGuard<'_>, ExcelComError> {
        let previous = self.automation_security()?;
        self.set_automation_security(value)?;
        Ok(AutomationSecurityGuard {
            application: self,
            previous,
            active: true,
        })
    }
    /// Runs a public Excel macro with up to thirty positional Automation arguments.
    ///
    /// Returned dispatch values are refused because this high-level API does
    /// not expose raw objects; macros returning Range-like objects should use a
    /// dedicated typed API in a future slice.
    pub fn run_macro(
        &self,
        macro_name: &str,
        arguments: &[AutomationValue],
    ) -> Result<AutomationValue, ExcelComError> {
        if arguments.len() > 30 {
            return Err(ExcelComError::Unsupported {
                detail: "Application.Run accepts at most 30 macro arguments",
            });
        }
        let mut values = PositionalArguments::new();
        values.push_result(text_bstr(macro_name))?;
        for argument in arguments {
            values.push_argument(
                AutomationArgument::Value(argument.clone()),
                ConversionPolicy::default(),
            )?;
        }
        let result = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.application.run"), false),
            values.into_inner(),
            false,
        )?;
        if result.vt() == VT_DISPATCH {
            return Err(ExcelComError::Unsupported {
                detail: "Application.Run returned a dispatch object",
            });
        }
        decode_variant(&result, ConversionPolicy::default())
    }
}

impl Workbooks {
    /// Opens a workbook while a guard forces Automation macros to remain disabled.
    pub fn open_safely<P: AsRef<Path>>(
        &self,
        filename: P,
        options: SafeWorkbookOpenOptions<'_>,
    ) -> Result<Workbook, ExcelComError> {
        let application: Application = get_object(
            self.dispatch_object(),
            "excel.workbooks.application",
            Application::from_dispatch,
        )?;
        let guard = application.automation_security_guard(AutomationSecurity::FORCE_DISABLE)?;
        let opened = self.open(filename, options.open);
        let restored = guard.restore();
        match (opened, restored) {
            (Ok(workbook), Ok(())) => Ok(workbook),
            (Err(error), _) => Err(error),
            (_, Err(error)) => Err(error),
        }
    }
}

impl Workbook {
    /// Activates this workbook.
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.workbook.activate", vec![])
    }
    /// Returns the active sheet as a safe heterogeneous variant.
    pub fn active_sheet(&self) -> Result<Sheet, ExcelComError> {
        get_sheet(self.dispatch_object(), "excel.workbook.activesheet")
    }
    /// Returns all sheets, including chart and macro sheets.
    pub fn sheets(&self) -> Result<Sheets, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.workbook.sheets",
            Sheets::from_dispatch,
        )
    }
    /// Returns workbook windows.
    pub fn windows(&self) -> Result<Windows, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.workbook.windows",
            Windows::from_dispatch,
        )
    }
    /// Returns whether the workbook contains a VBA project.
    pub fn has_vb_project(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.workbook.hasvbproject")
    }
    /// Returns whether workbook structure is protected.
    pub fn protect_structure(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.workbook.protectstructure")
    }
    /// Returns whether workbook windows are protected.
    pub fn protect_windows(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.workbook.protectwindows")
    }
    /// Protects workbook structure and windows using explicit optional positions.
    pub fn protect(&self, options: &WorkbookProtectOptions<'_>) -> Result<(), ExcelComError> {
        let mut arguments = PositionalArguments::new();
        push_optional_text(&mut arguments, options.password)?;
        arguments.push_optional(options.structure.map(OwnedVariant::bool));
        arguments.push_optional(options.windows.map(OwnedVariant::bool));
        call(
            self.dispatch_object(),
            "excel.workbook.protect-2029",
            arguments.into_inner(),
        )
    }
    /// Removes workbook protection using an optional password.
    pub fn unprotect(&self, password: Option<&str>) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        push_optional_text(&mut a, password)?;
        call(
            self.dispatch_object(),
            "excel.workbook.unprotect",
            a.into_inner(),
        )
    }
    /// Shows Excel's workbook print preview.
    pub fn print_preview(&self) -> Result<(), ExcelComError> {
        call(
            self.dispatch_object(),
            "excel.workbook.printpreview",
            vec![],
        )
    }
    /// Sends the workbook to Excel's `PrintOut` entry point.
    pub fn print_out(&self, options: &PrintOutOptions<'_>) -> Result<(), ExcelComError> {
        print_out(
            self.dispatch_object(),
            "excel.workbook.printout-2361",
            options,
        )
    }
    /// Exports the workbook through Excel's fixed-format renderer.
    pub fn export_as_fixed_format(
        &self,
        format: FixedFormatType,
        options: &FixedFormatOptions<'_>,
    ) -> Result<(), ExcelComError> {
        fixed_format(
            self.dispatch_object(),
            "excel.workbook.exportasfixedformat-3175",
            format,
            options,
        )
    }
}

impl Worksheet {
    /// Activates this worksheet.
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.worksheet.activate", vec![])
    }
    /// Selects this worksheet, optionally replacing the selection.
    pub fn select(&self, replace: Option<bool>) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_optional(replace.map(OwnedVariant::bool));
        call(
            self.dispatch_object(),
            "excel.worksheet.select",
            a.into_inner(),
        )
    }
    /// Copies this worksheet to a typed destination.
    pub fn copy_to(&self, destination: SheetDestination<'_>) -> Result<(), ExcelComError> {
        worksheet_copy_move(self, "excel.worksheet.copy", destination)
    }
    /// Moves this worksheet to a typed destination.
    pub fn move_to(&self, destination: SheetDestination<'_>) -> Result<(), ExcelComError> {
        worksheet_copy_move(self, "excel.worksheet.move", destination)
    }
    /// Deletes this worksheet and consumes its wrapper.
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.worksheet.delete", vec![])
    }
    /// Returns the sheet tab formatting object.
    pub fn tab(&self) -> Result<Tab, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.tab",
            Tab::from_dispatch,
        )
    }
    /// Returns page setup for this worksheet.
    pub fn page_setup(&self) -> Result<PageSetup, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.pagesetup",
            PageSetup::from_dispatch,
        )
    }
    /// Returns outline settings for this worksheet.
    pub fn outline(&self) -> Result<Outline, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.outline",
            Outline::from_dispatch,
        )
    }
    /// Returns horizontal page breaks.
    pub fn h_page_breaks(&self) -> Result<HPageBreaks, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.hpagebreaks",
            HPageBreaks::from_dispatch,
        )
    }
    /// Returns vertical page breaks.
    pub fn v_page_breaks(&self) -> Result<VPageBreaks, ExcelComError> {
        get_object(
            self.dispatch_object(),
            "excel.worksheet.vpagebreaks",
            VPageBreaks::from_dispatch,
        )
    }
    /// Asks Excel to reset all page breaks for this worksheet.
    pub fn reset_all_page_breaks(&self) -> Result<(), ExcelComError> {
        call(
            self.dispatch_object(),
            "excel.worksheet.resetallpagebreaks",
            vec![],
        )
    }
    /// Protects the worksheet with explicit Excel permissions.
    pub fn protect(&self, options: &WorksheetProtectOptions<'_>) -> Result<(), ExcelComError> {
        worksheet_protect(self.dispatch_object(), options)
    }
    /// Removes worksheet protection using an optional password.
    pub fn unprotect(&self, password: Option<&str>) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        push_optional_text(&mut a, password)?;
        call(
            self.dispatch_object(),
            "excel.worksheet.unprotect",
            a.into_inner(),
        )
    }
    /// Returns whether worksheet contents are protected.
    pub fn protect_contents(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.worksheet.protectcontents")
    }
    /// Returns whether drawing objects are protected.
    pub fn protect_drawing_objects(&self) -> Result<bool, ExcelComError> {
        get_bool(
            self.dispatch_object(),
            "excel.worksheet.protectdrawingobjects",
        )
    }
    /// Returns whether scenarios are protected.
    pub fn protect_scenarios(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.worksheet.protectscenarios")
    }
    /// Returns whether worksheet protection is in effect.
    pub fn protection_mode(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.worksheet.protectionmode")
    }
    /// Shows Excel's worksheet print preview.
    pub fn print_preview(&self) -> Result<(), ExcelComError> {
        call(
            self.dispatch_object(),
            "excel.worksheet.printpreview",
            vec![],
        )
    }
    /// Sends the worksheet to Excel's `PrintOut` entry point.
    pub fn print_out(&self, options: &PrintOutOptions<'_>) -> Result<(), ExcelComError> {
        print_out(
            self.dispatch_object(),
            "excel.worksheet.printout-2361",
            options,
        )
    }
    /// Exports the worksheet through Excel's fixed-format renderer.
    pub fn export_as_fixed_format(
        &self,
        format: FixedFormatType,
        options: &FixedFormatOptions<'_>,
    ) -> Result<(), ExcelComError> {
        fixed_format(
            self.dispatch_object(),
            "excel.worksheet.exportasfixedformat-3175",
            format,
            options,
        )
    }
}

impl Range {
    /// Activates this range in Excel.
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.range.activate", vec![])
    }
    /// Selects this range in Excel.
    pub fn select(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.range.select", vec![])
    }
    /// Merges the range, optionally across rows.
    pub fn merge(&self, across: Option<bool>) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        a.push_optional(across.map(OwnedVariant::bool));
        call(self.dispatch_object(), "excel.range.merge", a.into_inner())
    }
    /// Unmerges this range.
    pub fn unmerge(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.range.unmerge", vec![])
    }
    /// Returns whether the range has merged cells, or a mixed/empty result.
    pub fn merge_cells(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.mergecells", mixed_bool)
    }
    /// Returns the containing merged range.
    pub fn merge_area(&self) -> Result<Range, ExcelComError> {
        get_range(self.dispatch_object(), "excel.range.mergearea")
    }
    /// Returns range orientation, preserving a mixed selection.
    pub fn orientation(&self) -> Result<MixedValue<i32>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.orientation", mixed_i32)
    }
    /// Sets the range orientation using Excel's raw orientation value.
    pub fn set_orientation(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.orientation",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the indent level, preserving a mixed selection.
    pub fn indent_level(&self) -> Result<MixedValue<i32>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.indentlevel", mixed_i32)
    }
    /// Sets the indent level.
    pub fn set_indent_level(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.indentlevel",
            OwnedVariant::i32(value),
        )
    }
    /// Returns whether text shrinks to fit, preserving a mixed selection.
    pub fn shrink_to_fit(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(
            self.dispatch_object(),
            "excel.range.shrinktofit",
            mixed_bool,
        )
    }
    /// Changes whether text shrinks to fit.
    pub fn set_shrink_to_fit(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.shrinktofit",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the reading-order value, preserving a mixed selection.
    pub fn reading_order(&self) -> Result<MixedValue<ReadingOrder>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.readingorder", |v| {
            mixed_i32(v).map(|m| map_mixed(m, ReadingOrder::from_raw))
        })
    }
    /// Changes the reading order.
    pub fn set_reading_order(&self, value: ReadingOrder) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.readingorder",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the locked-cell setting, preserving a mixed selection.
    pub fn locked(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(self.dispatch_object(), "excel.range.locked", mixed_bool)
    }
    /// Changes the locked-cell setting.
    pub fn set_locked(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.locked",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the hidden-formula setting, preserving a mixed selection.
    pub fn formula_hidden(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(
            self.dispatch_object(),
            "excel.range.formulahidden",
            mixed_bool,
        )
    }
    /// Changes the hidden-formula setting.
    pub fn set_formula_hidden(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.formulahidden",
            OwnedVariant::bool(value),
        )
    }
    /// Groups the range using Excel's default outline options.
    pub fn group(&self) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        for _ in 0..4 {
            a.push_optional(None);
        }
        call(self.dispatch_object(), "excel.range.group", a.into_inner())
    }
    /// Ungroups the range using Excel's default outline options.
    pub fn ungroup(&self) -> Result<(), ExcelComError> {
        let mut a = PositionalArguments::new();
        for _ in 0..4 {
            a.push_optional(None);
        }
        call(
            self.dispatch_object(),
            "excel.range.ungroup",
            a.into_inner(),
        )
    }
    /// Returns whether this outlined range's detail is displayed.
    pub fn show_detail(&self) -> Result<bool, ExcelComError> {
        get_bool(self.dispatch_object(), "excel.range.showdetail")
    }
    /// Changes whether this outlined range's detail is displayed.
    pub fn set_show_detail(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.range.showdetail",
            OwnedVariant::bool(value),
        )
    }
    /// Shows Excel's range print preview.
    pub fn print_preview(&self) -> Result<(), ExcelComError> {
        call(self.dispatch_object(), "excel.range.printpreview", vec![])
    }
    /// Sends this range to Excel's `PrintOut` entry point.
    pub fn print_out(&self, options: &PrintOutOptions<'_>) -> Result<(), ExcelComError> {
        print_out(self.dispatch_object(), "excel.range.printout-2361", options)
    }
    /// Exports this range through Excel's fixed-format renderer.
    pub fn export_as_fixed_format(
        &self,
        format: FixedFormatType,
        options: &FixedFormatOptions<'_>,
    ) -> Result<(), ExcelComError> {
        fixed_format(
            self.dispatch_object(),
            "excel.range.exportasfixedformat-3175",
            format,
            options,
        )
    }
}

fn count_i32(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
) -> Result<i32, ExcelComError> {
    i32::try_from(collection_count(target, descriptor)?).map_err(|_| ExcelComError::Unsupported {
        detail: "collection Count exceeds i32",
    })
}

fn sheet_from_dispatch(dispatch: ComPtr<Dispatch>) -> Result<Sheet, ExcelComError> {
    let object = DispatchObject {
        dispatch,
        kind: "Sheet",
    };
    let sheet_type = get_i32(&object, "excel.worksheet.type", "Sheet.Type")?;
    if sheet_type == -4167 {
        Ok(Sheet::Worksheet(Worksheet::from_dispatch(object.dispatch)))
    } else {
        Ok(Sheet::Other(SheetObject::from_dispatch(object.dispatch)))
    }
}

fn get_object<T>(
    target: &DispatchObject,
    id: &'static str,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    Ok(from(result.take_dispatch()?))
}

fn get_sheet(target: &DispatchObject, id: &'static str) -> Result<Sheet, ExcelComError> {
    let mut result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    sheet_from_dispatch(result.take_dispatch()?)
}

fn get_range(target: &DispatchObject, id: &'static str) -> Result<Range, ExcelComError> {
    get_object(target, id, Range::from_dispatch)
}
fn get_string(target: &DispatchObject, id: &'static str) -> Result<String, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?.as_string()
}
fn optional_string(
    target: &DispatchObject,
    id: &'static str,
) -> Result<Option<String>, ExcelComError> {
    let value = get_string(target, id)?;
    Ok((!value.is_empty()).then_some(value))
}
fn get_bool(target: &DispatchObject, id: &'static str) -> Result<bool, ExcelComError> {
    let result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    result.as_bool().ok_or(ExcelComError::Conversion(
        ConversionError::UnsupportedVariantType {
            vartype: result.vt(),
        },
    ))
}
fn get_i32(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<i32, ExcelComError> {
    let result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    result
        .as_i32()
        .or_else(|| {
            result
                .as_f64()
                .filter(|v| {
                    v.is_finite()
                        && v.fract() == 0.0
                        && *v >= i32::MIN as f64
                        && *v <= i32::MAX as f64
                })
                .map(|v| v as i32)
        })
        .ok_or(ExcelComError::Unsupported { detail })
}
fn get_f64(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<f64, ExcelComError> {
    let result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    result
        .as_f64()
        .or_else(|| result.as_i32().map(f64::from))
        .ok_or(ExcelComError::Unsupported { detail })
}
fn put(
    target: &DispatchObject,
    id: &'static str,
    value: OwnedVariant,
) -> Result<(), ExcelComError> {
    let _ = property_put(&target.dispatch, member(MemberId::new(id), true), value)?;
    Ok(())
}
fn call(
    target: &DispatchObject,
    id: &'static str,
    arguments: Vec<OwnedVariant>,
) -> Result<(), ExcelComError> {
    let _ = invoke(
        &target.dispatch,
        member(MemberId::new(id), false),
        arguments,
        false,
    )?;
    Ok(())
}
fn one_based(value: usize, detail: &'static str) -> Result<OwnedVariant, ExcelComError> {
    if value == 0 {
        return Err(ExcelComError::Unsupported { detail });
    }
    Ok(OwnedVariant::i32(
        i32::try_from(value).map_err(|_| ExcelComError::Unsupported { detail })?,
    ))
}
fn finite(value: f64) -> Result<(), ExcelComError> {
    if value.is_finite() {
        Ok(())
    } else {
        Err(ExcelComError::Conversion(ConversionError::NonFiniteNumber))
    }
}
fn nonnegative(value: f64) -> Result<(), ExcelComError> {
    finite(value)?;
    if value >= 0.0 {
        Ok(())
    } else {
        Err(ExcelComError::Unsupported {
            detail: "page margin must be nonnegative",
        })
    }
}
fn optional_positive(
    value: Option<usize>,
    detail: &'static str,
) -> Result<OwnedVariant, ExcelComError> {
    match value {
        Some(value) => one_based(value, detail),
        None => Ok(OwnedVariant::bool(false)),
    }
}
fn push_optional_text(
    arguments: &mut PositionalArguments,
    value: Option<&str>,
) -> Result<(), ExcelComError> {
    match value {
        Some(value) => arguments.push_result(text_bstr(value)),
        None => {
            arguments.push_optional(None);
            Ok(())
        }
    }
}
fn page_zoom_get(target: &DispatchObject, id: &'static str) -> Result<PageZoom, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    if value.as_bool() == Some(false) {
        Ok(PageZoom::Automatic)
    } else {
        value
            .as_i32()
            .or_else(|| {
                value
                    .as_f64()
                    .filter(|v| {
                        v.is_finite()
                            && v.fract() == 0.0
                            && *v >= i32::MIN as f64
                            && *v <= i32::MAX as f64
                    })
                    .map(|v| v as i32)
            })
            .map(PageZoom::Percent)
            .ok_or(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType {
                    vartype: value.vt(),
                },
            ))
    }
}
fn page_zoom_put(
    target: &DispatchObject,
    id: &'static str,
    value: PageZoom,
) -> Result<(), ExcelComError> {
    match value {
        PageZoom::Automatic => put(target, id, OwnedVariant::bool(false)),
        PageZoom::Percent(value) if value > 0 => put(target, id, OwnedVariant::i32(value)),
        PageZoom::Percent(_) => Err(ExcelComError::Unsupported {
            detail: "zoom percentage must be positive",
        }),
    }
}
fn page_fit_dimension_get(
    target: &DispatchObject,
    id: &'static str,
) -> Result<Option<usize>, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    if value.as_bool() == Some(false) {
        return Ok(None);
    }
    let value = value
        .as_i32()
        .or_else(|| {
            value
                .as_f64()
                .filter(|value| {
                    value.is_finite()
                        && value.fract() == 0.0
                        && *value >= i32::MIN as f64
                        && *value <= i32::MAX as f64
                })
                .map(|value| value as i32)
        })
        .ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: value.vt(),
            },
        ))?;
    usize::try_from(value)
        .ok()
        .filter(|value| *value > 0)
        .ok_or(ExcelComError::Unsupported {
            detail: "PageSetup fit-to-pages dimension was not positive",
        })
        .map(Some)
}
fn add_page_break<T>(
    target: &DispatchObject,
    id: &'static str,
    before: &Range,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_object(before.dispatch_object());
    let mut result = invoke(
        &target.dispatch,
        member(MemberId::new(id), false),
        a.into_inner(),
        false,
    )?;
    Ok(from(result.take_dispatch()?))
}
fn worksheet_copy_move(
    source: &Worksheet,
    id: &'static str,
    destination: SheetDestination<'_>,
) -> Result<(), ExcelComError> {
    let mut a = PositionalArguments::new();
    match destination {
        SheetDestination::Before(value) => {
            a.push_object(value.dispatch_object());
            a.push_optional(None);
        }
        SheetDestination::After(value) => {
            a.push_optional(None);
            a.push_object(value.dispatch_object());
        }
        SheetDestination::NewWorkbook => {
            a.push_optional(None);
            a.push_optional(None);
        }
    };
    call(source.dispatch_object(), id, a.into_inner())
}
fn worksheet_protect(
    target: &DispatchObject,
    options: &WorksheetProtectOptions<'_>,
) -> Result<(), ExcelComError> {
    let mut a = PositionalArguments::new();
    push_optional_text(&mut a, options.password)?;
    for value in [
        options.drawing_objects,
        options.contents,
        options.scenarios,
        options.user_interface_only,
        options.allow_formatting_cells,
        options.allow_formatting_columns,
        options.allow_formatting_rows,
        options.allow_inserting_columns,
        options.allow_inserting_rows,
        options.allow_inserting_hyperlinks,
        options.allow_deleting_columns,
        options.allow_deleting_rows,
        options.allow_sorting,
        options.allow_filtering,
        options.allow_using_pivot_tables,
    ] {
        a.push_optional(value.map(OwnedVariant::bool));
    }
    call(target, "excel.worksheet.protect-2029", a.into_inner())
}
fn print_out(
    target: &DispatchObject,
    id: &'static str,
    options: &PrintOutOptions<'_>,
) -> Result<(), ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_optional(
        options
            .from
            .map(|v| one_based(v, "PrintOut.From"))
            .transpose()?,
    );
    a.push_optional(
        options
            .to
            .map(|v| one_based(v, "PrintOut.To"))
            .transpose()?,
    );
    a.push_optional(
        options
            .copies
            .map(|v| one_based(v, "PrintOut.Copies"))
            .transpose()?,
    );
    a.push_optional(options.preview.map(OwnedVariant::bool));
    push_optional_text(&mut a, options.active_printer)?;
    a.push_optional(options.print_to_file.map(OwnedVariant::bool));
    a.push_optional(options.collate.map(OwnedVariant::bool));
    push_optional_text(&mut a, options.pr_to_file_name)?;
    a.push_optional(options.ignore_print_areas.map(OwnedVariant::bool));
    call(target, id, a.into_inner())
}
fn fixed_format(
    target: &DispatchObject,
    id: &'static str,
    format: FixedFormatType,
    options: &FixedFormatOptions<'_>,
) -> Result<(), ExcelComError> {
    let mut a = PositionalArguments::new();
    a.push_required(OwnedVariant::i32(format.raw()));
    match options.output {
        Some(path) => a.push_result(path_bstr(path))?,
        None => a.push_optional(None),
    };
    a.push_optional(options.quality.map(|v| OwnedVariant::i32(v.raw())));
    a.push_optional(options.include_doc_properties.map(OwnedVariant::bool));
    a.push_optional(options.ignore_print_areas.map(OwnedVariant::bool));
    a.push_optional(
        options
            .from
            .map(|v| one_based(v, "ExportAsFixedFormat.From"))
            .transpose()?,
    );
    a.push_optional(
        options
            .to
            .map(|v| one_based(v, "ExportAsFixedFormat.To"))
            .transpose()?,
    );
    a.push_optional(options.open_after_publish.map(OwnedVariant::bool));
    a.push_optional(None);
    call(target, id, a.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn transparent_values_preserve_unknown_excel_constants() {
        assert_eq!(WindowView::from_raw(99).raw(), 99);
        assert_eq!(AutomationSecurity::FORCE_DISABLE.raw(), 3);
        assert_eq!(PageBreakType::MANUAL.raw(), -4135);
    }

    #[test]
    fn print_out_preserves_all_optional_positions() {
        let options = PrintOutOptions {
            copies: Some(2),
            collate: Some(true),
            ..Default::default()
        };
        let mut values = PositionalArguments::new();
        values.push_optional(
            options
                .from
                .map(|v| one_based(v, "PrintOut.From"))
                .transpose()
                .expect("valid"),
        );
        values.push_optional(
            options
                .to
                .map(|v| one_based(v, "PrintOut.To"))
                .transpose()
                .expect("valid"),
        );
        values.push_optional(
            options
                .copies
                .map(|v| one_based(v, "PrintOut.Copies"))
                .transpose()
                .expect("valid"),
        );
        values.push_optional(options.preview.map(OwnedVariant::bool));
        push_optional_text(&mut values, options.active_printer).expect("text");
        values.push_optional(options.print_to_file.map(OwnedVariant::bool));
        values.push_optional(options.collate.map(OwnedVariant::bool));
        push_optional_text(&mut values, options.pr_to_file_name).expect("text");
        values.push_optional(options.ignore_print_areas.map(OwnedVariant::bool));
        let values = values.into_inner();
        assert_eq!(values.len(), 9);
        assert_eq!(values[0].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[2].as_i32(), Some(2));
        assert_eq!(values[6].as_bool(), Some(true));
        assert_eq!(values[8].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn secret_debug_fields_are_redacted() {
        let worksheet = format!(
            "{:?}",
            WorksheetProtectOptions {
                password: Some("secret"),
                ..Default::default()
            }
        );
        let workbook = format!(
            "{:?}",
            WorkbookProtectOptions {
                password: Some("secret"),
                ..Default::default()
            }
        );
        let print = format!(
            "{:?}",
            PrintOutOptions {
                active_printer: Some("private printer"),
                pr_to_file_name: Some("private path"),
                ..Default::default()
            }
        );
        for text in [worksheet, workbook, print] {
            assert!(text.contains("REDACTED"));
            assert!(!text.contains("secret"));
            assert!(!text.contains("private"));
        }
    }
}
