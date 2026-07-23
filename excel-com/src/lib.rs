//! Experimental, unpublished building blocks for Excel COM Automation.
//!
//! # Scope and platform
//!
//! This Windows-only crate implements the narrow path
//! `Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range`.
//! It starts a local Excel Automation server; it does not attach to an
//! existing session, marshal interfaces, expose raw COM pointers, or offer a
//! generic collection, chart, macro, or event API. The public API
//! remains experimental and may change before a first release.
//!
//! # Apartments and lifetime
//!
//! Create [`ComApartment::sta`] on the calling thread, retain its guard for
//! longer than every wrapper created through it, and create [`Application`]
//! with that guard. The guard and wrappers are not `Send` or `Sync`; this crate
//! does not marshal COM interfaces between threads. Releasing a wrapper only
//! releases its COM reference. Call [`Application::quit`] explicitly to stop a
//! crate-created Excel server.
//!
//! # Values and Range shape
//!
//! [`AutomationValue`] preserves empty, null, bool, number, text, error,
//! OLE Automation date, currency, and rank-two rectangular-array distinctions
//! without exposing `VARIANT` or `SAFEARRAY`. [`Range::value`] uses `VT_DATE`
//! for dates while [`Range::value2`] reads and writes their numeric serials.
//! [`Range::formula`] and [`Range::formula2`] are separate Excel members.
//! A scalar may be written only to a 1x1 range, and an [`AutomationArray`] must
//! have exactly the target dimensions; otherwise no COM setter is called and
//! [`ExcelComError::Conversion`] contains [`ConversionError::ShapeMismatch`].
//!
//! Excel errors retain their complete signed SCODE through [`ExcelError`].
//! [`OaDate`] retains a finite serial, [`Currency`] retains its exact scaled
//! `CY` integer, and [`AutomationArgument::Missing`] encodes the Automation
//! `VT_ERROR` / `DISP_E_PARAMNOTFOUND` optional-argument marker. Unsupported
//! features are intentionally reported rather than guessed.
//!
//! # Workbook file lifecycle
//!
//! [`Workbooks::open`] and [`Workbook::save_as`] retain every declared
//! optional position as an explicit [`AutomationArgument::Missing`] marker;
//! the private dispatch layer performs the one required COM-order reversal.
//! Their typed options redact passwords in `Debug`. Input paths use Windows
//! [`std::path::Path`] / `OsStr` UTF-16 units directly, with no
//! canonicalization or lossy conversion; an embedded NUL returns
//! [`ExcelComError::InvalidPath`]. [`Application::display_alerts_guard`]
//! restores the previous `DisplayAlerts` value when dropped (or explicitly
//! through [`DisplayAlertsGuard::restore`]). [`Workbook::close`] consumes its
//! wrapper and takes [`WorkbookCloseOptions`] with [`SaveChanges`].
//!
//! # Collections and navigation
//!
//! [`Workbooks::iter`], [`Worksheets::iter`], and [`Areas::iter`] are
//! fallible, single-pass, apartment-bound cursors over Excel's `_NewEnum`.
//! Each yields `Result`; a COM or conversion failure fuses that cursor, and
//! dropping it early releases its enumerator. [`Workbook::is_same_object`],
//! [`Worksheet::is_same_object`], and [`Range::is_same_object`] compare the
//! canonical COM `IUnknown` identity and are deliberately fallible rather
//! than `PartialEq`. Excel may materialize logically equivalent Range values
//! as distinct COM objects, so addresses are not identity.
//!
//! # Selecting, naming, converting, and evaluating ranges
//!
//! A1 is the concise default: [`Worksheet::range`] selects one A1 reference,
//! while [`Worksheet::range_between`] accepts two corners. R1C1 is explicit
//! through [`Worksheet::range_r1c1`], which uses Excel's own
//! [`Application::convert_formula`] engine rather than a Rust parser. Numeric
//! [`Worksheet::cell`] and [`Worksheet::range_from_cells`] coordinates are
//! one-based and reject zero before COM. [`Range::address_a1`],
//! [`Range::address_r1c1`], and [`Range::address_with_options`] request
//! explicit output notation without relying on the global setting.
//!
//! [`ReferenceStyle`] preserves unknown Excel values. Because it is global
//! Application state, use [`Application::reference_style_guard`] for a
//! temporary change and explicitly call [`ReferenceStyleGuard::restore`] when
//! a restoration failure must be observed. [`FormulaConversionOptions`] and
//! [`RangeAddressOptions`] preserve omitted Automation positions as Missing.
//!
//! [`Workbook::names`] and [`Worksheet::names`] expose distinct Excel scopes.
//! [`Names::add`] accepts a Range, A1, R1C1, or formula target, but a valid
//! [`Name`] can still denote a scalar, formula, or external reference instead
//! of a Range. [`Name::range`] is therefore fallible. Evaluation is similarly
//! explicit: [`Application::evaluate_value`] requires an Automation value or
//! array, whereas [`Application::evaluate_range`] requires a Range object.
//! Equal addresses, names, and formula strings do not imply canonical COM
//! object identity.
//!
//! # Formula, calculation, and auditing
//!
//! [`Range::formula`] and [`Range::formula2`] return [`FormulaValue`]: scalar
//! text, an exact rectangular [`AutomationArray`], `Mixed`, or `Empty`. Scalar
//! formula setters require a 1x1 receiver; rectangular setters require exactly
//! matching dimensions before COM. Excel remains the formula parser and
//! calculation engine—this crate neither parses formulas nor recalculates them
//! in Rust. [`Range::formula2`] and [`Range::set_formula2`] preserve Excel's
//! dynamic-array behavior, while [`Range::formula_array`] and
//! [`Range::set_formula_array`] expose Excel's legacy array-formula member.
//! Spill and legacy-array restrictions remain Excel-owned and are returned as
//! structured Automation failures where Excel reports one.
//!
//! [`Application::calculation_mode_guard`] temporarily changes the process-wide
//! calculation mode and restores it on drop. [`Application::calculate`],
//! [`Application::calculate_full`], [`Application::calculate_full_rebuild`],
//! [`Worksheet::calculate`], and [`Range::calculate`] delegate directly to
//! Excel. [`Range::mark_dirty`] asks Excel to mark a range for recalculation.
//! Do not rely on a calculation-state snapshot as a general asynchronous
//! completion guarantee.
//!
//! [`Range::special_cells`], [`Range::find`], [`Range::find_all`], and
//! [`Range::replace`] supply typed Excel-backed discovery. Find defaults send
//! every remembered search option explicitly; [`RangeFindIter`] detects
//! wraparound from normalized external addresses instead of COM identity.
//!
//! ```no_run
//! # fn example(application: &excel_com::Application, range: &excel_com::Range) -> Result<(), excel_com::ExcelComError> {
//! use excel_com::{
//!     AutomationValue, CalculationMode, FindOptions, FormulaValue, SpecialCellType,
//! };
//! range.set_formula("=SUM(A1:A10)")?;
//! assert!(matches!(range.formula()?, FormulaValue::Text(_)));
//! range.set_formula2("=SEQUENCE(2,2)")?;
//! let spill = range.spilling_to_range()?;
//! let guard = application.calculation_mode_guard(CalculationMode::MANUAL)?;
//! spill.calculate()?;
//! guard.restore()?;
//! let formulas = spill.special_cells(SpecialCellType::FORMULAS, None)?;
//! let matches = formulas
//!     .find_all(&AutomationValue::Text("SUM".to_owned()), &FindOptions::default())?
//!     .collect::<Result<Vec<_>, _>>()?;
//! # drop(matches);
//! # Ok(())
//! # }
//! ```
//!
//! # Tables, filters, sorting, and structural edits
//!
//! [`Worksheet::list_objects`] exposes typed [`ListObjects`], [`ListObject`],
//! [`ListColumns`], and [`ListRows`] wrappers. Table indexes and filter fields
//! are one-based; `DataBodyRange` is explicitly optional for an empty table.
//! Excel owns table naming, structured-reference parsing, calculated-column
//! propagation, resizing, totals, and all data changes. A table's
//! [`ListObject::unlist`] and deletion consume their wrapper because the COM
//! target no longer represents a table afterward.
//!
//! [`Range::apply_auto_filter`] and [`AutoFilter`] retain Excel's stateful
//! filtering model. [`Range::sort`] and [`Sort::apply`] modify worksheet data
//! in place. [`Range::validation`] passes validation formulas through as Excel
//! syntax. [`Range::remove_duplicates`], [`Range::insert`], and
//! [`Range::delete`] are structural, in-place operations; [`Range::copy`],
//! [`Range::cut`], and [`Range::paste_special`] use Excel's cut/copy state and
//! do not inspect the operating-system clipboard. Hidden assignment is passed
//! to Excel exactly as supplied; use complete rows or columns where required.
//!
//! ```no_run
//! # fn example(worksheet: &excel_com::Worksheet) -> Result<(), excel_com::ExcelComError> {
//! use excel_com::{ListObjectAddOptions, TableHeaderMode, TotalsCalculation};
//! let source = worksheet.range("A1:C10")?;
//! let table = worksheet.list_objects()?.add_from_range(&ListObjectAddOptions {
//!     source: &source,
//!     has_headers: TableHeaderMode::YES,
//!     destination: None,
//!     table_style_name: None,
//! })?;
//! let amount = table.list_columns()?.add(None)?;
//! amount.set_name("Amount")?;
//! amount.set_calculated_column_formula("=[@Quantity]*[@Price]")?;
//! table.set_show_totals(true)?;
//! amount.set_totals_calculation(TotalsCalculation::SUM)?;
//! # Ok(())
//! # }
//! ```
//!
//! # Formatting ranges
//!
//! Formatting wrappers preserve Excel's mixed-selection results: a getter that
//! sees a concrete common value returns [`MixedValue::Uniform`], whereas an
//! Excel `Null` result becomes [`MixedValue::Mixed`] and `Empty` becomes
//! [`MixedValue::Empty`]. RGB values use [`ExcelColor`] and Excel's low-byte
//! red, then green, then blue integer order. Every formatting wrapper remains
//! apartment-bound and is neither `Send` nor `Sync`.
//!
//! | Need | API |
//! |---|---|
//! | Font | [`Range::font`] |
//! | Fill | [`Range::interior`] |
//! | Borders | [`Range::borders`] |
//! | Number format | [`Range::set_number_format`] |
//! | Horizontal alignment | [`Range::set_horizontal_alignment`] |
//! | Vertical alignment | [`Range::set_vertical_alignment`] |
//! | Wrap text | [`Range::set_wrap_text`] |
//! | Row height | [`Range::set_row_height`] |
//! | Column width | [`Range::set_column_width`] |
//! | AutoFit | [`Range::auto_fit`] |
//!
//! ```no_run
//! # fn example(range: &excel_com::Range) -> Result<(), excel_com::ExcelComError> {
//! use excel_com::{
//!     BorderIndex, BorderLineStyle, BorderWeight, ExcelColor, HorizontalAlignment, MixedValue,
//! };
//! let font = range.font()?;
//! font.set_bold(true)?;
//! font.set_size(12.0)?;
//! font.set_color(ExcelColor::from_rgb(20, 40, 180))?;
//! range.interior()?.set_color(ExcelColor::from_rgb(240, 240, 200))?;
//! range.set_number_format("0.00")?;
//! range.set_horizontal_alignment(HorizontalAlignment::CENTER)?;
//! let bottom = range.borders()?.item(BorderIndex::EDGE_BOTTOM)?;
//! bottom.set_line_style(BorderLineStyle::CONTINUOUS)?;
//! bottom.set_weight(BorderWeight::THIN)?;
//! range.entire_column()?.auto_fit()?;
//! match range.font()?.bold()? {
//!     MixedValue::Uniform(value) => println!("{value}"),
//!     MixedValue::Mixed => println!("selection uses mixed bold formatting"),
//!     MixedValue::Empty => println!("no concrete value"),
//! }
//! # Ok(())
//! # }
//! ```
//!
//! | Need | API |
//! |---|---|
//! | A1 selection | [`Worksheet::range`] |
//! | Two A1 corners | [`Worksheet::range_between`] |
//! | R1C1 selection | [`Worksheet::range_r1c1`] |
//! | Numeric cell | [`Worksheet::cell`] |
//! | Numeric rectangle | [`Worksheet::range_from_cells`] |
//! | A1 output | [`Range::address_a1`] |
//! | R1C1 output | [`Range::address_r1c1`] |
//! | Customized address | [`Range::address_with_options`] |
//! | Formula conversion | [`Application::convert_formula`] |
//! | Workbook names | [`Workbook::names`] |
//! | Worksheet names | [`Worksheet::names`] |
//! | Add name | [`Names::add`] |
//! | Resolve name | [`Name::range`] |
//! | Evaluate scalar | [`Application::evaluate_value`] |
//! | Evaluate Range | [`Application::evaluate_range`] |
//!
//! ```no_run
//! # fn example() -> Result<(), excel_com::ExcelComError> {
//! # use excel_com::{Application, ComApartment};
//! # let apartment = ComApartment::sta()?;
//! # let application = Application::new(&apartment)?;
//! # let workbook = application.workbooks()?.add()?;
//! let worksheets = workbook.worksheets()?;
//! let direct = worksheets.item_by_index(1)?;
//! let first = worksheets.iter()?.next().transpose()?.expect("worksheet");
//! assert!(direct.is_same_object(&first)?);
//! let base = direct.range("B2")?;
//! let shifted = base.resize(2, 3)?.cell(2, 3)?.offset(1, -1)?;
//! # drop(shifted);
//! # application.quit()?;
//! # Ok(())
//! # }
//! ```
//!
//! # Example
//!
//! ```no_run
//! use excel_com::{
//!     Application, ComApartment,
//!     SaveChanges, WorkbookCloseOptions, WorkbookOpenOptions,
//!     WorkbookSaveAsOptions, XlFileFormat,
//! };
//! use std::path::Path;
//!
//! # fn main() -> Result<(), excel_com::ExcelComError> {
//! let apartment = ComApartment::sta()?;
//! let application = Application::new(&apartment)?;
//! let workbooks = application.workbooks()?;
//! let alerts = application.display_alerts_guard(false)?;
//! let workbook = workbooks.open(
//!     Path::new("read-only.xlsx"),
//!     WorkbookOpenOptions { read_only: Some(true), ..WorkbookOpenOptions::new() },
//! )?;
//! workbook.save_copy_as(Path::new("backup.xlsx"))?;
//! workbook.close(WorkbookCloseOptions {
//!     save_changes: SaveChanges::Discard,
//!     ..WorkbookCloseOptions::new()
//! })?;
//! alerts.restore()?;
//! application.quit()?;
//! # Ok(())
//! # }
//! ```
//!
//! To save a workbook as `.xlsx`, use
//! `WorkbookSaveAsOptions { file_format: Some(XlFileFormat::OPEN_XML_WORKBOOK),
//! ..WorkbookSaveAsOptions::new() }` with [`Workbook::save_as`].
//!
//! # Charts, drawings, pictures, and sparklines
//!
//! [`Worksheet::chart_objects`] creates embedded charts with explicit
//! point-based [`ChartBounds`], while [`Workbook::charts`] exposes chart
//! sheets. [`Chart`] offers source-data, series, axes, labels, trendlines,
//! error bars, titles, legend, and installed-filter export operations. Excel
//! interprets source data and owns chart calculation; this crate does not
//! implement a chart renderer.
//!
//! [`Worksheet::shapes`] supports a bounded drawing surface: AutoShapes,
//! lines, local file-backed pictures, text boxes, ordering, placement, and
//! deletion. Shape grouping is deliberately unavailable. `Range::copy_picture`
//! uses Excel's own copy state rather than reading the operating-system
//! clipboard, and [`Application::cut_copy_mode`] exposes only Excel's
//! enum-like state. [`Range::sparkline_groups`] creates line, column, and
//! win/loss sparkline groups from Excel ranges.
//!
//! | Need | API |
//! |---|---|
//! | Embedded chart | [`Worksheet::add_chart`] / [`Worksheet::chart_objects`] |
//! | Chart sheet | [`Workbook::charts`] |
//! | Series and axes | [`Chart::series_collection`] / [`Chart::axes`] |
//! | Shapes and pictures | [`Worksheet::shapes`] |
//! | Excel-native picture copy | [`Range::copy_picture`] |
//! | Sparklines | [`Range::sparkline_groups`] |
//!
//! These wrappers are visible opt-in integration-test surfaces because their
//! result depends on the installed Excel version and its rendering stack. They
//! require an explicitly created local application and should be followed by
//! [`Application::quit`]; they never attach to an existing Excel session.
//!
//! # Conditional formatting, styles, comments, and hyperlinks
//!
//! The advanced-presentation surface remains apartment-bound and is neither
//! `Send` nor `Sync`. Excel evaluates conditional formatting; this crate does
//! not parse formulas or reproduce its rule engine. Formula text is passed to
//! Excel unchanged (apart from embedded-NUL rejection), and relative formulas
//! are interpreted by Excel relative to the upper-left cell of `AppliesTo`.
//! `FormatConditions` enumerates Excel's current rule order, which is expected
//! but not promised to be priority order until a controlled host observation
//! verifies it. `StopIfTrue` is an Excel rule property, not a Rust evaluator.
//!
//! | Need | API |
//! |---|---|
//! | Conditional rules | [`Range::format_conditions`] and [`FormatConditions`] |
//! | Effective display result | [`Range::display_format`] and [`DisplayFormat`] |
//! | Reusable workbook style | [`Workbook::styles`] and [`Styles`] |
//! | Assign a style | [`Range::set_style_by_name`] or [`Range::set_style`] |
//! | Theme colour and tint | [`ThemeColor`], Font, Interior, Border, and Tab methods |
//! | Legacy Note | [`Range::add_comment`] and [`Comment`] |
//! | Read-only threaded comments | [`Range::threaded_comment`] |
//! | Internal/external link representation | [`Hyperlinks::add`] and [`Hyperlink`] |
//!
//! [`DisplayFormat`] exposes Excel's displayed formatting after conditional
//! evaluation; it is intentionally read-only and is not interchangeable with
//! ordinary Range formatting. A [`Style`] has independent inclusion flags and
//! may be built-in or custom; Excel controls which built-in styles may be
//! deleted. Direct [`ExcelColor`] and [`ThemeColor`] assignments are both
//! passed to Excel. Their interaction, including tint and shade, remains
//! Excel-defined; every tint setter validates the inclusive `-1.0..=1.0`
//! range before COM.
//!
//! Modern Excel calls legacy [`Comment`] objects *Notes*. They are distinct
//! from account-dependent threaded comments, for which this crate provides
//! inspection only. Hyperlink `Address` and `SubAddress` are independent:
//! an internal link normally uses `SubAddress` without `Address`. The API
//! rejects embedded NUL text, does not validate a URL, and never follows a
//! hyperlink. Visible live acceptance tests are opt-in; see the normalized
//! evidence for their current runtime status.

#![cfg(windows)]
#![deny(missing_docs)]
#![deny(rustdoc::broken_intra_doc_links)]
#![warn(rustdoc::private_intra_doc_links)]

mod automation;
mod error;
mod excel;
mod internal;
mod object_model;

pub use automation::{
    AutomationArgument, AutomationArray, AutomationValue, ConversionError, ConversionPolicy,
    ComCallDisposition, ComMessageFilterGuard, ComRetryPolicy, Currency, ExcelError,
    InvocationRetrySafety, OaDate,
};
pub use error::{ExcelComError, ExcelRuntimeError};
pub use excel::{
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
pub use excel::{
    AdvancedFilterAction, AdvancedFilterOptions, AggregationFunction, Application, Areas,
    AreasIter, AskToUpdateLinksGuard, AutoFillType, AutoFilter, AutoFilterOperator,
    AutoFilterOptions, AutomationSecurity, AutomationSecurityGuard, Border, BorderIndex,
    BorderLineStyle, BorderWeight, Borders, BordersIter, CalculationMode, CalculationModeGuard,
    CalculationState, ConsolidateOptions, ConsolidationSource, DataSeriesDateUnit,
    DataSeriesOptions, DataSeriesType, DataTableInputs, DeleteShiftDirection, DisplayAlertsGuard,
    DynamicFilterCriteria, ExcelColor, ExcelColorIndex, ExternalLinkSource, FillPattern, Filter,
    FilterCriterion, Filters, FiltersIter, FindLookIn, FindMatchMode, FindOptions,
    FixedFormatOptions, FixedFormatQuality, FixedFormatType, Font, FormulaConversionOptions,
    FormulaValue, GoalSeekOptions, HPageBreak, HPageBreaks, HorizontalAlignment,
    InsertFormatOrigin, InsertShiftDirection, Interior, LinkStatus, LinkType, ListColumn,
    ListColumns, ListColumnsIter, ListObject, ListObjectAddOptions, ListObjectSourceType,
    ListObjects, ListObjectsIter, ListRow, ListRows, ListRowsIter, MixedValue, Name,
    NameAddOptions, NameRefersTo, Names, NamesIter, OpenTextOptions, Outline, PageBreakType,
    PageFit, PageOrientation, PageSetup, PageZoom, PaperSize, PasteOperation, PasteSpecialOptions,
    PasteType, PrintErrors, PrintLocation, PrintOrder, PrintOutOptions, Range, RangeAddressOptions,
    RangeFindIter, RangeInsertOptions, RangeSortOptions, ReadingOrder, ReferenceAbsoluteMode,
    ReferenceStyle, ReferenceStyleGuard, RemoveDuplicatesOptions, ReplaceOptions,
    SafeWorkbookOpenOptions, SaveChanges, Scenario, ScenarioAddOptions, ScenarioReportType,
    ScenarioSummaryOptions, Scenarios, ScenariosIter, SearchDirection, SearchOrder,
    SeriesOrientation, Sheet, SheetDestination, SheetObject, SheetType, SheetView, SheetVisibility,
    Sheets, SheetsIter, Sort, SortDataOption, SortField, SortFields, SortMethod, SortOrder,
    SortOrientation, SpecialCellType, SpecialCellValueMask, SummaryColumn, SummaryRow, Tab,
    TableHeaderMode, TextColumnSpec, TextColumnType, TextDelimiter, TextExportOptions,
    TextFileFormat, TextParsingType, TextPlatform, TextQualifier, TextToColumnsOptions,
    TotalsCalculation, UnderlineStyle, VPageBreak, VPageBreaks, Validation, ValidationAddOptions,
    ValidationAlertStyle, ValidationOperator, ValidationType, VerticalAlignment, Window,
    WindowView, Windows, Workbook, WorkbookCloseOptions, WorkbookOpenFormat, WorkbookOpenOptions,
    WorkbookProtectOptions, WorkbookSaveAsOptions, Workbooks, WorkbooksIter, Worksheet,
    WorksheetAddOptions, WorksheetProtectOptions, Worksheets, WorksheetsAddOptions, WorksheetsIter,
    XlCorruptLoad, XlFileFormat, XlPlatform, XlSaveAsAccessMode, XlSaveConflictResolution,
    XlSheetVisibility, XlUpdateLinks,
};
pub use excel::{
    AttachOptions, AttachedApplication, ExcelSession, ExcelSessionDiagnostics,
    ExistingInstanceSelection, OfficeBitness, OwnedApplication, ProcessExitReport,
    SessionOwnership,
};
pub use excel::{
    AutoShapeType, Axes, Axis, AxisGroup, AxisScaleType, AxisTitle, AxisType, Chart, ChartArea,
    ChartBounds, ChartCreateOptions, ChartExportOptions, ChartFormat, ChartObject, ChartObjects,
    ChartObjectsIter, ChartSheet, ChartSheetDestination, ChartTitle, ChartType, Charts, ChartsIter,
    CopyPictureFormat, CopyPictureOptions, CutCopyMode, DataLabel, DataLabelOptions, DataLabelType,
    DataLabels, ErrorBarDirection, ErrorBarInclude, ErrorBarOptions, ErrorBarType, FillFormat,
    Legend, LegendPosition, LineFormat, MarkerStyle, PictureAddOptions, PictureAppearance,
    PlotArea, PlotBy, Series, SeriesAddOptions, SeriesCollection, SeriesCollectionIter, SeriesData,
    Shape, ShapeBounds, ShapePlacement, ShapePoint, ShapeType, Shapes, ShapesIter, SparkScale,
    SparklineGroup, SparklineGroups, SparklineGroupsIter, SparklineType, TextBoxAddOptions,
    TextFrame, TextOrientation, TextRange, TickLabelPosition, TickLabels, TickMark, Trendline,
    TrendlineAddOptions, TrendlineType, Trendlines, TrendlinesIter, ZOrderCommand,
};
pub use excel::{
    CommandType, ConnectionDetails, ConnectionFileType, ConnectionType, Connections,
    ConnectionsIter, CredentialsMethod, OdbcConnection, OdbcConnectionAddOptions, OleDbConnection,
    OleDbConnectionAddOptions, QueryTable, QueryTables, QueryTablesIter, RefreshCancellationReport,
    RefreshWaitOptions, RefreshWaitReport, SecretStringValue, TextConnection, TextQueryAddOptions,
    WebConnection, WorkbookConnection, WorkbookQueries, WorkbookQueriesIter, WorkbookQuery,
};
pub use excel::{
    MissingItemsLimit, PivotCache, PivotCaches, PivotCachesIter, PivotDataField, PivotField,
    PivotFieldOrientation, PivotFieldPlacement, PivotFields, PivotFieldsIter, PivotFilter,
    PivotFilterType, PivotFilters, PivotFiltersIter, PivotItem, PivotItems, PivotItemsIter,
    PivotLabelFilterOptions, PivotLayoutOptions, PivotSourceType, PivotTable,
    PivotTableCreateOptions, PivotTableVersion, PivotTables, PivotTablesIter,
    PivotValueFilterOptions, Slicer, SlicerCache, SlicerCaches, SlicerCachesIter, Slicers,
    SlicersIter,
};
pub use internal::ComApartment;
pub use object_model::{
    DocumentationStatus, IMPLEMENTED_MEMBER_IDS, ImplementationStatus, MemberId, ObjectId,
    TestStatus,
};
