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
    Currency, ExcelError, OaDate,
};
pub use error::ExcelComError;
pub use excel::{
    Application, Areas, AreasIter, Border, BorderIndex, BorderLineStyle, BorderWeight, Borders,
    BordersIter, CalculationMode, CalculationModeGuard, CalculationState, DisplayAlertsGuard,
    ExcelColor, ExcelColorIndex, FillPattern, FindLookIn, FindMatchMode, FindOptions, Font,
    FormulaConversionOptions, FormulaValue, HorizontalAlignment, Interior, MixedValue, Name,
    NameAddOptions, NameRefersTo, Names, NamesIter, Range, RangeAddressOptions, RangeFindIter,
    ReferenceAbsoluteMode, ReferenceStyle, ReferenceStyleGuard, ReplaceOptions, SaveChanges,
    SearchDirection, SearchOrder, SpecialCellType, SpecialCellValueMask, UnderlineStyle,
    VerticalAlignment, Workbook, WorkbookCloseOptions, WorkbookOpenFormat, WorkbookOpenOptions,
    WorkbookSaveAsOptions, Workbooks, WorkbooksIter, Worksheet, Worksheets, WorksheetsAddOptions,
    WorksheetsIter, XlCorruptLoad, XlFileFormat, XlPlatform, XlSaveAsAccessMode,
    XlSaveConflictResolution, XlSheetVisibility, XlUpdateLinks,
};
pub use internal::ComApartment;
pub use object_model::{
    DocumentationStatus, IMPLEMENTED_MEMBER_IDS, ImplementationStatus, MemberId, ObjectId,
    TestStatus,
};
