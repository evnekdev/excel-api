//! Experimental, unpublished building blocks for Excel COM Automation.
//!
//! # Scope and platform
//!
//! This Windows-only crate implements the narrow path
//! `Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range`.
//! It starts a local Excel Automation server; it does not attach to an
//! existing session, marshal interfaces, expose raw COM pointers, or offer a
//! generic collection, formatting, chart, macro, or event API. The public API
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
//! # Example
//!
//! ```no_run
//! use excel_com::{
//!     Application, AutomationArgument, AutomationValue, ComApartment,
//!     WorksheetsAddOptions,
//! };
//!
//! # fn main() -> Result<(), excel_com::ExcelComError> {
//! let apartment = ComApartment::sta()?;
//! let application = Application::new(&apartment)?;
//! let workbook = application.workbooks()?.add()?;
//! let worksheet = workbook.worksheets()?.add(WorksheetsAddOptions::new())?;
//! let cell = worksheet.range(
//!     AutomationArgument::Value(AutomationValue::Text("A1".to_owned())),
//!     None,
//! )?;
//! cell.set_value2(AutomationValue::Number(42.0))?;
//! application.quit()?;
//! # Ok(())
//! # }
//! ```

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
    Application, Range, Workbook, Workbooks, Worksheet, Worksheets, WorksheetsAddOptions,
    XlSheetVisibility,
};
pub use internal::ComApartment;
pub use object_model::{
    DocumentationStatus, IMPLEMENTED_MEMBER_IDS, ImplementationStatus, MemberId, ObjectId,
    TestStatus,
};
