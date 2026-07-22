//! Experimental, unpublished building blocks for Excel COM Automation.
//!
//! The API is not stable. Semantic and wrapper types may change before the
//! first release. The crate deliberately exposes no raw COM pointers.

#![cfg(windows)]

mod automation;
mod error;
mod excel;
mod internal;
mod object_model;

pub use automation::{
    AutomationArgument, AutomationArray, AutomationValue, ConversionPolicy, Currency, ExcelError,
    OaDate,
};
pub use error::ExcelComError;
pub use excel::{Application, Workbook, Workbooks};
pub use internal::ComApartment;
pub use object_model::{
    DocumentationStatus, IMPLEMENTED_MEMBER_IDS, ImplementationStatus, MemberId, ObjectId,
    TestStatus,
};
