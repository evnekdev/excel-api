#![doc = "Safe building blocks for Rust-native Microsoft Excel XLL add-ins."]

pub mod borrowed;
pub mod context;
pub mod convert;
pub mod error;
pub mod registration;
pub mod value;

pub use borrowed::{
    DecodeError, ExcelArrayColumns, ExcelArrayElements, ExcelArrayRows, ExcelArrayView,
    ExcelMissing, ExcelMultiReference, ExcelNil, ExcelReference, ExcelReferenceArea,
    ExcelReferenceAreas, ExcelSingleReference, ExcelStr, ExcelValueRef, RawExcelValue,
};
pub use context::{MacroContext, ThreadSafeContext, WorksheetContext};
pub use convert::{ConversionLimits, FromExcel, IntoExcel};
pub use error::{ConversionError, ExcelError, OwnedValueError, Utf16ConversionError};
pub use registration::{AddInDescriptor, FunctionFlags, FunctionRegistration, RegistrationError};
pub use value::{ExcelArray, ExcelArrayColumn, ExcelString, ExcelValue, OptionalValue};

#[cfg(feature = "macros")]
pub use excel_api_macros::{excel_command, excel_function};

/// Common imports for XLL authors.
pub mod prelude {
    pub use crate::{
        AddInDescriptor, ConversionLimits, ExcelArray, ExcelError, ExcelString, ExcelValue,
        ExcelValueRef, FromExcel, FunctionFlags, FunctionRegistration, IntoExcel, OptionalValue,
    };

    #[cfg(feature = "macros")]
    pub use excel_api_macros::{excel_command, excel_function};
}
