#![doc = "Safe building blocks for Rust-native Microsoft Excel XLL add-ins."]

pub mod borrowed;
pub mod context;
pub mod convert;
pub mod error;
mod excel_call;
mod excel_owned;
pub mod metadata;
pub mod registration;
mod return_alloc;
pub mod return_plan;
pub mod runtime;
#[doc(hidden)]
pub mod thunk;
pub mod value;

pub use borrowed::{
    DecodeError, ExcelArrayColumns, ExcelArrayElements, ExcelArrayRows, ExcelArrayView,
    ExcelMissing, ExcelMultiReference, ExcelNil, ExcelReference, ExcelReferenceArea,
    ExcelReferenceAreas, ExcelSingleReference, ExcelStr, ExcelValueRef, RawExcelValue,
};
pub use context::{LifecycleContext, MacroContext, ThreadSafeContext, WorksheetContext};
pub use convert::{ConversionLimits, FromExcel, IntoExcel};
pub use error::{
    ConversionError, ExcelError, OwnedValueError, ReturnError, ReturnMaterializationError,
    ThunkError, Utf16ConversionError,
};
pub use excel_call::{
    AbortCheckMode, CallPermission, CoerceTarget, ExcelCallDescriptor, ExcelCallError,
    ExcelReturnCode, ResultRoot, SdkExcel12vBackend, XL_ABORT, XL_COERCE, XL_FREE, XL_GET_NAME,
    XL_SHEET_ID, XL_SHEET_NM, XLF_CALLER, XLF_REGISTER, XLF_SET_NAME, XLF_UNREGISTER,
};
pub use excel_owned::{
    ExcelOwnedConversionError, ExcelOwnedValue, ExcelReleaseError, ExcelReleasePolicy,
    ExcelXlFreeTransfer,
};
pub use metadata::{CountedUtf16Arg, ExcelReferenceArg, ExcelValueArg, NullTerminatedUtf16Arg};
pub use registration::{
    AddInDescriptor, ExcelArgumentType, ExcelReturnType, FunctionFlags, FunctionRegistration,
    FunctionSignature, RegistrationError,
};
pub use return_alloc::{ExcelReturn, xl_auto_free12};
pub use return_plan::{
    ExcelReturnArray, ExcelReturnValue, PlannedArray, PlannedArrayElement, PlannedText,
    PlannedValue, ReturnLimits, ReturnOwnershipStrategy, ReturnPlan, ReturnStorageTotals,
    ReturnText,
};
pub use runtime::{LifecycleError, LifecycleOutcome, Runtime, RuntimeDiagnostics, RuntimePhase};
pub use value::{ExcelArray, ExcelArrayColumn, ExcelString, ExcelValue, OptionalValue};

#[cfg(feature = "macros")]
pub use excel_api_macros::{excel_command, excel_function};

/// Common imports for XLL authors.
pub mod prelude {
    pub use crate::{
        AddInDescriptor, ConversionLimits, CountedUtf16Arg, ExcelArray, ExcelError,
        ExcelReferenceArg, ExcelReturnArray, ExcelReturnValue, ExcelString, ExcelValue,
        ExcelValueArg, ExcelValueRef, FromExcel, FunctionFlags, FunctionRegistration, IntoExcel,
        NullTerminatedUtf16Arg, OptionalValue, ReturnError, ReturnLimits,
        ReturnMaterializationError, ReturnPlan, ReturnText,
    };

    #[cfg(feature = "macros")]
    pub use excel_api_macros::{excel_command, excel_function};
}
