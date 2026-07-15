#![doc = "Safe building blocks for Rust-native Microsoft Excel XLL add-ins."]

pub mod async_udf;
pub mod borrowed;
pub mod context;
pub mod convert;
pub mod diagnostics;
pub mod dispatcher;
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

pub use async_udf::{
    AsyncCancellationToken, AsyncCompletionError, AsyncExecuteError, AsyncExecutor, AsyncJob,
    AsyncSubmitError, ThreadPoolExecutor,
};
pub use borrowed::{
    DecodeError, ExcelArrayColumns, ExcelArrayElements, ExcelArrayRows, ExcelArrayView,
    ExcelMissing, ExcelMultiReference, ExcelNil, ExcelReference, ExcelReferenceArea,
    ExcelReferenceAreas, ExcelSingleReference, ExcelStr, ExcelValueRef, RawExcelValue,
};
pub use context::{LifecycleContext, MacroContext, ThreadSafeContext, WorksheetContext};
pub use convert::{ConversionLimits, FromExcel, IntoExcel};
pub use diagnostics::{
    DiagnosticCode, DiagnosticEvent, DiagnosticSeverity, DiagnosticSink, set_user_sink,
};
pub use dispatcher::{
    DispatchCallbackKind, DispatchCancelOutcome, DispatchCompletionError, DispatchConfig,
    DispatchDiagnostics, DispatchDrainReport, DispatchEnqueueError, DispatchExecutionError,
    DispatchGeneration, DispatchOperation, DispatchRequirement, DispatchResult, DispatchTicket,
    current_generation as current_dispatch_generation, diagnostics as dispatch_diagnostics,
    enqueue as enqueue_dispatch,
};
pub use error::{
    ConversionError, ExcelError, OwnedValueError, ReturnError, ReturnMaterializationError,
    ThunkError, Utf16ConversionError,
};
pub use excel_call::{
    AbortCheckMode, CallPermission, CoerceTarget, ExcelCallDescriptor, ExcelCallError,
    ExcelReturnCode, ResultRoot, SdkExcel12vBackend, XL_ABORT, XL_COERCE, XL_EVENT_REGISTER,
    XL_FREE, XL_GET_NAME, XL_SHEET_ID, XL_SHEET_NM, XLF_CALLER, XLF_REGISTER, XLF_SET_NAME,
    XLF_UNREGISTER,
};
#[cfg(feature = "xlcontime-research")]
#[doc(hidden)]
pub use excel_call::{ExperimentalOnTimeOutcome, ExperimentalOnTimeValue, XLC_ON_TIME, XLF_NOW};
pub use excel_owned::{
    ExcelOwnedConversionError, ExcelOwnedValue, ExcelReleaseError, ExcelReleasePolicy,
    ExcelXlFreeTransfer,
};
pub use metadata::{CountedUtf16Arg, ExcelReferenceArg, ExcelValueArg, NullTerminatedUtf16Arg};
pub use registration::{
    AddInDescriptor, CommandRegistration, ExcelArgumentType, ExcelReturnType, FunctionFlags,
    FunctionRegistration, FunctionSignature, RegistrationError,
};
pub use return_alloc::{ExcelReturn, xl_auto_free12};
pub use return_plan::{
    ExcelReturnArray, ExcelReturnValue, PlannedArray, PlannedArrayElement, PlannedText,
    PlannedValue, ReturnLimits, ReturnOwnershipStrategy, ReturnPlan, ReturnStorageTotals,
    ReturnText,
};
pub use runtime::{LifecycleError, LifecycleOutcome, Runtime, RuntimeDiagnostics, RuntimePhase};

/// Installs the executor for the next generated asynchronous-UDF open generation.
///
/// A successful async close permanently shuts that executor down. The XLL must
/// install a fresh executor before reopening; an active or already-pending
/// generation returns the supplied executor unchanged.
pub fn install_async_executor(
    executor: std::sync::Arc<dyn AsyncExecutor>,
    maximum_in_flight: usize,
) -> Result<(), std::sync::Arc<dyn AsyncExecutor>> {
    async_udf::install_production_executor(executor, maximum_in_flight)
}

/// Installs bounds for the next cooperative-dispatcher runtime generation.
///
/// Enqueueing does not wake Excel. A successful close permanently shuts down
/// the generation; install a fresh configuration before reopening.
pub fn install_dispatcher(config: DispatchConfig) -> Result<(), DispatchConfig> {
    dispatcher::install_production_config(config)
}
pub use value::{ExcelArray, ExcelArrayColumn, ExcelString, ExcelValue, OptionalValue};

#[cfg(feature = "macros")]
pub use excel_api_macros::{excel_command, excel_function};

/// Common imports for XLL authors.
pub mod prelude {
    pub use crate::{
        AddInDescriptor, CommandRegistration, ConversionLimits, CountedUtf16Arg, ExcelArray,
        ExcelError, ExcelReferenceArg, ExcelReturnArray, ExcelReturnValue, ExcelString, ExcelValue,
        ExcelValueArg, ExcelValueRef, FromExcel, FunctionFlags, FunctionRegistration, IntoExcel,
        NullTerminatedUtf16Arg, OptionalValue, ReturnError, ReturnLimits,
        ReturnMaterializationError, ReturnPlan, ReturnText,
    };

    #[cfg(feature = "macros")]
    pub use excel_api_macros::{excel_command, excel_function};
}
