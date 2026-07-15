#![warn(missing_docs)]
#![doc = include_str!("../README.md")]
//!
//! # Overview
//!
//! `excel-api` is the safe, high-level layer for native 64-bit Microsoft Excel
//! XLL add-ins. It builds on [`excel_api_sys`] but does not make raw ABI
//! ownership safe by convention: callback inputs are represented by
//! callback-borrowed types such as [`ExcelValueRef`], while data that must outlive
//! a callback is represented by owned semantic types such as [`ExcelValue`] and
//! [`ExcelString`].
//!
//! Prefer the [`prelude`] for ordinary XLL authoring. The
//! [`excel_function!`][crate::excel_function] and
//! [`excel_command!`][crate::excel_command] attributes generate registration
//! metadata and panic-contained ABI thunks when the default `macros` feature is
//! enabled. Start with the repository's `examples/minimal-xll` example and the
//! [user guide](https://github.com/evnekdev/excel-api/tree/master/docs/guide).
//!
//! # Ownership and callback lifetime
//!
//! Excel owns callback arguments and may reclaim them when the callback returns;
//! they must never be retained. Convert a borrowed value with [`FromExcel`] when
//! work must outlive the callback. Return values are planned as pointer-free
//! [`ReturnPlan`]s and materialized into DLL-owned storage through
//! [`ExcelReturn`], which `xlAutoFree12` later releases. Results returned by an
//! Excel C API call are represented by [`ExcelOwnedValue`] and are released with
//! `xlFree` according to their documented policy.
//!
//! # Callback capabilities and threads
//!
//! Typed contexts are capabilities, not claims about a thread: use
//! [`WorksheetContext`], [`ThreadSafeContext`], [`MacroContext`], or
//! [`LifecycleContext`] only when Excel has issued the matching callback. Worker
//! threads must not call Excel. Async UDFs are **preview**: their worker bodies
//! own their data and publish through `xlAsyncReturn` only under the documented
//! lifecycle generation. The cooperative dispatcher is also **preview**: an
//! enqueue never wakes Excel, and work runs only when a legal callback drains it.
//!
//! # Platform and features
//!
//! The stable target is 64-bit Windows Excel using the Excel 12 C API. 32-bit
//! Excel and arbitrary background-thread Excel calls are unsupported. The
//! default `macros` feature re-exports the authoring attributes. The
//! `xlcontime-research` feature is experimental, doc-hidden compatibility
//! research and is not a supported wake mechanism. RTD, general COM, Ribbon,
//! custom task panes, and autonomous notification are outside this crate's
//! stable 1.0 scope.
//!
//! # Minimal shape
//!
//! The code below is an XLL entry point shape, so it is `no_run`: loading and
//! registration require a real Excel process.
//!
//! ```no_run
//! use excel_api::prelude::*;
//!
//! #[cfg(feature = "macros")]
//! #[excel_function(
//!     name = "RUST.ADD",
//!     thunk = "rust_add",
//!     category = "Rust",
//!     description = "Adds two numbers.",
//!     arguments(left = "First addend.", right = "Second addend.")
//! )]
//! fn add(left: f64, right: f64) -> f64 {
//!     left + right
//! }
//! # #[cfg(not(feature = "macros"))]
//! # fn main() {}
//! ```

/// Preview bounded asynchronous-UDF scheduling and cancellation primitives.
pub mod async_udf;
/// Callback-borrowed Excel values that cannot outlive their issuing callback.
pub mod borrowed;
/// Typed capability tokens for legal Excel callback contexts.
pub mod context;
/// Conversion between callback-borrowed and owned semantic values.
pub mod convert;
/// Bounded, panic-contained runtime diagnostics.
pub mod diagnostics;
/// Preview cooperative callback-drained dispatcher with no autonomous wake.
pub mod dispatcher;
/// Error types returned by safe value, conversion, and return-planning APIs.
pub mod error;
mod excel_call;
mod excel_owned;
/// Owned metadata arguments used to describe registered functions and commands.
pub mod metadata;
/// Manual registration descriptors and signature validation.
pub mod registration;
mod return_alloc;
/// Pointer-free planning of DLL-owned `XLOPER12` return storage.
pub mod return_plan;
/// Runtime initialization, registration, close, and cleanup state.
pub mod runtime;
#[doc(hidden)]
pub mod thunk;
/// Owned semantic Excel values with no callback-scoped pointers.
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
///
/// This intentionally exports the stable authoring vocabulary only. Contexts,
/// low-level calls, async scheduling, and dispatch are imported explicitly so
/// their callback and lifecycle restrictions are visible at the use site.
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
