//! Audited runtime support for macro-generated worksheet-function thunks.
//!
//! Procedural macro output delegates raw callback borrowing, conversion,
//! context construction, panic containment, error policy, return planning,
//! materialization, and DLLFree handoff to this module.

use core::panic::AssertUnwindSafe;

use excel_api_sys::{LPXLOPER12, XCHAR, XLOPER12, XLOPER12Value};

use crate::context::{MacroContext, ThreadSafeContext, WorksheetContext};
use crate::excel_call::CallCapability;
use crate::runtime::production_backend;
use crate::{
    ConversionError, ExcelError, ExcelReturnValue, ExcelValueRef, FromExcel, IntoExcel,
    RawExcelValue, ReturnError, ReturnMaterializationError, ThunkError,
};

#[doc(hidden)]
pub type RawXloper12 = LPXLOPER12;

#[doc(hidden)]
pub type RawXchar = XCHAR;

/// One unforgeable scope for values borrowed during an Excel callback.
///
/// Macro-generated thunks obtain this value only through [`with_callback`].
/// Every borrowed argument and injected context is tied to the scope borrow,
/// so callback-owned storage and capabilities cannot escape through a
/// supported worksheet-function result.
#[doc(hidden)]
pub struct CallbackScope<'backend> {
    capability: CallCapability<'backend>,
}

impl CallbackScope<'_> {
    /// Decode one Q/U callback argument within this callback scope.
    ///
    /// # Safety
    ///
    /// `raw` must point to a readable, immutable callback-owned XLOPER12 tree
    /// that remains valid for the borrow of `self`.
    pub unsafe fn decode<'call>(
        &'call self,
        raw: LPXLOPER12,
    ) -> Result<ExcelValueRef<'call>, ThunkError> {
        // SAFETY: required by this method's callback pointer contract.
        let raw = unsafe { raw.as_ref() }.ok_or(ThunkError::NullArgument)?;
        // SAFETY: the caller keeps the complete callback tree live and immutable.
        unsafe { RawExcelValue::from_callback(raw) }
            .decode()
            .map_err(ThunkError::Decode)
    }

    /// Decode one counted D% UTF-16 callback argument.
    ///
    /// # Safety
    ///
    /// `raw` must point to a readable length prefix and payload that remain
    /// immutable and live for the borrow of `self`.
    pub unsafe fn counted_utf16<'call>(
        &'call self,
        raw: *mut XCHAR,
    ) -> Result<crate::CountedUtf16Arg<'call>, ThunkError> {
        // SAFETY: required by this method's callback pointer contract.
        let raw = unsafe { raw.as_ref() }.ok_or(ThunkError::NullArgument)?;
        // SAFETY: the caller guarantees the counted direct-string extent.
        let value =
            unsafe { crate::ExcelStr::from_counted_direct(raw) }.map_err(ThunkError::Decode)?;
        Ok(crate::CountedUtf16Arg::new(value))
    }

    /// Decode one NUL-terminated C% UTF-16 callback argument.
    ///
    /// # Safety
    ///
    /// `raw` must remain readable and immutable through its first NUL or the
    /// Excel 12 scan limit for the borrow of `self`.
    pub unsafe fn null_terminated_utf16<'call>(
        &'call self,
        raw: *mut XCHAR,
    ) -> Result<crate::NullTerminatedUtf16Arg<'call>, ThunkError> {
        // SAFETY: required by this method's callback pointer contract.
        let raw = unsafe { raw.as_ref() }.ok_or(ThunkError::NullArgument)?;
        // SAFETY: the caller guarantees the bounded NUL-terminated extent.
        let value = unsafe { crate::ExcelStr::from_null_terminated_direct(raw) }
            .map_err(ThunkError::Decode)?;
        Ok(crate::NullTerminatedUtf16Arg::new(value))
    }

    pub fn with_worksheet_context<R>(&self, body: impl FnOnce(&WorksheetContext<'_>) -> R) -> R {
        let context = WorksheetContext::new(&self.capability);
        body(&context)
    }

    pub fn with_thread_safe_context<R>(&self, body: impl FnOnce(&ThreadSafeContext<'_>) -> R) -> R {
        let context = ThreadSafeContext::new(&self.capability);
        body(&context)
    }

    pub fn with_macro_context<R>(&self, body: impl FnOnce(&MacroContext<'_>) -> R) -> R {
        let context = MacroContext::new(&self.capability);
        body(&context)
    }
}

/// Run generated thunk work inside one callback lifetime and call capability.
#[doc(hidden)]
pub fn with_callback<R>(body: impl for<'call> FnOnce(&'call CallbackScope<'call>) -> R) -> R {
    let backend = production_backend();
    let scope = CallbackScope {
        capability: CallCapability::new(backend.as_ref()),
    };
    body(&scope)
}

/// Convert a decoded Q/U argument through the audited conversion trait.
#[doc(hidden)]
pub fn from_excel<'call, T: FromExcel<'call>>(
    value: ExcelValueRef<'call>,
) -> Result<T, ThunkError> {
    T::from_excel(value).map_err(ThunkError::Conversion)
}

/// Convert supported ordinary Rust results into the logical return domain.
#[doc(hidden)]
pub trait IntoThunkReturn {
    fn into_thunk_return(self) -> Result<ExcelReturnValue, ThunkError>;
}

impl<T: IntoExcel> IntoThunkReturn for T {
    fn into_thunk_return(self) -> Result<ExcelReturnValue, ThunkError> {
        self.into_excel()
            .map(ExcelReturnValue::from)
            .map_err(ThunkError::Conversion)
    }
}

impl IntoThunkReturn for ExcelReturnValue {
    fn into_thunk_return(self) -> Result<ExcelReturnValue, ThunkError> {
        Ok(self)
    }
}

/// Preserve the supported error family while entering thunk error policy.
#[doc(hidden)]
pub trait IntoThunkError {
    fn into_thunk_error(self) -> ThunkError;
}

impl IntoThunkError for ExcelError {
    fn into_thunk_error(self) -> ThunkError {
        ThunkError::Function(self)
    }
}

impl IntoThunkError for ConversionError {
    fn into_thunk_error(self) -> ThunkError {
        ThunkError::Conversion(self)
    }
}

impl IntoThunkError for ReturnError {
    fn into_thunk_error(self) -> ThunkError {
        ThunkError::ReturnPlanning(self)
    }
}

impl IntoThunkError for ReturnMaterializationError {
    fn into_thunk_error(self) -> ThunkError {
        ThunkError::Materialization(self)
    }
}

impl IntoThunkError for ThunkError {
    fn into_thunk_error(self) -> ThunkError {
        self
    }
}

#[doc(hidden)]
pub fn function_error(error: impl IntoThunkError) -> ThunkError {
    error.into_thunk_error()
}

/// Execute a Q-returning thunk and publish one fresh DLLFree root on success.
#[doc(hidden)]
pub fn xloper12_thunk(body: impl FnOnce() -> Result<ExcelReturnValue, ThunkError>) -> LPXLOPER12 {
    match std::panic::catch_unwind(AssertUnwindSafe(|| body().and_then(return_value))) {
        Ok(Ok(pointer)) => pointer,
        Ok(Err(error)) => static_error(error_for(&error)),
        Err(_) => static_error(ExcelError::Value),
    }
}

/// Execute a direct scalar thunk with a deterministic fallback.
///
/// Direct B/A/J registration forms cannot encode an Excel error value. Input
/// failure, an ordinary `Result::Err`, or panic therefore returns the supplied
/// zero/false fallback without allowing unwinding across the ABI boundary.
#[doc(hidden)]
pub fn scalar_thunk<T: Copy>(fallback: T, body: impl FnOnce() -> Result<T, ThunkError>) -> T {
    match std::panic::catch_unwind(AssertUnwindSafe(body)) {
        Ok(Ok(value)) => value,
        Ok(Err(_)) | Err(_) => fallback,
    }
}

fn return_value(value: ExcelReturnValue) -> Result<LPXLOPER12, ThunkError> {
    let plan = value.plan().map_err(ThunkError::ReturnPlanning)?;
    let allocation = plan.materialize().map_err(ThunkError::Materialization)?;
    Ok(allocation.into_raw_for_excel())
}

#[doc(hidden)]
pub fn error_for(error: &ThunkError) -> ExcelError {
    match error {
        ThunkError::Function(error) => *error,
        ThunkError::Conversion(ConversionError::UnsupportedReference) => ExcelError::Ref,
        ThunkError::Conversion(
            ConversionError::NonFiniteNumber
            | ConversionError::NonIntegralNumber
            | ConversionError::IntegerOutOfRange,
        ) => ExcelError::Num,
        ThunkError::ReturnPlanning(ReturnError::ReferenceUnsupported) => ExcelError::Ref,
        ThunkError::Materialization(ReturnMaterializationError::AllocationFailure { .. }) => {
            ExcelError::Na
        }
        _ => ExcelError::Value,
    }
}

#[repr(transparent)]
struct StaticScalar(XLOPER12);

// SAFETY: these roots contain immutable scalar error data and no pointer.
unsafe impl Sync for StaticScalar {}

static VALUE_ERROR: StaticScalar = error_root(excel_api_sys::xlerrValue);
static NULL_ERROR: StaticScalar = error_root(excel_api_sys::xlerrNull);
static DIV0_ERROR: StaticScalar = error_root(excel_api_sys::xlerrDiv0);
static REF_ERROR: StaticScalar = error_root(excel_api_sys::xlerrRef);
static NAME_ERROR: StaticScalar = error_root(excel_api_sys::xlerrName);
static NUM_ERROR: StaticScalar = error_root(excel_api_sys::xlerrNum);
static NA_ERROR: StaticScalar = error_root(excel_api_sys::xlerrNA);
static GETTING_DATA_ERROR: StaticScalar = error_root(excel_api_sys::xlerrGettingData);

const fn error_root(error: i32) -> StaticScalar {
    StaticScalar(XLOPER12 {
        val: XLOPER12Value { err: error },
        xltype: excel_api_sys::xltypeErr,
    })
}

#[doc(hidden)]
pub fn static_error(error: ExcelError) -> LPXLOPER12 {
    let root = match error {
        ExcelError::Null => &NULL_ERROR.0,
        ExcelError::Div0 => &DIV0_ERROR.0,
        ExcelError::Value => &VALUE_ERROR.0,
        ExcelError::Ref => &REF_ERROR.0,
        ExcelError::Name => &NAME_ERROR.0,
        ExcelError::Num => &NUM_ERROR.0,
        ExcelError::Na => &NA_ERROR.0,
        ExcelError::GettingData => &GETTING_DATA_ERROR.0,
    };
    root as *const XLOPER12 as LPXLOPER12
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_function_error_preserves_its_exact_excel_code() {
        let cases = [
            (ExcelError::Null, excel_api_sys::xlerrNull),
            (ExcelError::Div0, excel_api_sys::xlerrDiv0),
            (ExcelError::Value, excel_api_sys::xlerrValue),
            (ExcelError::Ref, excel_api_sys::xlerrRef),
            (ExcelError::Name, excel_api_sys::xlerrName),
            (ExcelError::Num, excel_api_sys::xlerrNum),
            (ExcelError::Na, excel_api_sys::xlerrNA),
            (ExcelError::GettingData, excel_api_sys::xlerrGettingData),
        ];
        for (error, code) in cases {
            let pointer = static_error(error);
            // SAFETY: every static fallback root is permanently live and has
            // the error tag selecting the `err` union member.
            unsafe {
                assert_eq!((*pointer).xltype, excel_api_sys::xltypeErr);
                assert_eq!((*pointer).val.err, code);
            }
            assert_eq!(error_for(&ThunkError::Function(error)), error);
        }
    }

    #[test]
    fn scalar_and_xloper_boundaries_contain_panics() {
        assert_eq!(
            scalar_thunk(0_i32, || -> Result<i32, ThunkError> {
                panic!("scalar boundary")
            }),
            0
        );
        let pointer = xloper12_thunk(|| -> Result<ExcelReturnValue, ThunkError> {
            panic!("xloper boundary")
        });
        assert_eq!(pointer, static_error(ExcelError::Value));
    }
}
