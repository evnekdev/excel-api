use core::panic::AssertUnwindSafe;
use std::sync::OnceLock;

use excel_api::{
    AddInDescriptor, ExcelArgumentType, ExcelArray, ExcelError, ExcelReference, ExcelReturnType,
    ExcelReturnValue, ExcelString, ExcelValueRef, FromExcel, FunctionFlags, FunctionRegistration,
    FunctionSignature, OptionalValue, RawExcelValue, Runtime, ThunkError,
};
use excel_api_sys::{LPXLOPER12, XLOPER12, XLOPER12Value};

const ADD_ARGS: &[ExcelArgumentType] = &[ExcelArgumentType::Number, ExcelArgumentType::Number];
const GENERAL_ARG: &[ExcelArgumentType] = &[ExcelArgumentType::GeneralValue];
const REFERENCE_ARG: &[ExcelArgumentType] = &[ExcelArgumentType::GeneralReference];
const PURE: FunctionFlags = FunctionFlags {
    volatile: false,
    thread_safe: true,
    macro_type: false,
    cluster_safe: false,
};

pub fn add(x: f64, y: f64) -> f64 {
    x + y
}
pub fn echo(value: ExcelString) -> ExcelString {
    value
}
pub fn array_echo(value: ExcelArray) -> ExcelArray {
    value
}

pub static FUNCTIONS: &[FunctionRegistration] = &[
    FunctionRegistration::new(
        "rust_add",
        "RUST.ADD",
        FunctionSignature::new(ExcelReturnType::Xloper12, ADD_ARGS),
    )
    .category("Rust")
    .description("Adds two numbers")
    .arguments(&["x", "y"], &["First number", "Second number"])
    .flags(PURE),
    FunctionRegistration::new(
        "rust_echo",
        "RUST.ECHO",
        FunctionSignature::new(ExcelReturnType::Xloper12, GENERAL_ARG),
    )
    .category("Rust")
    .description("Returns text without changing its UTF-16 code units")
    .arguments(&["value"], &["Text value"])
    .flags(PURE),
    FunctionRegistration::new(
        "rust_array_echo",
        "RUST.ARRAY.ECHO",
        FunctionSignature::new(ExcelReturnType::Xloper12, GENERAL_ARG),
    )
    .category("Rust")
    .description("Deep-copies a flat value-only mixed array")
    .arguments(&["value"], &["Value-only range or array"])
    .flags(PURE),
    FunctionRegistration::new(
        "rust_reference_kind",
        "RUST.REFERENCE.KIND",
        FunctionSignature::new(ExcelReturnType::Xloper12, REFERENCE_ARG),
    )
    .category("Rust")
    .description("Reports the kind of a reference-preserving argument")
    .arguments(&["reference"], &["Reference or value"]),
    FunctionRegistration::new(
        "rust_option_kind",
        "RUST.OPTION.KIND",
        FunctionSignature::new(ExcelReturnType::Xloper12, GENERAL_ARG),
    )
    .category("Rust")
    .description("Distinguishes omitted, empty, and supplied values")
    .arguments(&["value"], &["Optional value"])
    .flags(PURE),
];

pub static ADD_IN: AddInDescriptor = AddInDescriptor::new(
    "excel-api minimal XLL",
    "Manual Rust Excel12 XLL vertical slice",
    FUNCTIONS,
);

fn runtime() -> &'static Runtime {
    static RUNTIME: OnceLock<Runtime> = OnceLock::new();
    RUNTIME.get_or_init(Runtime::production)
}

#[repr(transparent)]
struct StaticScalar(XLOPER12);
// SAFETY: these roots contain only immutable scalar error data and no pointer.
unsafe impl Sync for StaticScalar {}
static VALUE_ERROR: StaticScalar = StaticScalar(XLOPER12 {
    val: XLOPER12Value {
        err: excel_api_sys::xlerrValue,
    },
    xltype: excel_api_sys::xltypeErr,
});
static REF_ERROR: StaticScalar = StaticScalar(XLOPER12 {
    val: XLOPER12Value {
        err: excel_api_sys::xlerrRef,
    },
    xltype: excel_api_sys::xltypeErr,
});
static NUM_ERROR: StaticScalar = StaticScalar(XLOPER12 {
    val: XLOPER12Value {
        err: excel_api_sys::xlerrNum,
    },
    xltype: excel_api_sys::xltypeErr,
});
static NA_ERROR: StaticScalar = StaticScalar(XLOPER12 {
    val: XLOPER12Value {
        err: excel_api_sys::xlerrNA,
    },
    xltype: excel_api_sys::xltypeErr,
});

fn static_error(error: ExcelError) -> LPXLOPER12 {
    let root = match error {
        ExcelError::Ref => &REF_ERROR.0,
        ExcelError::Num => &NUM_ERROR.0,
        ExcelError::Na => &NA_ERROR.0,
        _ => &VALUE_ERROR.0,
    };
    root as *const XLOPER12 as LPXLOPER12
}

fn error_for(error: &ThunkError) -> ExcelError {
    match error {
        ThunkError::Conversion(excel_api::ConversionError::UnsupportedReference) => ExcelError::Ref,
        ThunkError::Conversion(
            excel_api::ConversionError::NonFiniteNumber
            | excel_api::ConversionError::NonIntegralNumber
            | excel_api::ConversionError::IntegerOutOfRange,
        ) => ExcelError::Num,
        ThunkError::ReturnPlanning(excel_api::ReturnError::ReferenceUnsupported) => ExcelError::Ref,
        ThunkError::Materialization(excel_api::ReturnMaterializationError::AllocationFailure {
            ..
        }) => ExcelError::Na,
        _ => ExcelError::Value,
    }
}

fn return_value(value: ExcelReturnValue) -> Result<LPXLOPER12, ThunkError> {
    let plan = value.plan().map_err(ThunkError::ReturnPlanning)?;
    let allocation = plan.materialize().map_err(ThunkError::Materialization)?;
    Ok(allocation.into_raw_for_excel())
}

fn thunk(body: impl FnOnce() -> Result<ExcelReturnValue, ThunkError>) -> LPXLOPER12 {
    match std::panic::catch_unwind(AssertUnwindSafe(|| body().and_then(return_value))) {
        Ok(Ok(pointer)) => pointer,
        Ok(Err(error)) => static_error(error_for(&error)),
        Err(_) => static_error(ExcelError::Value),
    }
}

unsafe fn decode<'call>(raw: LPXLOPER12) -> Result<ExcelValueRef<'call>, ThunkError> {
    // SAFETY: required by this helper's callback argument contract.
    let raw = unsafe { raw.as_ref() }.ok_or(ThunkError::NullArgument)?;
    // SAFETY: the exported thunk's ABI contract keeps the callback argument tree live and immutable.
    unsafe { RawExcelValue::from_callback(raw) }
        .decode()
        .map_err(ThunkError::Decode)
}

#[unsafe(no_mangle)]
pub extern "system" fn rust_add(x: f64, y: f64) -> LPXLOPER12 {
    thunk(|| Ok(ExcelReturnValue::Number(add(x, y))))
}

#[unsafe(no_mangle)]
/// # Safety
/// `value` must be a readable callback-owned XLOPER12 tree for the call.
pub unsafe extern "system" fn rust_echo(value: LPXLOPER12) -> LPXLOPER12 {
    thunk(|| {
        // SAFETY: forwarded from the exported thunk contract.
        let borrowed = unsafe { decode(value) }?;
        let value = ExcelString::from_excel(borrowed).map_err(ThunkError::Conversion)?;
        Ok(ExcelReturnValue::from(echo(value)))
    })
}

#[unsafe(no_mangle)]
/// # Safety
/// `value` must be a readable callback-owned value-only XLOPER12 tree for the call.
pub unsafe extern "system" fn rust_array_echo(value: LPXLOPER12) -> LPXLOPER12 {
    thunk(|| {
        // SAFETY: forwarded from the exported thunk contract.
        let borrowed = unsafe { decode(value) }?;
        let value = ExcelArray::from_excel(borrowed).map_err(ThunkError::Conversion)?;
        Ok(ExcelReturnValue::from(array_echo(value)))
    })
}

#[unsafe(no_mangle)]
/// # Safety
/// `value` must be a readable callback-owned reference-preserving XLOPER12 tree for the call.
pub unsafe extern "system" fn rust_reference_kind(value: LPXLOPER12) -> LPXLOPER12 {
    thunk(|| {
        // SAFETY: forwarded from the exported thunk contract.
        let kind = match unsafe { decode(value) }? {
            ExcelValueRef::Reference(ExcelReference::Single(_)) => "SRef",
            ExcelValueRef::Reference(ExcelReference::Multiple(_)) => "Ref",
            ExcelValueRef::Array(_) => "multi",
            ExcelValueRef::Missing(_) => "missing",
            ExcelValueRef::Nil(_) => "nil",
            _ => "scalar",
        };
        Ok(ExcelReturnValue::from(kind))
    })
}

#[unsafe(no_mangle)]
/// # Safety
/// `value` must be a readable callback-owned XLOPER12 tree for the call.
pub unsafe extern "system" fn rust_option_kind(value: LPXLOPER12) -> LPXLOPER12 {
    thunk(|| {
        // SAFETY: forwarded from the exported thunk contract.
        let decoded = unsafe { decode(value) }?;
        let optional = OptionalValue::<excel_api::ExcelValue>::from_excel(decoded)
            .map_err(ThunkError::Conversion)?;
        Ok(ExcelReturnValue::from(match optional {
            OptionalValue::Missing => "missing",
            OptionalValue::Empty => "nil",
            OptionalValue::Value(_) => "value",
        }))
    })
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoOpen() -> i32 {
    std::panic::catch_unwind(|| runtime().initialize(&ADD_IN).map(|_| 1).unwrap_or(0)).unwrap_or(0)
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoClose() -> i32 {
    std::panic::catch_unwind(|| runtime().close().map(|_| 1).unwrap_or(0)).unwrap_or(0)
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoAdd() -> i32 {
    xlAutoOpen()
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoRemove() -> i32 {
    1
}

#[unsafe(no_mangle)]
/// # Safety
/// `action` must be a readable callback-owned numeric XLOPER12 for the call.
pub unsafe extern "system" fn xlAddInManagerInfo12(action: LPXLOPER12) -> LPXLOPER12 {
    thunk(|| {
        // SAFETY: forwarded from the exported callback contract.
        let action = unsafe { decode(action) }?;
        let supported = matches!(
            action,
            ExcelValueRef::Integer(1) | ExcelValueRef::Number(1.0)
        );
        if supported {
            Ok(ExcelReturnValue::from(ADD_IN.name))
        } else {
            Ok(ExcelReturnValue::Error(ExcelError::Value))
        }
    })
}

#[unsafe(no_mangle)]
/// # Safety
/// Excel must pass null or the unique live pointer handed off by this loaded XLL, exactly once.
pub unsafe extern "system" fn xlAutoFree12(value: LPXLOPER12) {
    // SAFETY: Excel must supply the unique pointer previously handed off by this loaded XLL.
    unsafe { excel_api::xl_auto_free12(value) };
}

#[unsafe(no_mangle)]
pub extern "system" fn SetExcel12EntryPt(callback: excel_api_sys::Excel12EntryPtFn) {
    let _ = std::panic::catch_unwind(|| runtime().set_excel12_entry_point(callback));
}

const _: excel_api_sys::XlAutoOpenFn = xlAutoOpen;
const _: excel_api_sys::XlAutoCloseFn = xlAutoClose;
const _: excel_api_sys::XlAutoAddFn = xlAutoAdd;
const _: excel_api_sys::XlAutoRemoveFn = xlAutoRemove;
const _: excel_api_sys::XlAddInManagerInfo12Fn = xlAddInManagerInfo12;
const _: excel_api_sys::XlAutoFree12Fn = xlAutoFree12;
const _: excel_api_sys::SetExcel12EntryPtFn = SetExcel12EntryPt;
const _: unsafe extern "system" fn(f64, f64) -> LPXLOPER12 = rust_add;
const _: unsafe extern "system" fn(LPXLOPER12) -> LPXLOPER12 = rust_echo;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn descriptors_have_exact_signatures_and_flags() {
        assert_eq!(ADD_IN.validate(), Ok(()));
        let texts: Vec<_> = FUNCTIONS
            .iter()
            .map(|function| function.type_text().unwrap())
            .collect();
        assert_eq!(texts, ["QBB$", "QQ$", "QQ$", "QU", "QQ$"]);
    }

    #[test]
    fn add_thunk_returns_per_call_dllfree_storage() {
        let first = rust_add(2.0, 3.0);
        let second = rust_add(4.0, 5.0);
        assert_ne!(first, second);
        // SAFETY: each is a distinct fresh handoff and is reclaimed once.
        unsafe {
            xlAutoFree12(first);
            xlAutoFree12(second);
        }
    }

    #[test]
    fn panics_are_mapped_to_an_immutable_scalar_error() {
        let pointer = thunk(|| -> Result<ExcelReturnValue, ThunkError> { panic!("test panic") });
        assert_eq!(pointer, static_error(ExcelError::Value));
        // Static fallback roots carry no ownership bit and must not be passed to AutoFree.
        // SAFETY: the static fallback pointer is permanently live.
        assert_eq!(unsafe { (*pointer).xltype }, excel_api_sys::xltypeErr);
    }

    #[test]
    fn null_input_is_a_controlled_value_error() {
        // SAFETY: null is intentionally supplied to exercise defensive validation.
        let pointer = unsafe { rust_echo(core::ptr::null_mut()) };
        assert_eq!(pointer, static_error(ExcelError::Value));
    }
}
