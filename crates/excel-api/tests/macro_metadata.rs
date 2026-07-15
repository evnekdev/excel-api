#![cfg(feature = "macros")]

use excel_api::{
    CountedUtf16Arg, ExcelError, ExcelReferenceArg, ExcelString, ExcelValue, ExcelValueArg,
    NullTerminatedUtf16Arg, OptionalValue, ThreadSafeContext, excel_function,
};
use excel_api_sys::{
    LPXLOPER12, XCHAR, XLOPER12, XLOPER12Value, xlbitDLLFree, xltypeErr, xltypeNil,
};

#[excel_function(
    name = "TEST.SCALAR",
    category = "Tests",
    description = "Scalar mapping",
    thunk = "test_scalar",
    arguments(number = "number", boolean = "boolean", integer = "integer")
)]
fn scalar(number: f64, boolean: bool, integer: i32) -> f64 {
    number + f64::from(integer) + if boolean { 1.0 } else { 0.0 }
}

#[excel_function(
    name = "TEST.EXPLICIT",
    category = "Tests",
    description = "Explicit general and direct string families",
    thunk = "test_explicit",
    arguments(
        value = "value-only",
        reference = "reference-preserving",
        counted = "counted UTF-16",
        terminated = "NUL-terminated UTF-16"
    )
)]
fn explicit_families(
    value: ExcelValueArg<'_>,
    reference: ExcelReferenceArg<'_>,
    counted: CountedUtf16Arg<'_>,
    terminated: NullTerminatedUtf16Arg<'_>,
) -> ExcelString {
    let _ = (
        value.into_inner(),
        reference.into_inner(),
        counted.into_inner(),
        terminated.into_inner(),
    );
    ExcelString::from("ok")
}

#[excel_function(
    name = "TEST.OPTIONAL",
    category = "Tests",
    description = "Optional values use Q",
    thunk = "test_optional",
    arguments(value = "optional value")
)]
fn optional(value: OptionalValue<f64>) -> ExcelValue {
    match value {
        OptionalValue::Missing => ExcelValue::Missing,
        OptionalValue::Empty => ExcelValue::Empty,
        OptionalValue::Value(value) => ExcelValue::Number(value),
    }
}

#[excel_function(
    name = "TEST.CONTEXT",
    category = "Tests",
    description = "Context injection and Result output",
    thunk = "test_context",
    thread_safe,
    arguments(value = "value")
)]
fn with_context(
    _context: &ThreadSafeContext<'_>,
    value: ExcelValue,
) -> Result<ExcelString, ExcelError> {
    let _ = value;
    Ok(ExcelString::from("ok"))
}

#[excel_function(
    name = "TEST.PANIC.DYNAMIC",
    thunk = "test_panic_dynamic",
    return_type = "xloper12"
)]
fn panic_dynamic() -> f64 {
    panic!("generated dynamic panic")
}

#[excel_function(name = "TEST.PANIC.SCALAR", thunk = "test_panic_scalar")]
fn panic_scalar() -> f64 {
    panic!("generated scalar panic")
}

#[excel_function(name = "TEST.RESULT.ERROR", thunk = "test_result_error")]
fn result_error() -> Result<ExcelString, ExcelError> {
    Err(ExcelError::Ref)
}

#[excel_function(name = "TEST.SCALAR.ERROR", thunk = "test_scalar_error")]
fn scalar_error() -> Result<f64, ExcelError> {
    Err(ExcelError::Value)
}

const _: unsafe extern "system" fn(f64, i16, i32) -> f64 = __excel_function_thunk_scalar;
const _: unsafe extern "system" fn(LPXLOPER12, LPXLOPER12, *mut XCHAR, *mut XCHAR) -> LPXLOPER12 =
    __excel_function_thunk_explicit_families;
const _: unsafe extern "system" fn(LPXLOPER12) -> LPXLOPER12 = __excel_function_thunk_optional;
const _: unsafe extern "system" fn(LPXLOPER12) -> LPXLOPER12 = __excel_function_thunk_with_context;

#[test]
fn compile_pass_examples_generate_the_closed_type_mapping() {
    assert_eq!(
        __EXCEL_FUNCTION_METADATA_SCALAR.type_text().as_deref(),
        Ok("BBAJ")
    );
    assert_eq!(
        __EXCEL_FUNCTION_METADATA_EXPLICIT_FAMILIES
            .type_text()
            .as_deref(),
        Ok("QQUD%C%")
    );
    assert_eq!(
        __EXCEL_FUNCTION_METADATA_OPTIONAL.type_text().as_deref(),
        Ok("QQ")
    );
    assert_eq!(
        __EXCEL_FUNCTION_METADATA_WITH_CONTEXT
            .type_text()
            .as_deref(),
        Ok("QQ$")
    );
}

#[test]
fn annotated_functions_remain_ordinary_callable_rust_functions() {
    assert_eq!(scalar(2.0, true, 3), 6.0);
    assert_eq!(optional(OptionalValue::Missing), ExcelValue::Missing);
    let _ordinary_signature: fn(
        &ThreadSafeContext<'_>,
        ExcelValue,
    ) -> Result<ExcelString, ExcelError> = with_context;
    let _explicit_signature: fn(
        ExcelValueArg<'_>,
        ExcelReferenceArg<'_>,
        CountedUtf16Arg<'_>,
        NullTerminatedUtf16Arg<'_>,
    ) -> ExcelString = explicit_families;
}

#[test]
fn generated_scalar_thunks_use_exact_abi_and_contain_failures() {
    // SAFETY: all arguments are immediate scalar values matching B/A/J.
    assert_eq!(unsafe { __excel_function_thunk_scalar(2.0, 1, 3) }, 6.0);
    // SAFETY: this thunk has no pointer arguments.
    assert_eq!(unsafe { __excel_function_thunk_panic_scalar() }, 0.0);
    // SAFETY: this thunk has no pointer arguments.
    assert_eq!(unsafe { __excel_function_thunk_scalar_error() }, 0.0);
}

#[test]
fn generated_dynamic_thunks_map_panics_and_result_errors() {
    // SAFETY: these thunks have no pointer arguments.
    let panic = unsafe { __excel_function_thunk_panic_dynamic() };
    assert_eq!(panic, excel_api::thunk::static_error(ExcelError::Value));

    // SAFETY: this thunk has no pointer arguments.
    let error = unsafe { __excel_function_thunk_result_error() };
    assert_eq!(error, excel_api::thunk::static_error(ExcelError::Ref));
    // SAFETY: the returned pointer is a permanently live scalar fallback.
    assert_eq!(unsafe { (*error).xltype }, xltypeErr);
}

#[test]
fn generated_direct_utf16_and_context_inputs_are_callback_scoped() {
    let mut nil = XLOPER12 {
        val: XLOPER12Value { w: 0 },
        xltype: xltypeNil,
    };
    let mut counted = [3_u16, b'A' as u16, 0, b'B' as u16];
    let mut terminated = [b'Z' as u16, 0];
    // SAFETY: all pointer arguments reference initialized callback-shaped
    // storage that remains live and immutable for the call.
    let returned = unsafe {
        __excel_function_thunk_explicit_families(
            &mut nil,
            &mut nil,
            counted.as_mut_ptr(),
            terminated.as_mut_ptr(),
        )
    };
    // SAFETY: a successful generated Q return is live until one AutoFree call.
    assert_ne!(unsafe { (*returned).xltype } & xlbitDLLFree, 0);
    // SAFETY: reclaim the unique fresh handoff exactly once.
    unsafe { excel_api::xl_auto_free12(returned) };

    // SAFETY: `nil` remains live for the callback and the injected context is
    // used only during the generated call.
    let returned = unsafe { __excel_function_thunk_with_context(&mut nil) };
    // SAFETY: reclaim the unique fresh handoff exactly once.
    unsafe { excel_api::xl_auto_free12(returned) };
}
