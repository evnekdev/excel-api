use excel_api::{
    CountedUtf16Arg, ExcelError, ExcelReferenceArg, ExcelString, ExcelValue, ExcelValueArg,
    NullTerminatedUtf16Arg, OptionalValue, ThreadSafeContext, excel_function,
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
