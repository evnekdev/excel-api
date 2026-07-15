use std::sync::{Arc, OnceLock};

use excel_api::{
    AddInDescriptor, AsyncCancellationToken, ExcelArray, ExcelError, ExcelReference,
    ExcelReferenceArg, ExcelReturnValue, ExcelString, ExcelValueRef, FunctionRegistration,
    MacroContext, OptionalValue, Runtime, ThreadPoolExecutor, excel_command, excel_function,
};
#[cfg(test)]
use excel_api::{ExcelArgumentType, ExcelReturnType, FunctionFlags, FunctionSignature};
use excel_api_sys::LPXLOPER12;

#[cfg(test)]
const ADD_ARGS: &[ExcelArgumentType] = &[ExcelArgumentType::Number, ExcelArgumentType::Number];
#[cfg(test)]
const GENERAL_ARG: &[ExcelArgumentType] = &[ExcelArgumentType::GeneralValue];
#[cfg(test)]
const REFERENCE_ARG: &[ExcelArgumentType] = &[ExcelArgumentType::GeneralReference];
#[cfg(test)]
const PURE: FunctionFlags = FunctionFlags {
    volatile: false,
    thread_safe: true,
    macro_type: false,
    cluster_safe: false,
};

#[excel_function(
    name = "RUST.ADD",
    category = "Rust",
    description = "Adds two numbers",
    thunk = "rust_add",
    return_type = "xloper12",
    thread_safe,
    arguments(x = "First number", y = "Second number")
)]
pub fn add(x: f64, y: f64) -> f64 {
    x + y
}

#[excel_function(
    name = "RUST.ECHO",
    category = "Rust",
    description = "Returns text without changing its UTF-16 code units",
    thunk = "rust_echo",
    thread_safe,
    arguments(value = "Text value")
)]
pub fn echo(value: ExcelString) -> ExcelString {
    value
}

#[excel_function(
    name = "RUST.ARRAY.ECHO",
    category = "Rust",
    description = "Deep-copies a flat value-only mixed array",
    thunk = "rust_array_echo",
    thread_safe,
    arguments(value = "Value-only range or array")
)]
pub fn array_echo(value: ExcelArray) -> ExcelArray {
    value
}

#[excel_function(
    name = "RUST.REFERENCE.KIND",
    category = "Rust",
    description = "Reports the kind of a reference-preserving argument",
    thunk = "rust_reference_kind",
    arguments(reference = "Reference or value")
)]
pub fn reference_kind(reference: ExcelReferenceArg<'_>) -> ExcelString {
    let kind = match reference.into_inner() {
        ExcelValueRef::Reference(ExcelReference::Single(_)) => "SRef",
        ExcelValueRef::Reference(ExcelReference::Multiple(_)) => "Ref",
        ExcelValueRef::Array(_) => "multi",
        ExcelValueRef::Missing(_) => "missing",
        ExcelValueRef::Nil(_) => "nil",
        _ => "scalar",
    };
    ExcelString::from(kind)
}

#[excel_function(
    name = "RUST.OPTION.KIND",
    category = "Rust",
    description = "Distinguishes omitted, empty, and supplied values",
    thunk = "rust_option_kind",
    thread_safe,
    arguments(value = "Optional value")
)]
pub fn option_kind(value: OptionalValue<excel_api::ExcelValue>) -> ExcelString {
    ExcelString::from(match value {
        OptionalValue::Missing => "missing",
        OptionalValue::Empty => "nil",
        OptionalValue::Value(_) => "value",
    })
}

#[excel_function(
    name = "RUST.ASYNC.DOUBLE",
    category = "Rust",
    description = "Doubles a number on the bounded async executor",
    thunk = "rust_async_double",
    asynchronous,
    thread_safe,
    arguments(value = "Number to double")
)]
pub fn async_double(value: f64, cancel: AsyncCancellationToken) -> Result<f64, ExcelError> {
    if cancel.is_cancellation_requested() {
        Err(ExcelError::Na)
    } else {
        Ok(value * 2.0)
    }
}

/// Minimal no-argument command used to verify the documented command ABI.
#[excel_command(
    name = "RUST.PING.COMMAND",
    description = "Verifies Rust XLL command registration and callback dispatch",
    thunk = "rust_ping_command"
)]
pub fn ping_command(_context: &MacroContext<'_>) -> Result<(), ExcelError> {
    Ok(())
}

pub static FUNCTIONS: &[FunctionRegistration] = &[
    __EXCEL_FUNCTION_METADATA_ADD,
    __EXCEL_FUNCTION_METADATA_ECHO,
    __EXCEL_FUNCTION_METADATA_ARRAY_ECHO,
    __EXCEL_FUNCTION_METADATA_REFERENCE_KIND,
    __EXCEL_FUNCTION_METADATA_OPTION_KIND,
    __EXCEL_FUNCTION_METADATA_ASYNC_DOUBLE,
];

pub static COMMANDS: &[excel_api::CommandRegistration] = &[__EXCEL_COMMAND_METADATA_PING_COMMAND];

#[cfg(test)]
pub static HANDWRITTEN_FUNCTIONS: &[FunctionRegistration] = &[
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
    "Macro-generated Rust Excel12 XLL vertical slice",
    FUNCTIONS,
)
.commands(COMMANDS);

/// Normative metadata for the handwritten M8 implementation.
///
/// M9 macro output must match these fixtures until an approved successor
/// oracle replaces them.  They intentionally describe the observable Excel
/// contract rather than providing a second registration implementation.
#[cfg(test)]
struct ManualFunctionFixture {
    rust_symbol: &'static str,
    excel_name: &'static str,
    type_text: &'static str,
    argument_names: &'static [&'static str],
    argument_descriptions: &'static [&'static str],
    flags: FunctionFlags,
    return_strategy: &'static str,
    error_mapping: &'static str,
}

#[cfg(test)]
const MANUAL_FUNCTION_FIXTURES: &[ManualFunctionFixture] = &[
    ManualFunctionFixture {
        rust_symbol: "rust_add",
        excel_name: "RUST.ADD",
        type_text: "QBB$",
        argument_names: &["x", "y"],
        argument_descriptions: &["First number", "Second number"],
        flags: PURE,
        return_strategy: "fresh DllOwnedXloper12 handoff",
        error_mapping: "panic/#VALUE!",
    },
    ManualFunctionFixture {
        rust_symbol: "rust_echo",
        excel_name: "RUST.ECHO",
        type_text: "QQ$",
        argument_names: &["value"],
        argument_descriptions: &["Text value"],
        flags: PURE,
        return_strategy: "fresh DllOwnedXloper12 handoff",
        error_mapping: "decode/conversion/#VALUE!, panic/#VALUE!",
    },
    ManualFunctionFixture {
        rust_symbol: "rust_array_echo",
        excel_name: "RUST.ARRAY.ECHO",
        type_text: "QQ$",
        argument_names: &["value"],
        argument_descriptions: &["Value-only range or array"],
        flags: PURE,
        return_strategy: "fresh DllOwnedXloper12 handoff",
        error_mapping: "reference/#REF!, numeric/#NUM!, allocation/#N/A, other/#VALUE!",
    },
    ManualFunctionFixture {
        rust_symbol: "rust_reference_kind",
        excel_name: "RUST.REFERENCE.KIND",
        type_text: "QU",
        argument_names: &["reference"],
        argument_descriptions: &["Reference or value"],
        flags: FunctionFlags {
            volatile: false,
            thread_safe: false,
            macro_type: false,
            cluster_safe: false,
        },
        return_strategy: "fresh DllOwnedXloper12 handoff",
        error_mapping: "decode/#VALUE!, panic/#VALUE!",
    },
    ManualFunctionFixture {
        rust_symbol: "rust_option_kind",
        excel_name: "RUST.OPTION.KIND",
        type_text: "QQ$",
        argument_names: &["value"],
        argument_descriptions: &["Optional value"],
        flags: PURE,
        return_strategy: "fresh DllOwnedXloper12 handoff",
        error_mapping: "decode/conversion/#VALUE!, panic/#VALUE!",
    },
];

#[cfg(test)]
const LIFECYCLE_EXPORT_FIXTURES: &[&str] = &[
    "xlAutoOpen",
    "xlAutoClose",
    "xlAutoAdd",
    "xlAutoRemove",
    "xlAddInManagerInfo12",
    "xlAutoFree12",
    "SetExcel12EntryPt",
];

fn runtime() -> &'static Runtime {
    static RUNTIME: OnceLock<Runtime> = OnceLock::new();
    RUNTIME.get_or_init(|| {
        let executor = ThreadPoolExecutor::new(2, 64).expect("constant executor bounds are valid");
        let _ = excel_api::install_async_executor(Arc::new(executor), 64);
        Runtime::production()
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
    std::panic::catch_unwind(|| runtime().close().map(|_| 1).unwrap_or(0)).unwrap_or(0)
}

#[unsafe(no_mangle)]
/// # Safety
/// `action` must be a readable callback-owned numeric XLOPER12 for the call.
pub unsafe extern "system" fn xlAddInManagerInfo12(action: LPXLOPER12) -> LPXLOPER12 {
    excel_api::thunk::xloper12_thunk(|| {
        excel_api::thunk::with_callback(|scope| {
            // SAFETY: forwarded from the exported callback contract.
            let action = unsafe { scope.decode(action) }?;
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
const _: unsafe extern "system" fn(f64, f64) -> LPXLOPER12 = __excel_function_thunk_add;
const _: unsafe extern "system" fn(LPXLOPER12) -> LPXLOPER12 = __excel_function_thunk_echo;
const _: unsafe extern "system" fn(f64, LPXLOPER12) = __excel_function_thunk_async_double;
const _: extern "system" fn() -> i16 = __excel_command_thunk_ping_command;

#[cfg(test)]
mod tests {
    use super::*;
    use excel_api_sys::{
        XLOPER12, XLOPER12Array, XLOPER12SRef, XLOPER12Value, XLREF12, XLTYPE_MASK, xlbitDLLFree,
        xlbitXLFree, xltypeBool, xltypeMissing, xltypeMulti, xltypeNum, xltypeSRef, xltypeStr,
    };

    fn normalize_snapshot_newlines(snapshot: &str) -> String {
        snapshot.lines().collect::<Vec<_>>().join("\n")
    }

    unsafe fn returned_text(pointer: LPXLOPER12) -> Vec<u16> {
        // SAFETY: the caller supplies a live generated xltypeStr handoff.
        let root = unsafe { &*pointer };
        assert_eq!(root.xltype & XLTYPE_MASK, xltypeStr);
        // SAFETY: the active union member follows from the validated tag.
        let text = unsafe { root.val.str };
        // SAFETY: a materialized counted string always has a readable prefix.
        let len = usize::from(unsafe { *text });
        // SAFETY: the counted buffer owns exactly prefix plus payload.
        unsafe { core::slice::from_raw_parts(text.add(1), len) }.to_vec()
    }

    unsafe fn assert_generated_handoff(pointer: LPXLOPER12, expected_type: u32) {
        // SAFETY: the caller supplies a live generated handoff.
        let xltype = unsafe { (*pointer).xltype };
        assert_eq!(xltype & XLTYPE_MASK, expected_type);
        assert_ne!(xltype & xlbitDLLFree, 0);
        assert_eq!(xltype & xlbitXLFree, 0);
    }

    #[test]
    fn descriptors_have_exact_signatures_and_flags() {
        assert_eq!(ADD_IN.validate(), Ok(()));
        let texts: Vec<_> = FUNCTIONS
            .iter()
            .map(|function| function.type_text().unwrap())
            .collect();
        assert_eq!(texts, ["QBB$", "QQ$", "QQ$", "QU", "QQ$", ">BX$"]);
        assert_eq!(COMMANDS.len(), 1);
        assert_eq!(COMMANDS[0].type_text(), "I");
    }

    #[test]
    fn handwritten_registration_matches_the_m8_oracle() {
        assert!(FUNCTIONS.len() >= MANUAL_FUNCTION_FIXTURES.len());
        for (function, fixture) in HANDWRITTEN_FUNCTIONS.iter().zip(MANUAL_FUNCTION_FIXTURES) {
            assert_eq!(function.rust_symbol, fixture.rust_symbol);
            assert_eq!(function.excel_name, fixture.excel_name);
            assert_eq!(function.type_text().as_deref(), Ok(fixture.type_text));
            assert_eq!(function.argument_names, fixture.argument_names);
            assert_eq!(
                function.argument_descriptions,
                fixture.argument_descriptions
            );
            assert_eq!(function.flags, fixture.flags);
            assert_eq!(fixture.return_strategy, "fresh DllOwnedXloper12 handoff");
            assert!(!fixture.error_mapping.is_empty());
        }
    }

    #[test]
    fn generated_metadata_exactly_matches_the_handwritten_m8_oracle() {
        assert!(FUNCTIONS.len() >= HANDWRITTEN_FUNCTIONS.len());
        for (generated, handwritten) in FUNCTIONS.iter().zip(HANDWRITTEN_FUNCTIONS) {
            assert_eq!(generated.rust_symbol, handwritten.rust_symbol);
            assert_eq!(generated.excel_name, handwritten.excel_name);
            assert_eq!(generated.signature, handwritten.signature);
            assert_eq!(generated.type_text(), handwritten.type_text());
            assert_eq!(generated.category, handwritten.category);
            assert_eq!(generated.description, handwritten.description);
            assert_eq!(generated.argument_names, handwritten.argument_names);
            assert_eq!(
                generated.argument_descriptions,
                handwritten.argument_descriptions
            );
            assert_eq!(generated.flags, handwritten.flags);
        }
    }

    #[test]
    fn generated_metadata_expansion_snapshot_is_stable() {
        let snapshot = FUNCTIONS
            .iter()
            .take(HANDWRITTEN_FUNCTIONS.len())
            .map(|function| {
                format!(
                    "{}|{}|{}|{}|{}|{}|{}|volatile={},thread_safe={},macro_type={},cluster_safe={}",
                    function.rust_symbol,
                    function.excel_name,
                    function.type_text().unwrap(),
                    function.category.unwrap_or_default(),
                    function.description.unwrap_or_default(),
                    function.argument_names.join(","),
                    function.argument_descriptions.join(","),
                    function.flags.volatile,
                    function.flags.thread_safe,
                    function.flags.macro_type,
                    function.flags.cluster_safe,
                )
            })
            .collect::<Vec<_>>()
            .join("\n");
        assert_eq!(
            snapshot,
            normalize_snapshot_newlines(include_str!(
                "../tests/snapshots/m8-generated-metadata.snap"
            ))
        );
    }

    #[test]
    fn async_metadata_uses_exact_void_handle_registration() {
        let registration = &FUNCTIONS[5];
        assert!(registration.is_asynchronous());
        assert_eq!(registration.type_text().as_deref(), Ok(">BX$"));
        assert_eq!(registration.argument_names, ["value"]);
        assert_eq!(registration.signature.arguments.len(), 2);
        assert_eq!(
            registration.signature.arguments[1],
            ExcelArgumentType::AsyncHandle
        );
    }

    #[test]
    fn metadata_snapshot_accepts_windows_and_unix_newlines() {
        assert_eq!(
            normalize_snapshot_newlines("first\nsecond\n"),
            normalize_snapshot_newlines("first\r\nsecond\r\n")
        );
    }

    #[test]
    fn lifecycle_and_thunk_exports_match_the_m8_oracle() {
        assert_eq!(LIFECYCLE_EXPORT_FIXTURES.len(), 7);
        // The typed const assertions above prove the callback ABIs. CI checks
        // the final PE export table because unit tests cannot inspect it.
        assert!(LIFECYCLE_EXPORT_FIXTURES.contains(&"xlAutoFree12"));
        assert!(LIFECYCLE_EXPORT_FIXTURES.contains(&"SetExcel12EntryPt"));
    }

    #[test]
    fn thunk_error_mapping_matches_the_m8_oracle() {
        assert_eq!(
            excel_api::thunk::error_for(&excel_api::ThunkError::NullArgument),
            ExcelError::Value
        );
        assert_eq!(
            excel_api::thunk::error_for(&excel_api::ThunkError::Conversion(
                excel_api::ConversionError::UnsupportedReference
            )),
            ExcelError::Ref
        );
        assert_eq!(
            excel_api::thunk::error_for(&excel_api::ThunkError::Conversion(
                excel_api::ConversionError::NonFiniteNumber
            )),
            ExcelError::Num
        );
        assert_eq!(
            excel_api::thunk::error_for(&excel_api::ThunkError::ReturnPlanning(
                excel_api::ReturnError::ReferenceUnsupported
            )),
            ExcelError::Ref
        );
        assert_eq!(
            excel_api::thunk::error_for(&excel_api::ThunkError::Materialization(
                excel_api::ReturnMaterializationError::AllocationFailure { storage: "fixture" }
            )),
            ExcelError::Na
        );
    }

    #[test]
    fn add_thunk_returns_per_call_dllfree_storage() {
        // SAFETY: the scalar inputs satisfy the generated ABI contract.
        let first = unsafe { __excel_function_thunk_add(2.0, 3.0) };
        // SAFETY: the scalar inputs satisfy the generated ABI contract.
        let second = unsafe { __excel_function_thunk_add(4.0, 5.0) };
        assert_ne!(first, second);
        // SAFETY: each is a distinct fresh handoff and is reclaimed once.
        unsafe {
            xlAutoFree12(first);
            xlAutoFree12(second);
        }
    }

    #[test]
    fn generated_command_has_the_documented_short_success_abi() {
        assert_eq!(__excel_command_thunk_ping_command(), 1);
    }

    #[test]
    fn generated_thunks_match_m8_values_tags_q_u_and_ownership() {
        // SAFETY: scalar arguments match the generated B/B ABI.
        let add = unsafe { __excel_function_thunk_add(2.0, 3.0) };
        // SAFETY: `add` is a live generated handoff.
        unsafe { assert_generated_handoff(add, xltypeNum) };
        // SAFETY: the numeric tag selects the numeric union member.
        assert_eq!(unsafe { (*add).val.num }, 5.0);
        // SAFETY: reclaim the unique handoff exactly once.
        unsafe { xlAutoFree12(add) };

        let mut counted = [3_u16, b'A' as u16, 0, b'B' as u16];
        let mut text = XLOPER12 {
            val: XLOPER12Value {
                str: counted.as_mut_ptr(),
            },
            xltype: xltypeStr,
        };
        // SAFETY: `text` and its counted backing remain live for the call.
        let echo = unsafe { __excel_function_thunk_echo(&mut text) };
        // SAFETY: `echo` is a live generated handoff.
        unsafe { assert_generated_handoff(echo, xltypeStr) };
        // SAFETY: the returned tag and storage were just validated.
        assert_eq!(unsafe { returned_text(echo) }, &counted[1..]);
        // SAFETY: reclaim the unique handoff exactly once.
        unsafe { xlAutoFree12(echo) };

        let mut elements = [
            XLOPER12 {
                val: XLOPER12Value { num: 2.5 },
                xltype: xltypeNum,
            },
            XLOPER12 {
                val: XLOPER12Value { xbool: 1 },
                xltype: xltypeBool,
            },
        ];
        let mut array = XLOPER12 {
            val: XLOPER12Value {
                array: XLOPER12Array {
                    lparray: elements.as_mut_ptr(),
                    rows: 1,
                    columns: 2,
                },
            },
            xltype: xltypeMulti,
        };
        // SAFETY: `array` is a valid flat Q callback tree for the call.
        let echoed_array = unsafe { __excel_function_thunk_array_echo(&mut array) };
        // SAFETY: `echoed_array` is a live generated handoff.
        unsafe { assert_generated_handoff(echoed_array, xltypeMulti) };
        // SAFETY: the multi tag selects a materialized two-element array.
        let returned_array = unsafe { (*echoed_array).val.array };
        assert_eq!((returned_array.rows, returned_array.columns), (1, 2));
        // SAFETY: the returned multi owns two initialized elements.
        let returned_elements = unsafe { core::slice::from_raw_parts(returned_array.lparray, 2) };
        assert_eq!(returned_elements[0].xltype, xltypeNum);
        assert_eq!(returned_elements[1].xltype, xltypeBool);
        // SAFETY: tags select the corresponding union members.
        assert_eq!(unsafe { returned_elements[0].val.num }, 2.5);
        // SAFETY: the Boolean tag selects the Boolean union member.
        assert_eq!(unsafe { returned_elements[1].val.xbool }, 1);
        // SAFETY: reclaim the unique handoff exactly once.
        unsafe { xlAutoFree12(echoed_array) };

        let mut reference = XLOPER12 {
            val: XLOPER12Value {
                sref: XLOPER12SRef {
                    count: 1,
                    reference: XLREF12 {
                        rwFirst: 0,
                        rwLast: 1,
                        colFirst: 0,
                        colLast: 1,
                    },
                },
            },
            xltype: xltypeSRef,
        };
        // SAFETY: `reference` is a valid U callback value for the call.
        let kind = unsafe { __excel_function_thunk_reference_kind(&mut reference) };
        // SAFETY: the generated U thunk returned a live text handoff.
        let kind_text = unsafe { returned_text(kind) };
        assert_eq!(kind_text, "SRef".encode_utf16().collect::<Vec<_>>());
        // SAFETY: reclaim the unique handoff exactly once.
        unsafe { xlAutoFree12(kind) };

        let mut missing = XLOPER12 {
            val: XLOPER12Value { w: 0 },
            xltype: xltypeMissing,
        };
        // SAFETY: `missing` is a valid Q callback value for the call.
        let optional = unsafe { __excel_function_thunk_option_kind(&mut missing) };
        // SAFETY: the generated optional thunk returned a live text handoff.
        let optional_text = unsafe { returned_text(optional) };
        assert_eq!(optional_text, "missing".encode_utf16().collect::<Vec<_>>());
        // SAFETY: reclaim the unique handoff exactly once.
        unsafe { xlAutoFree12(optional) };

        assert_eq!(FUNCTIONS[2].signature.arguments, GENERAL_ARG);
        assert_eq!(FUNCTIONS[3].signature.arguments, REFERENCE_ARG);
    }

    #[test]
    fn panics_are_mapped_to_an_immutable_scalar_error() {
        let pointer = excel_api::thunk::xloper12_thunk(
            || -> Result<ExcelReturnValue, excel_api::ThunkError> { panic!("test panic") },
        );
        assert_eq!(pointer, excel_api::thunk::static_error(ExcelError::Value));
        // Static fallback roots carry no ownership bit and must not be passed to AutoFree.
        // SAFETY: the static fallback pointer is permanently live.
        assert_eq!(unsafe { (*pointer).xltype }, excel_api_sys::xltypeErr);
    }

    #[test]
    fn null_input_is_a_controlled_value_error() {
        // SAFETY: null is intentionally supplied to exercise defensive validation.
        let pointer = unsafe { __excel_function_thunk_echo(core::ptr::null_mut()) };
        assert_eq!(pointer, excel_api::thunk::static_error(ExcelError::Value));
    }
}
