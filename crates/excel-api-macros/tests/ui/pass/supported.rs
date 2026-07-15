use excel_api::{
    CountedUtf16Arg, ExcelArray, ExcelError, ExcelReferenceArg, ExcelString, ExcelValue,
    ExcelValueArg, MacroContext, NullTerminatedUtf16Arg, OptionalValue, ThreadSafeContext,
    WorksheetContext,
};
use excel_api_macros::excel_function;

#[excel_function(name = "PASS.SCALARS", thunk = "pass_scalars", arguments(a = "a", b = "b", c = "c"))]
fn scalars(a: f64, b: bool, c: i32) -> i16 { let _ = (b, c); a as i16 }

#[excel_function(name = "PASS.VALUES", thunk = "pass_values", arguments(value = "value", reference = "reference", counted = "counted", nul = "nul"))]
fn values(value: ExcelValueArg<'_>, reference: ExcelReferenceArg<'_>, counted: CountedUtf16Arg<'_>, nul: NullTerminatedUtf16Arg<'_>) -> ExcelString {
    let _ = (value, reference, counted, nul); ExcelString::from("ok")
}

#[excel_function(name = "PASS.OPTIONAL", thunk = "pass_optional", arguments(value = "value"))]
fn optional(value: OptionalValue<f64>) -> ExcelValue { let _ = value; ExcelValue::Empty }

#[excel_function(name = "PASS.WORKSHEET", thunk = "pass_worksheet")]
fn worksheet(_context: &WorksheetContext<'_>) -> String { String::new() }

#[excel_function(name = "PASS.THREAD", thunk = "pass_thread", thread_safe)]
fn thread(_context: &ThreadSafeContext<'_>) -> ExcelArray { unreachable!() }

#[excel_function(name = "PASS.MACRO", thunk = "pass_macro", macro_type)]
fn macro_context(_context: &MacroContext<'_>) -> Result<ExcelString, ExcelError> { Ok(ExcelString::from("ok")) }

fn main() {}
