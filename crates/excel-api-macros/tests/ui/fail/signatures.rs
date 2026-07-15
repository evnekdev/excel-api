use excel_api::{CountedUtf16Arg, ExcelStr, ExcelValueRef};
use excel_api_macros::excel_function;

#[excel_function(name = "BAD.INPUT", thunk = "bad_input", arguments(value = "value"))]
fn unsupported_input(value: Vec<u8>) -> f64 { let _ = value; 0.0 }

#[excel_function(name = "BAD.BORROW", thunk = "bad_borrow")]
fn borrowed_return() -> &'static str { "no" }

#[excel_function(name = "BAD.GENERIC", thunk = "bad_generic")]
fn generic<T>() -> f64 { 0.0 }

struct Method;
impl Method {
    #[excel_function(name = "BAD.METHOD", thunk = "bad_method")]
    fn method(&self) -> f64 { 0.0 }
}

#[excel_function(name = "BAD.ASYNC", thunk = "bad_async")]
async fn asynchronous() -> f64 { 0.0 }

#[excel_function(name = "BAD.VARIADIC", thunk = "bad_variadic")]
unsafe extern "C" fn variadic(_: i32, _: ...) -> f64 { 0.0 }

#[excel_function(name = "BAD.Q", thunk = "bad_q", arguments(value = "value"))]
fn ambiguous(value: ExcelValueRef<'_>) -> f64 { let _ = value; 0.0 }

#[excel_function(name = "BAD.STRING", thunk = "bad_string", arguments(value = "value"))]
fn optional_direct_string(value: Option<CountedUtf16Arg<'_>>) -> f64 { let _ = value; 0.0 }

#[excel_function(name = "BAD.RETURN", thunk = "bad_return")]
fn direct_return() -> ExcelStr<'static> { unreachable!() }

fn main() {}
