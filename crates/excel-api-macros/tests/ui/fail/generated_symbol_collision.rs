use excel_api_macros::excel_function;

#[excel_function(name = "BAD.FOO", thunk = "bad_foo")]
fn foo() -> f64 { 0.0 }

#[excel_function(name = "BAD.FOO.UPPER", thunk = "bad_foo_upper")]
fn FOO() -> f64 { 0.0 }

fn main() {}
