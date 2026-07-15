use excel_api::{ExcelError, ExcelReferenceArg, MacroContext, ThreadSafeContext, WorksheetContext};
use excel_api_macros::excel_function;

#[excel_function(name = "BAD.OVERRIDE", thunk = "bad_override", return_type = "boolean")]
fn incompatible_override() -> f64 { 0.0 }

#[excel_function(name = "BAD.CLUSTER", thunk = "bad_cluster", cluster_safe, arguments(reference = "reference"))]
fn unjustified_cluster(reference: ExcelReferenceArg<'_>) -> f64 { let _ = reference; 0.0 }

#[excel_function(name = "BAD.CONTEXT", thunk = "bad_context")]
fn missing_thread_flag(_context: &ThreadSafeContext<'_>) -> f64 { 0.0 }

#[excel_function(name = "BAD.CONTEXT2", thunk = "bad_context2", thread_safe)]
fn broad_context(_context: &WorksheetContext<'_>) -> f64 { 0.0 }

#[excel_function(name = "BAD.CONTEXT3", thunk = "bad_context3")]
fn missing_macro_flag(_context: &MacroContext<'_>) -> f64 { 0.0 }

#[excel_function(name = "BAD.FLAGS", thunk = "bad_flags", macro_type, thread_safe)]
fn incompatible_flags() -> f64 { 0.0 }

#[excel_function(name = "BAD.RESULT", thunk = "bad_result")]
fn unsupported_error() -> Result<f64, String> { Err(String::new()) }

#[excel_function(name = "BAD.ATTRIBUTE", thunk = "bad_attribute", unknown)]
fn invalid_attribute() -> f64 { 0.0 }

#[excel_function(name = "BAD.HELP", thunk = "bad_help", arguments(missing = "missing"))]
fn metadata_mismatch(value: f64) -> f64 { value }

#[excel_function(name = "BAD.EXPORT", thunk = "not-an-export")]
fn invalid_export() -> f64 { 0.0 }

#[excel_function(name = "BAD.DUPLICATE", name = "BAD.DUPLICATE2", thunk = "bad_duplicate")]
fn duplicate_attribute() -> f64 { 0.0 }

fn main() {}
