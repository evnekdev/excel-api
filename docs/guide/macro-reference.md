# Macro reference

The macro crate deliberately accepts a closed signature catalogue. The exact
compile-fail diagnostics are maintained by `crates/excel-api-macros/tests/ui`.

## `#[excel_function]`

```rust,ignore
#[excel_function(
    name = "RUST.NAME", thunk = "ascii_export",
    category = "Optional category", description = "Optional help",
    return_type = "number|boolean|integer|xloper12",
    volatile, thread_safe, macro_type, cluster_safe, asynchronous,
    arguments(argument_name = "Excel help", ...)
)]
fn function(/* supported arguments */) -> /* supported result */ { /* ... */ }
```

`name` and `thunk` are required. `arguments(...)` must describe every
Excel-visible argument exactly once. The macro retains the function, generates
`__EXCEL_FUNCTION_METADATA_<NAME>`, and exports a panic-contained thunk.

| Function kind | Inputs | Result | Context/flags |
| --- | --- | --- | --- |
| Synchronous UDF | `f64`, `bool`, integer, `ExcelValueArg`, `ExcelReferenceArg`, UTF-16 wrappers | scalar, `ExcelString`, `String`, `ExcelArray`, `ExcelValue`, `ExcelReturnValue`, `ExcelError`, or supported `Result` | `WorksheetContext` or no context. |
| Thread-safe UDF | Same closed value inputs | Same supported results | `thread_safe` and optional `ThreadSafeContext`. |
| Macro-sheet function | Supported synchronous inputs | Supported synchronous result | `macro_type` and `MacroContext`; never thread/cluster-safe. |
| Async UDF (preview) | Owned `Send + 'static` inputs only | Supported owned result; registration is async void | `asynchronous`; optional `AsyncCancellationToken`; no Excel context. |

Rejected forms include borrowed returns, generic/async/variadic methods,
destructuring arguments, unsupported types, more than one injected context,
thread-safe `WorksheetContext`, macro plus thread/cluster flags, reference
arguments with `cluster_safe`, and non-async functions with cancellation token.

## `#[excel_command]`

```rust,ignore
#[excel_command(name = "RUST.COMMAND", thunk = "rust_command", description = "Optional help")]
fn command(context: &MacroContext<'_>) -> Result<(), ExcelError> { Ok(()) }
```

Commands accept only `name`, `thunk`, and `description`. Their function must
take exactly one `&MacroContext` and return `()` or a supported `Result<(), E>`.
The macro generates command registration metadata and a short success ABI thunk.

## Diagnostics and testing

The macros reject unsupported forms at compile time rather than emitting a
best-effort ABI. Add a `trybuild` pass or fail fixture before extending the
catalogue. Macro expansion is deterministic and generated symbols are checked
for collisions by the existing compile-fail tests.
