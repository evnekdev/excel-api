# Procedural Macro Architecture

## Status

M9A metadata generation and M9B worksheet-function ABI thunks are implemented.
M10 retains diagnostic refinement and generated-symbol compatibility review.

## Stable M9A attribute syntax

```rust
#[excel_function(
    name = "RUST.ADD",
    category = "Rust",
    description = "Adds two numbers",
    thunk = "rust_add",
    return_type = "xloper12",
    thread_safe,
    arguments(x = "First number", y = "Second number")
)]
fn add(x: f64, y: f64) -> f64 { x + y }
```

`name` and the exact `thunk` export name are required. `category` and
`description` are optional. `arguments(...)` must name every Excel-visible
parameter exactly once and excludes injected context parameters. Supported
flags are `volatile`, `thread_safe`, `macro_type`, and `cluster_safe`.

`return_type` is an optional explicit registration override with values
`number`, `boolean`, `integer`, or `xloper12`. It exists so the handwritten M8
oracle's DLLFree `Q` return contract can be expressed without changing the
ordinary Rust function's logical return type.

## Closed type mapping

Inputs:

| Rust type | Registration family |
|---|---|
| `f64` | `B` |
| `bool` | `A` |
| `i16`, `i32`, `u16` | `J` |
| `ExcelString`, `String`, `ExcelArray`, `ExcelValue` | `Q` |
| `ExcelValueArg<'_>` | explicit `Q` |
| `ExcelReferenceArg<'_>` | explicit `U` |
| `CountedUtf16Arg<'_>` | `D%` |
| `NullTerminatedUtf16Arg<'_>` | `C%` |
| `Option<T>`, `OptionalValue<T>` | `Q`, or `U` for an explicit reference wrapper |

Raw `ExcelValueRef<'_>` is rejected because it cannot choose Q versus U.
Direct UTF-16 arguments cannot be optional.

Outputs map scalar types to `B`, `A`, or `J`; owned `ExcelString`, `String`,
`ExcelArray`, `ExcelValue`, `ExcelReturnValue`, and `ExcelError` map to `Q`.
`Result<T, E>` maps from `T` when `E` is one of the documented excel-api error
types. Borrowed and direct dynamic-string returns are rejected.

`&WorksheetContext<'_>`, `&ThreadSafeContext<'_>`, and `&MacroContext<'_>` are
callback-scope-injected capabilities and do not add registration arguments.
Thread-safe and macro contexts require their corresponding flags, and a
thread-safe function cannot inject the broader worksheet context.

## Generated items

The ordinary function is emitted unchanged. The macro adds one public,
doc-hidden constant named
`__EXCEL_FUNCTION_METADATA_<UPPERCASE_FUNCTION_NAME>`. It contains a typed
`FunctionRegistration` and `FunctionSignature`.

M9B also adds a doc-hidden Rust item named
`__excel_function_thunk_<lowercase_function_name>` and gives it the exact
ASCII export name supplied by `thunk`. The same internal `ArgumentKind` and
`ResultKind` values select registration metadata and raw ABI parameters, so
type text and the exported signature cannot diverge.

The function name makes both Rust item names deterministic and collision-
resistant within a module. Rust detects generated-item collisions; the linker
detects duplicate exact export names.

## ABI and callback pipeline

| Registration | Generated ABI |
|---|---|
| `B` | `f64` |
| `A` | `i16` (`short`, 0/1) |
| `J` | `i32` |
| `Q`, `U` | `*mut XLOPER12` |
| `D%`, `C%` | `*mut XCHAR` |

All pointer thunks are unsafe `extern "system"` exports. Generated bodies are
thin: they enter `excel_api::thunk::with_callback`, borrow Q/U and direct
UTF-16 inputs through the callback scope, use `FromExcel`, inject the one legal
context, call the ordinary function, and delegate result handling.

Q returns use the audited logical-plan/materialize/DLLFree path and one fresh
root per success. Panics and failures map to immutable exact Excel error roots.
Direct B/A/J returns cannot represent an Excel error; their documented
fallback for conversion failure, `Result::Err`, or panic is zero/false. A
scalar `return_type` override must match the Rust success type; `xloper12` is
the only cross-family override.

The callback scope creates `WorksheetContext`, `ThreadSafeContext`, or
`MacroContext` with the shared production callback backend and prevents a
supported result from carrying callback borrows out of the invocation.

## Rejections

The macro rejects generics, methods, async functions, unsafe/ABI functions,
variadics, destructuring patterns, `impl Trait`, unsupported types, ambiguous
Q/U inputs, borrowed returns, unsupported `Result` errors, and direct dynamic
string returns. M9B additionally rejects invalid export identifiers and
incompatible scalar return overrides. M10 may improve diagnostic presentation
without widening this closed set.

Official ABI source: [xlfRegister (Form 1)](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1).
