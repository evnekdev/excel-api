# Procedural Macro Architecture

## Status

M9A metadata generation is implemented. Exported thunks, conversions, panic
boundaries, and callback lifetime construction remain M9B work.

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

`name` and the future `thunk` symbol association are required. `category` and
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
metadata-only injected capabilities and do not add registration arguments.
Thread-safe and macro contexts require their corresponding flags, and a
thread-safe function cannot inject the broader worksheet context.

## Generated item

The ordinary function is emitted unchanged. The macro adds one public,
doc-hidden constant named
`__EXCEL_FUNCTION_METADATA_<UPPERCASE_FUNCTION_NAME>`. It contains a typed
`FunctionRegistration` and `FunctionSignature`; it contains no raw type text,
FFI thunk, export, unsafe block, callback conversion, or panic boundary.

The function name makes the symbol deterministic and collision-resistant
within a Rust module. Rust's ordinary duplicate-item error detects a collision
instead of silently choosing one descriptor.

## Rejections

M9A rejects generics, methods, async functions, unsafe/ABI functions,
variadics, destructuring patterns, `impl Trait`, unsupported types, ambiguous
Q/U inputs, borrowed returns, unsupported `Result` errors, and direct dynamic
string returns. M10 may improve diagnostic presentation without widening this
closed set.
