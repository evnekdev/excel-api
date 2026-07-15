# Worksheet functions

Use `#[excel_function]` for the supported closed signature model. Supply an
Excel name, an ASCII x64 export name (`thunk`), argument help for every
Excel-visible argument, and optional category/description/flags.

```rust,no_run
use excel_api::prelude::*;

#[excel_function(
    name = "RUST.DOUBLE",
    thunk = "rust_double",
    arguments(value = "Number to double.")
)]
fn double(value: f64) -> f64 { value * 2.0 }
```

Scalar inputs include `f64`, `bool`, and supported integer forms. General
value (`Q`) and reference-preserving (`U`) inputs use metadata wrappers; see
[values](values-and-conversions.md) and [references](references.md). Dynamic
returns use owned `ExcelString`, `ExcelArray`, `ExcelValue`, `ExcelError`, or
`ExcelReturnValue`; borrowed returns are rejected at compile time.
