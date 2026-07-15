# Getting started

`excel-api` builds native `.xll` add-ins for 64-bit Windows Excel. Use the
high-level crate and its default `macros` feature; use `excel-api-sys` only for
an audited raw ABI boundary.

```toml
[dependencies]
excel-api = "0.1"
```

```rust,no_run
use excel_api::prelude::*;

#[excel_function(
    name = "RUST.ADD",
    thunk = "rust_add",
    category = "Rust",
    description = "Adds two values.",
    arguments(left = "First value.", right = "Second value.")
)]
fn add(left: f64, right: f64) -> f64 { left + right }
```

The attribute retains `add` as normal Rust code, generates registration
metadata, and emits an Excel ABI thunk named `rust_add`. Add lifecycle exports
and registration using the repository's `examples/minimal-xll` as the complete
reference. The snippet is `no_run` because registration and loading require an
Excel process.

**Stable target:** 64-bit Windows Excel. **Unsupported:** 32-bit Excel and
calling Excel from arbitrary worker threads.
