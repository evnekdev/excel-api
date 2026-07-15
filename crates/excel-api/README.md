# excel-api

Safe Rust building blocks for native 64-bit Microsoft Excel XLL add-ins.

`excel-api` is the high-level companion to `excel-api-sys`: callback inputs are
borrowed for exactly their Excel callback lifetime, owned semantic values are
pointer-free, return planning allocates stable DLL-owned `XLOPER12` storage, and
typed contexts restrict Excel calls to documented callback capabilities.

## Quick start

```rust,no_run
use excel_api::prelude::*;

#[cfg(feature = "macros")]
#[excel_function(
    name = "RUST.ADD", thunk = "rust_add",
    arguments(left = "First addend.", right = "Second addend.")
)]
fn add(left: f64, right: f64) -> f64 { left + right }

# fn main() {}
```

The default `macros` feature re-exports `excel_function` and `excel_command`.
`xlcontime-research` is an experimental, doc-hidden compatibility probe; it is
not a supported autonomous wake mechanism. RTD, COM/Ribbon UI, task panes, and
autonomous notification are outside the stable core.

Async UDFs and the cooperative dispatcher are preview features: automated
coverage is present, while their full Excel lifecycle/pump validation remains
pending. Enqueueing dispatcher work never wakes Excel.

See the [repository user guide](https://github.com/evnekdev/excel-api/tree/master/docs/guide),
the [minimal XLL example](https://github.com/evnekdev/excel-api/tree/master/examples/minimal-xll),
and the crate Rustdoc for API-specific ownership and callback restrictions.

Supported target: 64-bit Windows Excel using the Excel 12 C API. 32-bit Excel
and arbitrary background-thread Excel calls are unsupported.

Licensed under either Apache-2.0 or MIT.
