# excel-api-sys

Raw, `no_std` Rust definitions for the 64-bit Microsoft Excel 12 C API ABI.
The crate mirrors the checked-in `XLCALL.H` contract and intentionally exposes
FFI types, unions, constants, and callback signatures without adding ownership
or lifetime guarantees.

Most add-ins should depend on `excel-api` and use this crate only when a raw ABI
boundary is required. The initial stable support target is 64-bit Excel;
32-bit Excel is unsupported.

Licensed under either Apache-2.0 or MIT.
