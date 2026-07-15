# excel-api-sys

Raw, `no_std`, ABI-focused Rust definitions for the 64-bit Microsoft Excel 12 C
API. The crate mirrors the checked-in `XLCALL.H` contract: primitive aliases,
constants, C-layout structs/unions, and callback signatures.

This is not a safe ownership layer. `XLOPER12` tags and ownership bits describe
Excel protocol state; they do not validate pointers, grant Rust ownership, or
permit a callback value to escape its lifetime. Direct callers must verify the
Excel-issued callback capability, calling convention, tag/union member,
pointer alignment and extent, allocation origin, aliasing, and cleanup protocol.

Most XLL authors should depend on `excel-api`, which provides borrowed views,
owned values, return allocation, `xlFree` RAII, registration, and typed
contexts. The supported target is 64-bit Windows Excel; 32-bit Excel and
non-Windows hosts are unsupported.

Licensed under either Apache-2.0 or MIT.
