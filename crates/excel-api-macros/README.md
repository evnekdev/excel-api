# excel-api-macros

Procedural macros for declaring `excel-api` worksheet functions and commands.
The macros validate the supported signature catalogue, generate deterministic
registration metadata, and emit panic-contained Excel ABI thunks.

This crate is normally enabled through the default `macros` feature of
`excel-api`; direct use is uncommon.

Licensed under either Apache-2.0 or MIT.
