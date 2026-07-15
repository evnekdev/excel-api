# excel-api-macros

Procedural attributes for the `excel-api` native XLL framework. Normally use
them through `excel-api`'s default `macros` feature rather than depending on
this crate directly.

`#[excel_function]` and `#[excel_command]` retain the annotated Rust function,
generate deterministic registration metadata, and emit a panic-contained native
Excel ABI thunk. Their signature model is deliberately closed: unsupported
argument/result types, borrowed returns, generic functions, incompatible
context/flag combinations, and malformed metadata are compile-time errors.

See the complete [macro reference](https://github.com/evnekdev/excel-api/blob/master/docs/guide/macro-reference.md)
and the `trybuild` fixtures in the repository for supported and rejected forms.
Async function support is preview and requires owned inputs; it never permits a
worker to call Excel directly.

Licensed under either Apache-2.0 or MIT.
