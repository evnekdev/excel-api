# Procedural Macro Architecture

Status: planned after manual registration works.

Macros generate:

- exported thunk;
- callback lifetime scope;
- argument parser/converter;
- context token injection;
- panic boundary;
- logical return conversion;
- registration metadata;
- compile-time errors.

The macro maps Rust types to a verified registration signature. It must never
accept an arbitrary unsupported type and defer failure to Excel.
