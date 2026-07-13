# Unified architecture and prompt replacement package

This package is intended to replace the architecture and Codex prompt files
previously generated for `excel-api`.

Copy the contents into the repository root while preserving paths.

## Scope

The revision incorporates:

- the uploaded second edition of Steve Dalton's book;
- official Excel C API memory and callback rules;
- the previously agreed Rust ownership model;
- the existing workspace and three-crate publication plan;
- corrections for strings, arrays, references, lifecycle, registration,
  thread-safe returns, Excel-owned results, and C API call legality.

## Important policy changes

1. Use `XLOPER12`/`xloper12` as the primary modern public ABI.
2. Treat `R/U` and `P/Q` registration forms as semantically different:
   reference-preserving versus value-only/coerced arguments.
3. Keep three string ABI parsers:
   `xltypeStr`, counted wide direct strings, and null-terminated wide strings.
4. Prefer deep-copy ownership for DLL-created arrays.
5. Support `xlbitXLFree` only as a consuming transfer of Excel-owned results.
6. Use per-call heap-owned return roots for the first thread-safe implementation.
7. Do not use thread-local return slots as the initial design.
8. Make lifecycle initialization idempotent because Excel may call add-in
   callbacks in surprising orders.
9. Keep reference values distinct from arrays and values.
10. Do not support arrays-of-arrays or arrays containing references as return
    values in the first implementation.
