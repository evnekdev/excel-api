You are implementing one milestone in the `excel-api` Rust workspace.

Before editing:

1. inspect the repository and current branch;
2. read all architecture files relevant to this milestone;
3. read ADRs and checklists under `docs/`;
4. run the current test suite;
5. do not modify unrelated files.

Global rules:

- target Windows x64 MSVC and Excel 12+;
- keep raw ABI in `excel-api-sys`;
- keep ownership/conversion/runtime policy in `excel-api`;
- callback inputs are borrowed and read-only;
- Excel API results use `ExcelOwnedValue`;
- XLL returns use `ExcelReturn`;
- Excel-owned memory is released with `xlFree` or consuming XLFree transfer;
- XLL-owned return memory uses DLLFree and `xlAutoFree12`;
- deep-copy pointer-bearing elements into DLL-owned multis;
- no arrays-of-arrays or arrays containing references as returns;
- no static mutable return root in thread-safe UDFs;
- no panic crosses FFI;
- document every unsafe invariant;
- stop and document uncertainty rather than guessing.

At completion run:

```text
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
```

Report changed files, test commands, verified behavior, and remaining unknowns.

# Prompt 01: Verify the Excel 12 ABI

## Goal

Make `excel-api-sys` authoritative for the supported modern ABI.

## Required work

Verify:

- `XLOPER12` union and 32-bit `xltype`;
- `XLREF12`;
- variable-length `XLMREF12`;
- `FP12`;
- `xltypeBigData`;
- row/column/integer/error widths;
- `Excel12`, `Excel12v`, and lifecycle signatures;
- type and ownership bit constants;
- C API return codes;
- `xlFree` function ID;
- registration type-text codes for:
  - reference-preserving general XLOPER12;
  - value-only general XLOPER12;
  - counted wide string;
  - null-terminated wide string;
  - thread-safe, macro-sheet, volatile, cluster-safe, modify-in-place modifiers.

Add an optional C/Rust ABI-check tool using official headers.

## Tests

- size/alignment/offsets;
- constant values;
- root/union capacity;
- no SDK required for normal workspace tests.

## Non-goals

No safe parsing, conversion, calls, allocation, or registration.

## Acceptance

No placeholder raw definition is presented as verified.
