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

# Prompt 08: First real manually registered XLL

## Goal

Build the first working vertical slice without proc macros.

## Required work

- idempotent runtime linking/initialization;
- manual `xlfRegister` via Excel12v;
- typed descriptors and generated type text;
- registration ID storage/unregistration;
- lifecycle callbacks;
- panic boundaries;
- functions:
  - scalar add;
  - XLOPER12 text length;
  - UTF-16 echo;
  - UTF-8 uppercase;
  - flat mixed-array echo;
  - reference-preserving probe;
  - value-only/coerced range probe;
  - counted direct string input probe;
  - null-terminated direct string input probe;
  - Excel-owned result returned through consuming XLFree transfer.

## Critical validation

- R/U reference-preserving versus P/Q value-only behavior;
- missing versus nil;
- current versus active context;
- no dynamic direct simple-string return;
- thread-safe functions use fresh per-call roots;
- thread-safe and macro-sheet flags rejected together;
- real Excel repeated recalculation and unload.

## Acceptance

Excel loads, functions register, values are correct, no leaks, no panic crash,
and lifecycle remains safe under duplicate initialization.
