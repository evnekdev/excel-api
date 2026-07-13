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

# Prompt 02: Borrowed callback values

## Goal

Implement safe callback-lifetime views.

## Required types

```rust
RawExcelValue<'call>
ExcelValueRef<'call>
ExcelStr<'call>
ExcelArrayView<'call>
ExcelReference<'call>
```

## Required behavior

- one audited tag decoder;
- mask ownership bits before base type;
- union access only after tag validation;
- callback values never freed/modified;
- references remain references;
- arrays are flat row-major views;
- no arrays-of-arrays/references accepted as normal array elements;
- separate parsers for counted XLOPER12 strings, counted direct strings, and
  null-terminated direct strings;
- borrowed views not Send/Sync.

## Tests

All supported tags, malformed tags, missing/nil distinction, strings, dimensions,
references, lifetime constraints, non-Send/Sync.

## Non-goals

No deep copy, Excel calls, return memory, or registration.
