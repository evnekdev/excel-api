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

# Prompt 03: Owned semantic values

## Goal

Create values independent of Excel memory.

## Required types

```rust
ExcelString
ExcelValue
ExcelArray
OptionalValue<T>
OwnedExcelReference // only if semantics are fully specified
```

Prefer:

```rust
ExcelValue::Text(ExcelString)
```

Deep-copy:

- borrowed strings;
- borrowed arrays;
- every pointer-bearing element.

Do not silently coerce references to values.

Implement checked `FromExcel` conversions. Integer conversion rejects fraction,
NaN, infinity, and out-of-range values.

Add limits for strings, elements, bytes, and recursion.

## Tests

Deep-copy independence, UTF-16 preservation, arrays, limits, worker-thread
movement.
