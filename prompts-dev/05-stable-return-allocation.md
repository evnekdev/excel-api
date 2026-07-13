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

# Prompt 05: Stable DLL-owned return allocation

## Goal

Materialize a validated plan into one RAII-owned return tree.

## Required design

```rust
#[repr(C)]
ReturnAllocation {
    root: XLOPER12, // offset zero
    ...
}
```

Own stable buffers for:

- counted UTF-16;
- multi elements;
- every string nested in a multi;
- future reference memory.

Use final boxed storage before pointer publication.

Implement `ExcelReturn` with normal local drop.

Do not set DLLFree yet.

## Tests

Root offset, pointer ranges, strings, mixed multis, partial failure, high-volume
construct/drop, zero live allocations.
