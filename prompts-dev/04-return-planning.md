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

# Prompt 04: Logical return planning

## Goal

Validate every return before raw allocation.

## Required types

```rust
ExcelReturnValue
ReturnText
ReturnPlan
ReturnError
```

## Rules

- planner is safe Rust;
- no raw pointers;
- checked arithmetic;
- strings may originate as UTF-8 or UTF-16;
- arrays are flat and deeply owned;
- references are separate return variants and deferred unless fully supported;
- no arrays-of-arrays/references;
- strategy for ordinary dynamic returns is DLL-owned XLOPER12;
- XLFree transfer is not produced from ordinary Rust values.

## Tests

Limits, overflows, mixed arrays, strings, deterministic plan.
