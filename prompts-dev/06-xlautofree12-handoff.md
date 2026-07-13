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

# Prompt 06: DLLFree handoff and xlAutoFree12

## Goal

Implement exactly-once ownership transfer.

## Required work

- consuming `ExcelReturn::into_raw_for_excel`;
- apply DLLFree only at final handoff;
- no fallible work afterward;
- panic-safe `xlAutoFree12`;
- reconstruct and drop the exact top-level allocation;
- free root even when payload subtype is scalar if root was heap allocated;
- never call `xlFree` from the chosen deep-owned return tree;
- never process direct simple-string buffers here.

## Tests

Scalar roots, strings, multis, repeated loops, panic before handoff,
cross-thread cleanup where valid.
