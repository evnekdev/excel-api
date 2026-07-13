You are working in the `excel-api` repository, a Rust workspace for building
native Microsoft Excel XLL add-ins through the Excel C API.

Before editing:

1. inspect the repository structure and all architecture documents relevant to
   this task;
2. read `ARCHITECTURE.md`, `MEMORY_OWNERSHIP_ARCHITECTURE.md`,
   `MEMORY_OWNERSHIP_ROADMAP.md`, and the ADR files under `docs/adr/`;
3. inspect the current implementation in `crates/excel-api-sys`,
   `crates/excel-api`, `crates/excel-api-macros`, and examples;
4. run the current test suite to establish a baseline;
5. do not modify unrelated files.

Global constraints:

- preserve the crate layering:
  `excel-api-sys` -> `excel-api` -> user XLL;
- keep raw ABI definitions in `excel-api-sys`;
- keep safe ownership policy in `excel-api`;
- no Rust panic may cross an FFI boundary;
- no ordinary user-facing API may require manipulation of raw pointers or Excel
  ownership bits;
- every unsafe function or block must state and justify its safety invariants;
- do not guess undocumented Excel ABI details;
- if an SDK detail cannot be verified from repository sources or authoritative
  local headers, isolate the uncertainty, document it clearly, and stop short
  of claiming production correctness;
- do not introduce COM, Ribbon, RTD, async runtimes, or packaging work in this
  task;
- keep the workspace compiling after every logical step;
- use focused tests that exercise both success and failure paths;
- avoid broad refactors unless they are required by this task.

At completion:

- run `cargo fmt --all --check`;
- run `cargo clippy --workspace --all-targets --all-features -- -D warnings`;
- run `cargo test --workspace --all-features`;
- report changed files, design decisions, tests run, and remaining uncertainties.


# Task: implement consuming handoff and `xlAutoFree12`

## Objective

Complete the exactly-once ownership transfer from Rust to Excel and the matching
cleanup path through `xlAutoFree12`.

## Prerequisites

Assume prompt 05 provides a stable RAII-owned `ExcelReturn`.

## Required work

### 1. Consuming handoff

Implement an internal method similar to:

```rust
impl ExcelReturn {
    pub(crate) fn into_raw_for_excel(self) -> *mut XLOPER12;
}
```

Requirements:

- consume `self`;
- prevent double handoff in safe code;
- set the DLL-free ownership bit only at final handoff;
- transition debug state to handed-off;
- use `Box::into_raw` or equivalent exactly once;
- return a pointer to the offset-zero root;
- perform no fallible operation after ownership is relinquished.

### 2. Cleanup callback

Implement a minimal exported/internal cleanup function for:

```rust
xlAutoFree12
```

Requirements:

- correct calling convention;
- no unwind across FFI;
- reconstruct the exact top-level allocation;
- validate debug header/state without relying on it for normal ownership;
- drop exactly once;
- avoid allocating, formatting, or calling back into Excel;
- do not assume the cleanup thread equals the creation thread.

### 3. Boundary helpers

Use a shared panic boundary for no-return callbacks.

The callback implementation itself should be designed not to panic even before
the catch boundary.

### 4. Ownership flags

Centralize ownership-bit application and masking.

No other module should manually OR the DLL-free bit.

### 5. Debug checks

Under `memory-debug`, add:

- live handed-off allocation count;
- handoff/free counters;
- state poisoning before drop;
- double-free detection where possible;
- leak summary hook usable during tests and shutdown diagnostics.

The debug mechanism must remain observational.

### 6. Tests

Add pure-Rust tests that simulate Excel:

1. build an `ExcelReturn`;
2. hand it off;
3. invoke the internal cleanup path with the returned pointer;
4. verify exactly one drop and zero live allocations.

Also test:

- null handling according to verified contract;
- scalar returns that do not require heap allocation;
- string and array returns;
- repeated handoff/free loops;
- panic before handoff;
- local drop without handoff;
- no local drop after handoff;
- multi-threaded cleanup simulation if the design allows.

### 7. Example integration

Update the minimal XLL lifecycle export to delegate to the shared cleanup path,
but do not yet add real function registration unless necessary for testing.

## Explicit non-goals

Do not implement:

- Excel-owned API result release;
- `Excel12v`;
- `xlfRegister`;
- async completion;
- COM objects.

## Acceptance criteria

The task is complete when:

1. handoff consumes ownership exactly once;
2. DLL-free ownership is applied only at handoff;
3. `xlAutoFree12` reconstructs and drops the same allocation;
4. no panic crosses FFI;
5. cleanup does not require origin-thread state;
6. repeated handoff/free tests show zero live allocations;
7. failure before handoff leaks nothing;
8. all tests and lints pass.
