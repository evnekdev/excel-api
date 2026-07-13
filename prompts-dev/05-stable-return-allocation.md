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


# Task: implement stable XLL-owned return allocation

## Objective

Materialize validated return plans into stable XLL-owned memory without
publishing pointers into movable or temporary storage.

## Prerequisites

Assume prompt 04 provides a safe `ExcelReturnValue` and `ReturnPlan`.

## Required work

### 1. Memory module

Create a focused internal module, for example:

```text
crates/excel-api/src/memory/
  mod.rs
  allocation.rs
  materialize.rs
  debug.rs
```

Do not expose raw allocation internals publicly.

### 2. Return allocation layout

Implement an internal layout similar to:

```rust
#[repr(C)]
struct ReturnAllocation {
    root: XLOPER12,
    header: ReturnHeader,
    strings: Vec<Box<[u16]>>,
    arrays: Vec<Box<[XLOPER12]>>,
}
```

The exact storage may differ, but preserve these invariants:

- `root` is at offset zero;
- every published payload pointer targets stable final storage;
- all backing allocations are owned by the top-level allocation;
- no pointer targets stack memory;
- no pointer targets a `Vec` that may reallocate;
- the layout is versioned for diagnostics.

### 3. Planning-to-materialization pipeline

Implement:

```text
ExcelReturnValue
    -> ReturnPlan
    -> ReturnAllocation
    -> ExcelReturn
```

The materializer must:

1. allocate stable backing storage;
2. populate all payloads;
3. create raw pointers only after addresses are stable;
4. build array element `XLOPER12` values;
5. build the root value;
6. keep DLL-free ownership unset before handoff.

### 4. RAII owner

Implement opaque `ExcelReturn`.

Before handoff:

- normal `Drop` frees every backing buffer;
- partial construction failures clean up automatically;
- no global pointer registry is required.

### 5. Strings

Returned strings must:

- use verified Excel length-prefixed UTF-16;
- store prefix and payload together in stable memory;
- preserve interior NUL code units;
- reject overlength values before allocation.

### 6. Arrays

Returned arrays must:

- use stable boxed element storage;
- support only verified flat element types initially;
- keep all nested string buffers alive through the top-level owner;
- preserve documented row-major ordering;
- reject nested arrays.

### 7. Failure injection

Design construction so tests can simulate failure at multiple steps.

Verify that every failure before handoff leaves no live allocation.

### 8. Debug instrumentation

Under a feature such as `memory-debug`, add:

- atomic live allocation count;
- allocation sequence numbers;
- magic/layout version;
- counts of backing buffers and estimated bytes.

The production ownership path must not depend on a registry.

### 9. Tests

Add tests for:

- root offset zero;
- pointer stability after construction;
- string payload pointers inside owned backing buffers;
- array element pointers inside owned backing buffers;
- local drop cleanup;
- partial failure cleanup;
- mixed string/scalar arrays;
- high-volume construct/drop loops;
- debug live count returning to zero.

## Explicit non-goals

Do not:

- hand ownership to Excel yet;
- set DLL-free ownership bits;
- implement `xlAutoFree12`;
- add global authoritative pointer maps;
- return references or nested arrays.

## Acceptance criteria

The task is complete when:

1. all complex returns are fully owned by one RAII top-level object;
2. every raw pointer targets stable memory;
3. local drop frees everything;
4. partial failures leak nothing;
5. root offset and backing ranges are tested;
6. no ownership bit claims Excel handoff yet;
7. all tests and lints pass.
