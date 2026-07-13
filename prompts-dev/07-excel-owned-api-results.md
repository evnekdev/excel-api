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


# Task: implement RAII ownership for Excel API results

## Objective

Represent and release values returned by `Excel12v` or related Excel APIs
without confusing them with callback inputs or XLL-owned return allocations.

## Prerequisites

Assume the XLL-owned return path and `xlAutoFree12` are complete.

## Required work

### 1. Result ownership model

Introduce:

```rust
pub(crate) struct ExcelOwnedValue {
    raw: XLOPER12,
    release: ExcelReleasePolicy,
}
```

Define an explicit `ExcelReleasePolicy` based on verified Excel contracts.

Possible states may include:

- no release required;
- release through a specific Excel API mechanism;
- release restricted to a valid Excel/main-thread context.

Do not infer policies casually from one ownership bit.

### 2. Construction

Only the Excel-call layer may construct `ExcelOwnedValue`.

Construction must capture:

- originating call category;
- raw result status;
- verified ownership flags;
- release requirements;
- thread/context restrictions.

### 3. Safe access

Allow:

- borrowed inspection through the same validated decoding path;
- deep copy into `ExcelValue`;
- explicit conversion errors.

Do not expose unrestricted raw ownership mutation.

### 4. Drop

`Drop` must:

- invoke only the verified release mechanism;
- never use `xlAutoFree12`;
- never panic;
- record diagnostics on release failure;
- respect any required thread/context restrictions.

If safe automatic release cannot be guaranteed for a call category, do not
support that category yet.

### 5. Excel call abstraction

Introduce only the minimum internal call wrapper needed to exercise result
ownership.

Keep call IDs and context legality explicit.

### 6. Tests

Add tests using an injectable/mock release backend where practical:

- no-release policy;
- exactly-once release;
- deep-copy before release;
- error-path release;
- drop after conversion failure;
- release failure diagnostics;
- wrong-context rejection;
- proof that `xlAutoFree12` is never used.

### 7. Documentation

For every supported Excel call category, document:

- result ownership;
- release function;
- valid release context;
- whether the result may be copied;
- relevant SDK source/provenance.

## Explicit non-goals

Do not broadly wrap the Excel C API.

Do not add registration, worksheet commands, or async behavior except what is
strictly needed to validate ownership.

## Acceptance criteria

The task is complete when:

1. Excel API results have a distinct RAII wrapper;
2. release policy is explicit and verified;
3. every supported path releases exactly once;
4. XLL-return cleanup and Excel-result cleanup cannot be confused;
5. unsupported ownership contracts are unavailable safely;
6. all tests and lints pass.
