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


# Task: implement borrowed Excel input views

## Objective

Implement a safe, callback-lifetime-bound view over raw `XLOPER12` arguments
received from Excel.

This task must not allocate for ordinary borrowed inspection and must never free
callback inputs.

## Prerequisites

Assume prompt 01 has completed and the supported raw ABI subset is verified.

## Required work

### 1. Internal raw wrapper

Introduce an internal type similar to:

```rust
pub(crate) struct RawExcelValue<'call> {
    ptr: NonNull<XLOPER12>,
    _lifetime: PhantomData<&'call XLOPER12>,
}
```

Construction must be internal and unsafe, with documented requirements for:

- non-null pointer;
- alignment;
- callback lifetime;
- tag/union correspondence;
- ownership domain;
- valid Excel-provided memory.

### 2. Public borrowed value model

Implement or refine:

```rust
pub enum ExcelValueRef<'call> {
    Number(f64),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(ExcelStr<'call>),
    Array(ExcelArrayView<'call>),
    Reference(ExcelReference<'call>),
}
```

If references are not yet verified, keep them internal or return an explicit
unsupported-value error rather than guessing.

### 3. Centralized tag decoding

Add one internal decoding path that:

- masks ownership bits before matching the base type;
- rejects contradictory or unsupported flags;
- accesses only the union field corresponding to the verified tag;
- never frees callback values;
- converts malformed or unsupported values into controlled errors.

No other module should read raw union fields directly.

### 4. Borrowed UTF-16 strings

Implement `ExcelStr<'call>`:

- borrow the payload code units without copying;
- validate the verified length-prefix representation;
- permit interior NUL code units;
- expose `as_units`;
- provide strict `to_string`;
- optionally provide explicit lossy conversion;
- reject invalid lengths and malformed UTF-16 safely.

### 5. Borrowed arrays

Implement `ExcelArrayView<'call>`:

- validate signed dimensions before conversion;
- use checked multiplication;
- reject null pointers for non-empty arrays;
- expose rows, columns, length, checked indexing, and iteration;
- yield element conversion errors rather than constructing invalid references;
- reject nested arrays initially unless already verified and intentionally
  supported.

### 6. Threading traits

Ensure borrowed views are not accidentally `Send` or `Sync` by default.

Use a sound marker strategy rather than relying on comments.

### 7. Tests

Add tests for:

- each supported primitive tag;
- ownership-bit masking;
- missing and empty distinction;
- invalid/unsupported tags;
- null and zero-length string behavior;
- strict UTF-16 conversion;
- invalid UTF-16;
- negative dimensions;
- multiplication overflow;
- null array payload;
- indexing and row-major iteration;
- nested array rejection;
- compile-time lifetime behavior where practical.

Consider property tests or fuzz-friendly parser entry points, but do not add a
heavy framework unless justified.

## Explicit non-goals

Do not implement:

- owned deep-copy arrays beyond what tests require;
- `ExcelReturn`;
- `xlAutoFree12`;
- Excel-owned API result cleanup;
- async retention;
- actual exported thunks.

## Acceptance criteria

The task is complete when:

1. raw callback inputs can be parsed through one audited internal path;
2. borrowed views cannot safely outlive the callback;
3. borrowed views do not free memory;
4. strings and flat arrays can be inspected safely;
5. malformed supported inputs return errors;
6. borrowed wrappers are not `Send` or `Sync`;
7. all tests and lints pass.
