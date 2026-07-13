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


# Task: complete owned Excel semantic values and conversions

## Objective

Provide fully owned Rust representations that are independent of Excel memory
and safe to retain, cache, clone, or move to worker threads.

## Prerequisites

Assume borrowed callback views from prompt 02 exist and are tested.

## Required work

### 1. Owned value model

Refine:

```rust
pub enum ExcelValue {
    Number(f64),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(String),
    Array(ExcelArray),
}
```

Ensure no variant contains a raw Excel pointer.

### 2. Owned rectangular arrays

Implement `ExcelArray` with:

- checked `rows * columns`;
- exact shape invariant;
- boxed or vector-owned elements;
- rows, columns, len, is_empty;
- checked indexing;
- row/column access where useful;
- iteration;
- clear row-major semantics;
- no jagged representation.

### 3. Deep-copy conversions

Implement deep conversion from:

- `ExcelStr<'call>` to `String`;
- `ExcelArrayView<'call>` to `ExcelArray`;
- `ExcelValueRef<'call>` to `ExcelValue`.

Add recursion and aggregate-size limits.

Nested arrays should remain rejected unless the architecture explicitly changes.

### 4. Conversion traits

Refine:

```rust
pub trait FromExcel<'call>: Sized {
    fn from_excel(value: ExcelValueRef<'call>)
        -> Result<Self, ConversionError>;
}
```

Implement checked conversions for:

- `f64`;
- `bool`;
- `String`;
- integer types supported by project policy;
- `ExcelError`;
- `Option<T>`;
- `OptionalValue<T>`;
- `ExcelValue`;
- owned arrays where appropriate.

Integer conversion must reject:

- fractional values;
- NaN;
- infinity;
- out-of-range values.

### 5. Missing and empty semantics

Keep exact distinction available through `OptionalValue<T>`.

Document the default `Option<T>` policy explicitly.

### 6. Conversion limits

Introduce a configuration or internal limits type covering:

- maximum string units;
- maximum array elements;
- maximum recursion depth;
- maximum aggregate copied bytes.

Defaults should be conservative and documented.

### 7. Error model

Make conversion failures precise enough to distinguish:

- unexpected type;
- unsupported type;
- invalid UTF-16;
- invalid dimensions;
- size limit exceeded;
- integer overflow/range;
- nested-array rejection.

### 8. Tests

Add tests for:

- deep-copy independence from source storage;
- every primitive conversion;
- NaN/infinity integer rejection;
- missing/empty behavior;
- exact array shape;
- aggregate-size limits;
- recursion limits;
- strict and lossy string conversion;
- worker-thread movement of owned values where expected.

## Explicit non-goals

Do not implement:

- raw ABI return allocation;
- `ExcelReturn`;
- `xlAutoFree12`;
- `Excel12v`;
- registration.

## Acceptance criteria

The task is complete when:

1. owned values contain no Excel memory;
2. borrowed inputs can be deeply copied safely;
3. conversion limits prevent pathological allocation requests;
4. integer and string conversions are checked;
5. exact missing/empty behavior is documented and tested;
6. owned values can outlive callbacks safely;
7. all tests and lints pass.
