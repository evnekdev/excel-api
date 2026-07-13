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


# Task: implement safe logical return planning

## Objective

Introduce a safe, two-phase return pipeline that validates all semantic and size
requirements before any raw pointer is created or published.

## Prerequisites

Assume owned semantic values and conversions from prompt 03 exist.

## Required work

### 1. Logical return type

Introduce an internal or carefully scoped type:

```rust
pub(crate) enum ExcelReturnValue {
    Number(f64),
    Boolean(bool),
    Error(ExcelError),
    Empty,
    Text(String),
    Array(ReturnArray),
}
```

Exclude states that are not valid outputs.

Do not reuse every `ExcelValue` variant automatically.

### 2. `IntoExcel`

Refine the trait so it produces a logical return, not raw ABI memory:

```rust
pub trait IntoExcel {
    fn into_excel_value(self)
        -> Result<ExcelReturnValue, ConversionError>;
}
```

Support initial Rust outputs:

- `f64`;
- `bool`;
- supported integers;
- `String`;
- `&str`;
- `ExcelError`;
- `Result<T, E>` according to documented policy;
- flat rectangular arrays.

### 3. Return plan

Implement a safe `ReturnPlan` that computes and validates:

- root type;
- string buffer counts and code-unit lengths;
- array element count;
- total backing allocation count;
- aggregate byte estimate;
- alignment requirements if needed;
- recursion depth;
- unsupported nested structures.

All arithmetic must be checked.

### 4. No pointer publication

The planning phase must:

- allocate no ABI backing buffers;
- create no pointers into future storage;
- mutate no ownership bits;
- be testable using safe Rust only.

### 5. Limits

Apply explicit limits for:

- maximum Excel string length;
- maximum array elements;
- maximum aggregate return bytes;
- maximum nesting depth.

### 6. Error model

Introduce `ReturnError` or equivalent with precise variants for:

- unsupported return type;
- string too long;
- array too large;
- size overflow;
- nested array unsupported;
- invalid shape;
- aggregate limit exceeded.

### 7. Tests

Add tests for:

- scalar plans;
- empty and maximum accepted strings;
- string-too-long rejection;
- valid mixed flat arrays;
- nested-array rejection;
- checked multiplication/addition overflow;
- aggregate byte limits;
- deterministic plan output;
- zero raw pointers in plan structures.

## Explicit non-goals

Do not implement:

- `ReturnAllocation`;
- raw `XLOPER12` pointer patching;
- DLL-free ownership bits;
- `xlAutoFree12`;
- global allocation registries.

## Acceptance criteria

The task is complete when:

1. supported Rust outputs convert to a validated logical return;
2. every required size is checked before materialization;
3. unsupported structures fail before allocation;
4. the planner contains no unsafe code;
5. the planner publishes no raw pointers;
6. tests cover all size and overflow boundaries;
7. all tests and lints pass.
