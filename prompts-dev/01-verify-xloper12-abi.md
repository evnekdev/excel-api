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


# Task: verify and complete the initial Excel C API ABI layer

## Objective

Make `excel-api-sys` an authoritative, testable representation of the supported
subset of the Excel 12 C API ABI for Windows x64 MSVC.

This task is about layout correctness only. Do not implement safe value parsing,
return allocation, function registration, or Excel runtime calls beyond raw
signatures.

## Required work

### 1. Audit raw types

Inspect the current definitions and verify or correct:

- `XLOPER12`;
- its tagged union;
- `XLREF12`;
- single-reference and multi-reference structures;
- rectangular multi-cell storage structures;
- `FP12` or the exact modern floating-point array equivalent;
- UTF-16 string pointer representation;
- integer widths and signedness;
- row and column dimensions;
- function-pointer calling conventions.

Only retain fields and structures that can be verified.

### 2. Audit constants

Verify or correct:

- base `xltype*` values;
- ownership bits;
- error discriminants;
- relevant return codes;
- maximum row and column constants;
- initial function identifiers needed by later registration work.

Keep raw constants named consistently with the SDK where practical.

### 3. Verify function signatures

Define exact raw signatures for:

- `Excel12`;
- `Excel12v`;
- `xlAutoOpen`;
- `xlAutoClose`;
- `xlAutoFree12`;
- `xlAddInManagerInfo12`.

Do not yet choose the final linking/loading strategy for `Excel12v`; represent
the ABI only.

### 4. Add ABI verification

Create an ABI verification mechanism that compares Rust and C definitions.

Preferred structure:

```text
tools/abi-check/
  Cargo.toml
  build.rs
  src/main.rs
  native/layout_check.c
```

The C helper should compile against an authoritative local Excel SDK header if
available and expose or print:

- `sizeof`;
- alignment;
- selected field offsets;
- selected enum/constant values.

Rust tests or the tool should compare those values against the Rust
definitions.

If the official header is unavailable in the environment:

- make the ABI checker conditional;
- keep ordinary workspace tests working;
- document exactly how a developer with the SDK can run the comparison;
- do not fabricate expected values.

### 5. Add compile-time/runtime checks

Add tests for:

- size and alignment;
- offset-zero requirements where relevant;
- union capacity;
- constant values;
- calling-convention type compatibility where testable.

## Required design decisions

Document:

- supported target: Windows x64 MSVC first;
- whether bindings are curated, generated once, or generated during builds;
- provenance of every copied constant/layout;
- unsupported structures and why they remain deferred.

## Explicit non-goals

Do not implement:

- `ExcelValueRef`;
- UTF-16 decoding;
- array iteration;
- `ExcelReturn`;
- `xlAutoFree12` logic;
- `Excel12v` invocation;
- `xlfRegister`;
- proc-macro thunk generation.

## Acceptance criteria

The task is complete when:

1. all supported raw definitions are verified or clearly marked as pending;
2. the workspace compiles on supported targets;
3. ABI tests/check tooling exist;
4. no placeholder field is presented as authoritative;
5. all unsafe ABI assumptions are documented;
6. ordinary tests pass without requiring the Excel SDK;
7. an SDK-enabled verification path is documented and reproducible.
