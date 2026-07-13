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


# Task: implement the first real manually registered Excel function

## Objective

Complete the first production-oriented vertical slice:

- load or resolve `Excel12v`;
- register one static worksheet function manually;
- parse scalar inputs;
- invoke Rust code;
- return a safe scalar result;
- then validate string and array returns through the completed memory layer.

Do not introduce procedural-macro thunk generation yet.

## Prerequisites

Assume prompts 01 through 07 are complete.

## Required work

### 1. Excel API resolution

Choose and implement the smallest verified strategy for obtaining `Excel12v`:

- import library;
- delayed linking;
- or runtime `GetProcAddress`.

Document why the selected strategy is appropriate for the first supported
Windows x64 MSVC target.

Failures must produce clear diagnostics and must not crash Excel.

### 2. Safe/internal call wrapper

Implement a focused wrapper around `Excel12v` for:

- `xlfRegister`;
- unregistering during shutdown;
- any minimal metadata call required by registration.

Keep raw call mechanics internal.

### 3. Static registration descriptor

Use the existing `FunctionRegistration` and `AddInDescriptor`.

Extend descriptors only as needed for:

- exported procedure name;
- Excel-visible name;
- type text;
- argument names;
- category;
- descriptions;
- thread-safety flags.

Validate metadata before calling Excel.

### 4. Manual exported thunk

Create one explicit exported thunk for:

```text
RUST.ADD
```

The thunk must:

1. establish the callback lifetime;
2. parse two numeric inputs using `ExcelValueRef`;
3. call a normal Rust function;
4. convert the result using `IntoExcel`;
5. materialize and hand off the return if needed;
6. catch panics;
7. return an appropriate Excel error on failure.

For an immediate numeric return, use the safest verified ABI strategy. Do not
introduce unsound static mutable storage.

### 5. Lifecycle

Implement or complete:

- `xlAutoOpen`;
- `xlAutoClose`;
- `xlAutoFree12`;
- `xlAddInManagerInfo12`.

`xlAutoOpen` should register functions and retain registration IDs.

`xlAutoClose` should unregister and release runtime state without assuming that
all work succeeded previously.

### 6. Incremental return validation

After `RUST.ADD` works:

- add one string-return function;
- add one flat dynamic-array-return function.

These functions exist to validate the memory architecture, not to broaden the
public API.

### 7. Diagnostics

Add minimal diagnostics for:

- failure to resolve `Excel12v`;
- registration failure;
- invalid arguments;
- panic containment;
- cleanup counters under `memory-debug`.

Diagnostics must not allocate or re-enter Excel from sensitive cleanup paths.

### 8. Tests

Add:

- unit tests for registration argument construction;
- mock-call tests for `Excel12v`;
- thunk tests using synthetic raw inputs;
- panic-path tests;
- registration/unregistration state tests;
- string and array return handoff/free tests;
- Windows integration instructions for loading the built `.xll`.

### 9. Example

Update `examples/minimal-xll` into a real loadable example.

Document exact build and manual Excel loading steps for Windows x64.

## Explicit non-goals

Do not implement:

- `#[excel_function]` code generation;
- automatic distributed registration;
- async functions;
- COM/Ribbon/RTD;
- broad Excel API wrappers;
- packaging CLI.

## Acceptance criteria

The task is complete when:

1. Excel loads the XLL;
2. `xlAutoOpen` registers `RUST.ADD`;
3. the function appears in the Function Wizard;
4. valid scalar calls calculate correctly;
5. invalid inputs return controlled Excel errors;
6. deliberate Rust panics do not crash Excel;
7. a string return is repeatedly recalculated without leaks;
8. a flat array return is repeatedly recalculated without leaks;
9. `xlAutoClose` unregisters cleanly;
10. all Rust tests and lints pass;
11. manual Windows/Excel validation steps are documented.
