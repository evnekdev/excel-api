# Testing Architecture

## Test layers

### Pure Rust unit tests

- conversion;
- shape validation;
- string parsing;
- return planning;
- state machines.

### ABI tests

- C/Rust size/alignment/offset comparison;
- constants and type text;
- calling conventions.

### Compile-fail tests

- unsupported macro signatures;
- illegal flag combinations;
- escaping borrowed lifetimes;
- non-Send callback values.

### Mock Excel backend

Injectable call table for:

- `Excel12v`;
- `xlFree`;
- registration;
- call errors;
- ownership transfer.

### Memory tests

- partial-failure cleanup;
- exactly-once handoff/free;
- Excel-owned copy/release;
- XLFree transfer;
- nested string arrays;
- no live allocations.

### Real Excel smoke tests

- load/unload;
- registration;
- scalar/string/array functions;
- reference/value-only inputs;
- MTR;
- Function Wizard;
- workbook close/cancel;
- repeated recalculation.

### Stress/fuzz

- malformed tags;
- invalid lengths/dimensions;
- repeated allocation;
- large arrays;
- concurrency.

The borrowed-value suite currently includes a deterministic malformed-xltype
regression loop, but no dedicated cargo-fuzz target. A coverage-guided Prompt
02 fuzz target remains deferred testing work; Prompt 03 does not mix that
harness setup into owned-value implementation.

The M4 pure-Rust suite covers every scalar/error return, UTF-8 and arbitrary
UTF-16 planning, Excel/project string boundaries, flat arrays and ordering,
zero dimensions, ABI dimensions, checked accounting overflow, byte/allocation/
element/depth limits, deterministic metadata, and natural `Send + Sync +
'static` behavior. ABI materialization and exactly-once free tests remain M5-M6.

## Historical book guidance

The book's sample code is valuable for behavior and pitfalls, but the project
does not copy its ownership flexibility blindly. Tests enforce the stricter
Rust invariants chosen here.
