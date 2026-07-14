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

The M5 suite verifies every scalar tag/member, all Excel errors, counted-string
prefixes and arbitrary UTF-16, maximum strings, one contiguous row-major multi,
deep nested string pointers, root/element/string address stability after moves,
all plan totals, offset-zero root layout, and absence of ownership bits. Test-
only failure injection covers every construction stage. Atomic live root,
string, and element-buffer counters return to zero after failures, normal drop,
and 1,000 repeated construction/drop cycles.

The M6 suite verifies consuming handoff for numbers, integers, Booleans, every
Excel error, missing, empty, text, and mixed multis. It proves root/allocation
pointer identity, offset-zero layout, unchanged base types and nested pointers,
root-only DLLFree, absence of XLFree, base-only nested tags, embedded NUL,
unpaired surrogates, maximum strings, scalar heap-root cleanup, movement before
handoff, cross-thread callback cleanup, null tolerance, and exact callback ABI
typing. Test-only atomics distinguish live backing storage, outstanding
handed-off roots, and cumulative callback frees across 1,000 handoff/callback
cycles.

Compile-fail docs prove that handoff consumes `ExcelReturn`, so the same owner
cannot be handed off twice, and that callback reclamation requires `unsafe`.
Tests deliberately do not call the callback twice on one pointer because that
would be use-after-free, not a recoverable behavior test. A test-only panic
hook before reclamation verifies that panic does not cross the extern wrapper
and that the still-valid allocation is then reclaimed. Production destruction
is designed not to panic; recovery from a panic during arbitrary partial
destruction is not promised.

## Historical book guidance

The book's sample code is valuable for behavior and pitfalls, but the project
does not copy its ownership flexibility blindly. Tests enforce the stricter
Rust invariants chosen here.

The M7 mock-backend suite covers explicit release, Drop fallback, no-release
scalars, bit-masked borrowing, lossless UTF-16 copy, mixed multi copy,
top-level-only release, malformed conversion, combined conversion/release
failures, exact Excel codes, invalid context, not-thread-safe and unavailable
backends, contained backend/conversion panics, DLLFree absence, and 1,000
exactly-once cycles. The pre-commit transfer token is non-duplicable, performs
no premature release or ownership-bit mutation, exposes no pointer, and falls
back to release on Drop. Tests never call a live Excel process.
