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
## M8 coverage

The mock Excel12v backend records calls, type text, IDs, reverse-order
unregistration, top-level `xlFree`, partial failure, retry, duplicate open, and
idempotent close. Thunk tests cover per-call roots, null conversion failure,
and panic fallback. The real 64-bit Excel procedure and pending result record
are in `docs/manual-tests/m8-excel-smoke-test.md`. The COM harness passed two
fresh real 64-bit Excel processes with MTR enabled; visible Function Wizard,
Add-in Manager UI, and embedded-NUL UI cases remain manual.

## M15 real-Excel stress coverage

`scripts/excel-stress-harness.ps1` separates a two-cycle smoke run from a
25-cycle/50,000-rebuild soak run. A parent PowerShell process directly tracks a
fresh worker for each cycle. The worker maps `Application.Hwnd` through
`GetWindowThreadProcessId`, immediately persists the exact Excel PID and start
time, and the parent applies a hard timeout only to that verified process pair.
Artifacts retain
workbook outputs, worker exit status/logs, timings, process handle and memory
snapshots, Excel version/build/MTR settings, and readable Windows crash-event
evidence. It covers all sample functions and command, scalars, direct UTF-16
strings, Q arrays, U references, missing/blank values, error values, controlled
fallbacks, MTR, and unload/reload. The sample has no public panic test hook;
Rust thunk tests remain the authoritative panic-fallback coverage.

Pure PowerShell assertions cover identity/start-time matching, unrelated PID
exclusion, missing coordination, cleanup selection, mode timeout defaults,
Excel-only event filtering, sample aggregation, schema fields, and preflight
classification. The current live runner fails plain `Workbooks.Add()` before
XLL registration, so M15 remains implementation-complete but live-blocked.

## M16 asynchronous UDF coverage

Registration tests require the exact `>` return plus one `X` handle and reject
stray handles and cluster-safe async functions. Macro expansion tests prove
that generated thunks return void, append the hidden handle, deep-copy only
owned input families, inject only `AsyncCancellationToken`, and reject
callback-borrowed/reference/direct-string inputs.

Controller tests use a deterministic queued executor and mock Excel backend to
verify opaque handle copying, exact two-argument `xlAsyncReturn`, Boolean
acceptance decoding, preservation of `xlretInvAsynchronousContext`, no
`xlbitDLLFree` handoff, capacity/rejection paths, cancellation, shutdown, and a
concurrent at-most-once completion race. The release XLL export test includes
the async function and both event procedures. The stress harness probes a real
async result when its Excel host can create workbooks. Real Excel cancellation,
recalculation, and unload validation remains pending because the M15 host gate
is still blocked.

## M9B coverage

Macro tests prove B/A/J/Q/U/C%/D% raw signatures derive from the same kinds as
registration metadata, exact export naming, scalar override rejection, Q/U
decoding, direct UTF-16 parsing, context injection, `Result` mapping, and both
scalar and Q panic boundaries. The minimal XLL parity suite covers all five M8
functions, values, root/nested tags, Q/U registration, fresh DLLFree-only
roots, and exact AutoFree. The release PE contains exactly the frozen 12 named
exports.

## M10 macro conformance coverage

`excel-api-macros/tests/trybuild.rs` runs compile-pass fixtures across every
supported argument, result, context, and flag family. Its checked diagnostic
snapshots reject unsupported inputs and outputs, borrowed/direct-string forms,
generics, methods/receivers, Rust `async fn`/variadic functions, ambiguous Q/U,
incompatible flags and contexts, unjustified cluster safety, invalid or
duplicate attributes, metadata mismatch, invalid exports, unsupported Result
errors, and deterministic generated-symbol collisions. Unit expansion checks
continue to assert that the same closed kind model emits matching metadata and
raw thunk ABI tokens.

## M17 xlcOnTime compatibility spike

The default minimal-XLL tests and PE inspection assert that no `RUST.ONTIME.*`
registration or research export is present. Compile-fail API examples prove
that the ordinary build has no scheduling method and that even the research
lifecycle bridge cannot be called by safe code. Separate all-feature tests
assert the checked-in `xlfNow`/`xlcOnTime` IDs, exact 2/3/4 argument
counts and tags, counted command text, missing-plus-FALSE cancellation form,
Boolean/error decoding, raw return-code preservation, context rejection,
bounded diagnostics, and absence of `xlFree` for immediate results. The ABI
checker compares both new constants against `XLCALL.H`; a separate research PE
check requires all experimental command/status exports.

`scripts/excel-ontime-validation.ps1` explicitly builds the research feature,
owns one Excel COM process, and enables
bootstrap only through a coordination marker containing that exact PID. It
records Excel build/architecture, plain and post-XLL workbook creation,
schedule/cancel/callback diagnostics, main-thread and MacroContext evidence,
bootstrap outcome, cancellation before unload, and process exit. It does not
infer bootstrap success from `RegisterXLL`. If pending cancellation fails, it
records an unsafe/inconclusive result and terminates only the PID/start-time
matched Excel process; it does not rely on `xlAutoClose` to prevent unload.
Security, recalculation, modal/edit,
copy/paste/undo, latency, and unload/reload cases remain a manual matrix.

The 2026-07-15 host could not create a plain workbook and did not enter the XLL
test bootstrap after `RegisterXLL`; COM and XLM command invocation were also
unavailable without a workbook. Therefore the live result is inconclusive and
must not be counted as `xlcOnTime` validation.

## M17 cooperative dispatcher coverage

Deterministic unit tests cover empty/FIFO/bounded selection, capacity
rejection, compatible work behind an incompatible head, expiration, detached
ticket drop, cancellation at queued/selected/running boundaries, exact-once
retirement, shutdown races, stale generations, reopen, bounded waiting, and
panic-safe nested-drain suppression. Typed-context tests pin the explicit
compatibility matrix and the Macro-only preserving `xlAbort` descriptor,
argument count, Boolean decode, no `xlFree`, and raw return-code preservation.

Runtime integration tests prove generation removal before cleanup, dispatch
remaining disabled in `CleanupRequired`, retry close, and backend unlink only
after active synchronous work has ended. Default minimal-XLL registration and
PE inspection require `RUST.DISPATCH.PUMP` while continuing to reject all
`RUST.ONTIME.*` production exports. Manual live steps are recorded in
`docs/manual-tests/m17-cooperative-dispatcher.md`; they remain pending because
the available host cannot create a plain workbook.

The M17 hardening suite additionally verifies that active waits select the
earlier caller/request deadline, tolerate spurious notifications, return
completion/cancellation/shutdown exactly, and do not apply queued expiry after
selection or Running commitment. Fault-injection hooks deterministically panic
after Running commitment and remove operation storage. Tests prove controlled
failure, exact retirement, zero pending/running counts, nonblocking shutdown,
reopen, repeated-path underflow protection, and callback-depth guard cleanup.
