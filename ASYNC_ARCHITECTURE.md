# Asynchronous UDF Architecture

## Status

M16 implementation and lifecycle-race hardening are complete in automated
tests. Real-Excel cancellation, recalculation, and unload validation is still
required before this design is stable.

M16 started by explicit maintainer direction while the M15 live-smoke gate is
still blocked by an intermittent host COM `Workbooks.Add` failure. This does
not change M15's status and is a recorded validation risk, not evidence that
the stress harness or async behavior has passed live Excel.

## Verified Excel contract

An asynchronous worksheet function registers a void return (`>`) and exactly
one `X` argument. Excel passes that argument as an `xltypeBigData` asynchronous
handle. The entry thunk returns without a result; a later operation calls
`xlAsyncReturn(handle, result)`. Inputs needed after thunk return must be deep
copied. `xlAsyncReturn` is the only Excel C API callback permitted from the
non-calculation worker thread during recalculation, and the XLL retains and
frees its result allocation after that synchronous callback.

Excel exposes `CalculationCanceled` and `CalculationEnded` through
`xlEventRegister`. Cancellation asks the XLL to stop asynchronous activity;
ended follows cancellation and permits calculation-scoped resource cleanup.
The event procedures are parameterless exported procedures. No general Excel
callback capability is made available to either event handlers or workers.

Microsoft documents registration, but the verified sources document neither an
event-removal function nor repeated-registration replacement semantics.
Consequently each event is registered at most once per loaded production
`Runtime`; successful individual registrations survive a later initialization
failure, missing registrations are retried, and callbacks are safe no-ops when
there is no active generation.

## Ownership boundary

`AsyncHandle` copies only the opaque `xltypeBigData` handle fields supplied by
Excel. It never dereferences or frees the handle value. The copied handle can
cross threads solely so the exact bits can be reconstructed for one
`xlAsyncReturn` call.

Macro-generated async thunks decode and deep-copy every Excel-visible input
before returning. Borrowed arguments, reference-preserving arguments, direct
UTF-16 views, and callback contexts are rejected for async functions. Worker
jobs therefore contain only `Send + 'static` Rust values and an owned request.

Return planning and materialization occur off the calculation thread. The
materialized allocation remains locally owned across `xlAsyncReturn`; it is
not marked `xlbitDLLFree` and is dropped by Rust immediately after the callback
returns. `xlAutoFree12` is never involved in an async result.

## State machine and shutdown

Each request has one atomic execution state and one idempotent retirement bit:

```text
Scheduled -> Running -> Completing -> Completed
    |            |             |
    |            +-> CancelRequested -> Canceled
    +--------------------------------> Canceled
```

The worker checks the active generation and wins `Scheduled -> Running` before
calling user code. Thus cancellation while queued retires the request without
calling the function body. Only `Running -> Completing` may call
`xlAsyncReturn`; a cancellation that wins first becomes `CancelRequested` and
suppresses an ignored result. Retirement atomically removes the registry entry
and decrements capacity at most once on every terminal path.

An open consumes a fresh controller and executor generation. Shutdown removes
the active generation before draining it. Scheduling clones a stable `Arc`, so
stale jobs retain only their old controller. The old controller is inactive
and its executor permanently closed before runtime unlink. A new `xlAutoOpen`
must install a new executor.

Calculation cancellation marks every current request canceled. Calculation
ended removes terminal request records. XLL close performs the same stop gate,
then drains/cancels tracked requests and waits for the executor's shutdown
contract. Executor implementations must make accepted work cooperative and
bounded so unload cannot wait indefinitely.

## Executor boundary and bounds

The crate defines an `AsyncExecutor` adapter accepting boxed `Send + 'static`
jobs. A rejection returns ownership of the closure, proving that it was not
run. `shutdown` is an irreversible rejection boundary and must wait until all
accepted closures can no longer execute XLL code. It is idempotent and must not
detach work. Third-party adapters use one executor per open generation.

The built-in executor is `New -> Running -> ShuttingDown -> Closed`. `execute`
and the shutdown transition share one mutex, so queue acceptance and permanent
closure are linearizable. Shutdown drops the sole sender and joins every
worker without holding an executor/controller/global lock. Partial worker
startup closes the queue, joins workers already created, and leaves the
executor permanently closed. Queue-full, closed, and worker-start failures are
distinct.

Controller admission reserves capacity and registers the request, then takes a
second short commit check under the lifecycle mutex while the executor records
acceptance. Shutdown uses that same mutex. A request paused between registration
and submission is rejected if shutdown wins; an accepted request precedes the
shutdown boundary and is canceled and joined. Workers cannot pass their start
check until the commit mutex is released, and no gate is held while user code,
Excel, or a join runs. Executors must enqueue rather than synchronously run or
wait for a submitted closure inside `execute`.

Worker panics are contained before the FFI completion boundary and mapped to
Excel `#VALUE!` when completion is still legal. Submission failure and capacity
exhaustion are observable diagnostics; the initial void thunk cannot return an
ordinary worksheet error synchronously.

If Excel unregistration fails during close, the runtime enters
`CleanupRequired`: async work remains disabled, the backend stays linked only
for cleanup, initialize is rejected, and retry-close is permitted. It never
reports `Initialized` with a drained async generation.

## Validation status

Unit, race, macro expansion, ABI, and lifecycle tests cover registration text,
deep-copy restrictions, exact callback arguments, at-most-once completion,
cancellation, stale epochs, shutdown, panic containment, and bounded
submission. Real Excel still must cover recalc, cancellation, close/unload,
and late-result suppression before M16 is stable.

ThreadSanitizer is not available for the Windows/MSVC target used by this
repository. Deterministic barriers/channels cover the critical boundaries;
the project does not claim TSan or Loom validation.

## Authoritative sources

- [Asynchronous user-defined functions](https://learn.microsoft.com/en-us/office/client-developer/excel/asynchronous-user-defined-functions)
- [xlAsyncReturn](https://learn.microsoft.com/en-us/office/client-developer/excel/xlasyncreturn)
- [xlEventRegister](https://learn.microsoft.com/en-us/office/client-developer/excel/xleventregister)
- [Handling Events](https://learn.microsoft.com/en-us/office/client-developer/excel/handling-events)
- [xlfRegister (Form 1)](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1)
# Relationship to M17 dispatch

The M17 cooperative dispatcher has its own replaceable controller generation,
queue, tickets, and shutdown gate. It does not share async handles or the M16
executor/controller and does not use `xlAsyncReturn`. Both systems follow the
same rule that their generation is removed and drained before backend unlink.
