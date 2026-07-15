# Asynchronous UDF Architecture

## Status

M16 implementation in progress. The automated contract can be completed on
ordinary CI; real-Excel cancellation, recalculation, and unload validation is
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

Each request has one atomic state:

```text
Scheduled -> Completing -> Completed
    |             |
    +----------> Canceled
```

Only a successful compare-and-exchange from `Scheduled` to `Completing` may
call `xlAsyncReturn`. Cancellation or shutdown changes `Scheduled` to
`Canceled`. A controller epoch and active flag reject stale work after close
and after a later reopen. The controller disables completion before lifecycle
unregistration and before the Excel callback backend is unlinked.

Calculation cancellation marks every current request canceled. Calculation
ended removes terminal request records. XLL close performs the same stop gate,
then drains/cancels tracked requests and waits for the executor's shutdown
contract. Executor implementations must make accepted work cooperative and
bounded so unload cannot wait indefinitely.

## Executor boundary and bounds

The crate defines an `AsyncExecutor` adapter accepting boxed `Send + 'static`
jobs. It does not depend on or initialize Tokio, async-std, or another runtime.
An XLL installs one executor explicitly and configures a maximum number of
in-flight requests. Capacity is reserved before submitting a job and released
on every completion, cancellation, panic, or executor rejection path.

Worker panics are contained before the FFI completion boundary and mapped to
Excel `#VALUE!` when completion is still legal. Submission failure and capacity
exhaustion are observable diagnostics; the initial void thunk cannot return an
ordinary worksheet error synchronously.

## Validation status

Unit, race, macro expansion, ABI, and lifecycle tests cover registration text,
deep-copy restrictions, exact callback arguments, at-most-once completion,
cancellation, stale epochs, shutdown, panic containment, and bounded
submission. Real Excel still must cover recalc, cancellation, close/unload,
and late-result suppression before M16 is stable.

## Authoritative sources

- [Asynchronous user-defined functions](https://learn.microsoft.com/en-us/office/client-developer/excel/asynchronous-user-defined-functions)
- [xlAsyncReturn](https://learn.microsoft.com/en-us/office/client-developer/excel/xlasyncreturn)
- [xlEventRegister](https://learn.microsoft.com/en-us/office/client-developer/excel/xleventregister)
- [xlfRegister (Form 1)](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1)
