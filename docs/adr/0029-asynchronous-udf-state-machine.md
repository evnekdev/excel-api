# ADR-0029: Owned asynchronous UDF state machine

## Decision

Represent an Excel asynchronous call as an owned opaque handle plus an atomic
request state managed by a bounded controller. Macro-generated `asynchronous`
thunks use the documented `>...X` ABI, deep-copy all supported inputs during
the callback, and submit a `Send + 'static` job through an application-installed
`AsyncExecutor`.

Only the request transition from scheduled to completing can issue
`xlAsyncReturn`, and that narrow operation is the only Excel callback exposed
to worker threads. Its result allocation remains Rust-owned throughout the
synchronous callback and is dropped afterward. Cancellation, calculation end,
and XLL shutdown invalidate current requests; shutdown disables completion
before registration cleanup and backend unlink.

The runtime does not select or start a third-party async runtime. Executor
installation and the in-flight bound are explicit application policy.

## Consequences

Callback borrows and ordinary callback capabilities cannot escape into worker
state. Duplicate, canceled, stale-generation, and post-unload completions are
locally rejected. Executor rejection, capacity exhaustion, callback failure,
and invalid Excel async handles remain distinguishable diagnostics.

The event-handler ABI is confined to two parameterless exports registered by
name with `xlEventRegister`; handlers only mutate the async controller and do
not call Excel. Real Excel validation remains necessary for cancellation,
recalculation, and unload timing.

M16 began under explicit maintainer direction despite M15's blocked live-smoke
gate. That exception does not upgrade M15's status and leaves live integration
risk for both milestones.
