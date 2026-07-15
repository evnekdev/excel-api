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

## Lifecycle-race amendment

Each open consumes a fresh executor/controller generation. Shutdown removes
the active `Arc` generation, marks it inactive, permanently closes its executor
and joins every accepted closure before Excel registrations or the callback
backend can disappear. The built-in executor is linearizable under one state
mutex (`New -> Running -> ShuttingDown -> Closed`) and cannot restart.

Request admission reserves capacity and inserts the request under a controller
mutex. A second check holds that short mutex only while the executor commits
acceptance, and shutdown uses the same transition. Shutdown therefore either
precedes submission (which returns the unexecuted job) or follows acceptance
(which is joined). Executors enqueue and return rather than synchronously
running the closure under this commit check. A worker must transition
`Scheduled -> Running` while its generation is active before user code runs.
Cancellation while queued skips user code, running cancellation is exposed by
the token, and one idempotent retirement operation owns registry removal and
capacity release.

Microsoft documents event registration and says CalculationCanceled is
immediately followed by CalculationEnded, but the verified pages provide no
event-unregister function or repeated-registration contract. The production
runtime registers each event at most once per loaded binary and retries only a
registration that did not succeed. Event callbacks are no-ops without an
active generation.

An Excel unregister failure produces `RuntimePhase::CleanupRequired`, not
`Initialized`. Async work stays disabled; retry-close is permitted and reopen
requires successful cleanup plus a fresh executor.

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

Third-party executor conformance checklist: accept a job exactly once or return
it unexecuted; make shutdown irreversible and idempotent; wait until accepted
closures cannot execute; never detach work past XLL unload; do not require
shutdown from inside a job unless explicitly supported; never synchronously run
or wait for the submitted closure inside `execute`.
