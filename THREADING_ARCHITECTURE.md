# Threading Architecture

## Excel threading domains

- Excel UI/main thread;
- Excel multi-threaded recalculation workers;
- XLL-created worker threads;
- async completion workers;
- future dispatcher thread interactions.

The non-default, feature-gated M17 `xlcOnTime` probe is initiated only from
Excel callback contexts and never calls Excel from a background thread. Its registered
callback records the Windows thread ID and also performs a verified
MacroContext `xlAbort` poll; thread identity alone is not considered proof of
legal context. No dispatcher worker or queue is implemented by the spike.

## Thread-safe UDF rules

A function marked thread-safe:

- cannot also be macro-sheet equivalent;
- must not use thread-unsafe C API operations;
- must not use shared mutable static return storage;
- must protect shared resources;
- must not assume callback order.

## Return memory

Initial strategy:

- allocate one return root per call;
- free through `xlAutoFree12`;
- avoid TLS return slots.

This is easier to reason about than persistent thread-local return objects.

M6 implements this strategy without thread-local ownership. The handed-off
allocation contains `Box<[XLOPER12]>`, `Box<[XCHAR]>`, ordinary scalar
metadata, and test-only atomic tracking. These fields use the process allocator
and have no thread affinity, locks, COM state, workbook state, or creator-thread
dependency, so the matching callback may drop them on a thread different from
construction. A cross-thread reclamation test exercises that contract.

`ExcelReturn` is not given manual `Send` or `Sync` implementations. The
required boundary is the documented ownership transfer of its unique raw
pointer to Excel, not a blanket promise that the local wrapper is shareable.
`xlAutoFree12` makes no Excel calls and acquires no project locks.

## C API calls

Thread-safe contexts expose only verified thread-safe operations.

The book identifies C API-only operations that are generally thread-safe, while
many worksheet/XLM calls are not. The safe wrapper must encode a whitelist,
not an optimistic default.

## XLL-created threads

A background thread must not call the Excel C API or COM directly unless a
specific documented mechanism permits it.

Worker threads receive owned Rust data only.

M16 makes `xlAsyncReturn` the sole exception to the general background-call
ban. Generated async thunks deep-copy supported inputs before returning and
move only `Send + 'static` values, an opaque copied Excel handle, and an
`AsyncCancellationToken` to the installed executor. No worksheet, lifecycle,
macro, COM, or general C API capability crosses the thread boundary.

The optional standard-library executor has a bounded queue, fixed worker count,
and permanent `New -> Running -> ShuttingDown -> Closed` state. Its acceptance
and shutdown boundary share one lock. Closed executors cannot restart, partial
worker startup is joined, and no lock is held while worker threads are joined.

Applications may install another `AsyncExecutor`, one per XLL open generation.
`execute` must accept exactly once or return the unexecuted job. `shutdown`
must establish an irreversible rejection boundary, be idempotent, join/drain
all accepted jobs, and never detach XLL code past unload. The runtime does not
invoke shutdown from an executor job.

M7's `ExcelOwnedValue<'call>` is deliberately neither `Send` nor `Sync`. Its
boxed root is stable, but the ability to call back into Excel is scoped to the
originating callback. Microsoft documents `xlFree` as thread-safe during MTR;
that does not make calls from arbitrary XLL-created threads legal. Convert to
ordinary `ExcelValue` before worker-thread use.

## Shutdown

Lifecycle shutdown must account for:

- outstanding workers;
- pending callbacks;
- Excel closing/cancelling;
- late cleanup calls;
- cancelled close sequences.

M16 close removes and disables the active controller generation, cancels
scheduled/running requests, and
calls the executor's blocking shutdown contract. It unregisters functions and
unlinks the Excel callback only after no accepted executor job can execute XLL
code. Queued cancellation skips the user body. A stale generation or inactive
controller suppresses late `xlAsyncReturn`. Failed Excel cleanup enters
`CleanupRequired` with async scheduling still disabled and the backend linked
only for retry.

## Shared state

Prefer:

- immutable descriptors;
- atomics for counters/state;
- scoped locks;
- no lock in `xlAutoFree12` if avoidable.
## M8 classifications

`RUST.ADD`, `RUST.ECHO`, `RUST.ARRAY.ECHO`, and `RUST.OPTION.KIND` are pure and
registered `$`; each produces fresh per-call DLLFree storage. The
reference-preserving probe is deliberately not `$`. Registration runs only in
lifecycle context. Panic and catastrophic static scalar fallbacks are
immutable, pointer-free, and contain no ownership bit.

## M9B generated callback contexts

Generated thunks create callback-scoped context capabilities from the same
process-local production backend used by lifecycle `Runtime`. The backend
entry is atomic, the per-call capability is borrowed, and no runtime mutex is
held while the ordinary function executes. Generated returns retain no context
or callback borrow. Direct scalar failures and panics return zero/false; Q
failures use immutable pointer-free Excel error roots.

## M11 cancellation polling

`xlAbort` is a verified C API-only call and may be polled from the documented
worksheet, thread-safe, and macro callback capabilities. Microsoft documents
the caveat that a thread-safe UDF cannot clear a break condition; callers use
an explicit preserve/clear mode and receive the exact Excel return code if it
is rejected. This reports a user break request only, never application
calculation progress or state.
# M17 cooperative dispatch

Background producers may enqueue owned dispatcher operations and use owned
tickets, but they never receive an Excel context and never call Excel. Excel
work occurs only in a later typed callback drain. A callback-depth guard rejects
ticket waits from callback scopes and suppresses recursive drain attempts.
Queue/controller locks are released before user-independent operation execution
or Excel calls. Shutdown waits for synchronously running operations before the
backend may be unlinked.

Ticket waits use a spurious-wakeup-safe loop and observe the earlier applicable
caller/request deadline without a timer thread. Request expiry applies only
while `Queued`. Every Running request is owned by an RAII guard that releases
running accounting and signals shutdown waiters even when execution unwinds or
an internal invariant fails.

# M18 RTD COM boundary

RTD methods are COM calls governed by the server's registered apartment model
and COM marshaling. Microsoft does not promise that every method runs on
Excel's main thread, and an observed thread ID would not establish legal
Excel12/Excel12v use. The proposed RTD server therefore creates no Excel C API
context and makes no XLL backend call.

Producer threads exchange only bounded, owned topic values. A raw
`IRTDUpdateEvent` pointer never crosses apartments; a future notifier must use
a COM proxy obtained through standard marshaling or the Global Interface
Table. Shutdown stops and joins producers, prevents new notifications, drains
already-committed notification calls, revokes callback access, and releases
COM references before termination.
