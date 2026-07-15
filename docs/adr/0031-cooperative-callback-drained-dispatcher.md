# ADR-0031: Cooperative callback-drained dispatcher

## Status

Accepted for M17. Live Excel validation is pending because the available host
cannot create a plain workbook.

## Decision

M17 uses a bounded queue that is drained only from a genuine typed Excel
callback. Enqueueing does not wake Excel and does not guarantee progress or
latency. The initial production entry point is `RUST.DISPATCH.PUMP` with a real
`MacroContext`; no autonomous wake source is enabled.

Each runtime open owns a fresh controller generation. Operations are sealed,
own their arguments, and declare one exact requirement: context-neutral,
thread-safe worksheet, worksheet, macro, or lifecycle. There is no public
closure receiving an Excel context and no context-free drain.

The queue is bounded by pending capacity and drain batch, with optional request
and drain-duration limits. Selection may skip incompatible requests but
preserves FIFO order within a requirement class. Locks are released before
operation execution or Excel calls. Nested drain attempts are suppressed by a
panic-safe callback-depth guard.

Tickets detach on drop. Cancellation is effective before execution commitment;
terminal publication and registry retirement are exact-once. Shutdown removes
the generation, rejects enqueue, retires queued and selected work, waits for
already-running synchronous work, and permits unlink only afterward. A failed
unregister leaves the runtime in `CleanupRequired` with dispatch disabled.

## Consequences

Manual pumping is observable and safe but cannot guarantee progress. A future
RTD, COM, or otherwise verified notification adapter may request an Excel
callback without changing dispatcher semantics. ADR-0030 remains unchanged:
`xlcOnTime` is inconclusive and not approved.
