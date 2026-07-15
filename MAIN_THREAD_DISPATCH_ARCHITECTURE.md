# Main-Thread Dispatch Architecture

## Current decision

M17 is implemented as a bounded, cooperative, callback-drained dispatcher.
**Enqueueing a request does not wake Excel.** Progress requires a later genuine
Excel-issued callback, initially the explicit `RUST.DISPATCH.PUMP` command.
There is no timer, hidden window, `PostMessage`, COM, RTD, or production
`xlcOnTime` wake source.

The dispatcher is deliberately independent of notification. A future verified
adapter may cause Excel to issue a callback, but it does not change request
ownership, tickets, capability matching, execution, retirement, or shutdown.

## Cooperative dispatcher contract

Each successful runtime-open cycle installs a fresh replaceable controller
generation. Producers retain a stable generation handle; stale handles reject
enqueue and can never target a later generation. Close removes global access,
commits shutdown under the controller lock, retires queued/selected requests,
waits for synchronous running work, unregisters, then unlinks the backend.

The sealed operation catalogue initially contains owned `EchoOwned` and a
Macro-only preserving `xlAbort` cancellation poll. It accepts no arbitrary
context closure. Compatibility is explicit: context-neutral work may run in
any typed drain; thread-safe, worksheet, macro, and lifecycle requirements run
only in the identically typed callback. Lifecycle close performs retirement,
not an ordinary drain.

Configuration bounds pending requests, batch size, optional request lifetime,
and optional drain duration. Selection scans FIFO order without letting an
incompatible head block compatible work. A selected request is removed under
the queue lock, then executed after unlocking. New work enqueued during a drain
waits for a later callback. A thread-local RAII depth guard suppresses nested
drains and rejects ticket waits from callback scopes.

Tickets own only Rust state and may be polled, timeout-waited off callback
threads, or canceled. Dropping a ticket detaches it; the controller retains and
retires the request. Cancellation wins until the Selected-to-Running
commitment. One idempotent retirement path publishes a terminal result,
releases registry capacity, and signals waiters at most once.

M17 has not selected an autonomous wake mechanism. A research-only
`xlcOnTime` compatibility probe exists, but the decision is **inconclusive**
until its complete contract, security behavior, cancellation, and unload
behavior are reproduced on a working current Excel host.

The production fallback remains a manually pumped cooperative dispatcher. The
research code does not create a dispatcher queue and is not a public arbitrary
XLM-command API.

## Historical evidence and verified boundary

Steve Dalton, sections 9.10.1 and 9.11.9, describes `xlcOnTime` as a polling
bridge that asks Excel to invoke a registered XLL command. Checked-in
`XLCALL.H` confirms only the IDs:

- `xlfNow = 74`;
- `xlcOnTime = 148 | xlCommand = 32916 = 0x8094`.

Microsoft's current C API documentation confirms that a registered XLL command
called by Excel is a class-3 context and can make command-equivalent calls. It
also states that Excel12v cannot be called from a background thread or an
operating-system timer callback. Current Microsoft documentation does not
publish the modern `xlcOnTime` argument/result/cancellation contract. The
similar VBA `Application.OnTime` contract is corroborating evidence, not proof
of the C API command.

## Experimental surface

The typed spike accepts only:

- two-argument schedule: Excel serial time and exact registered command name;
- schedule with a latest Excel serial time;
- four-argument cancellation: time, command name, missing latest time, FALSE;
- zero-argument `xlfNow` for Excel's own serial clock.

Every call keeps its counted command string and all XLOPER12 roots live through
Excel12v. It records both the raw C API return code and the immediate Boolean,
Excel-error, or unexpected result tag. Immediate scalar/error results create no
`ExcelOwnedValue` and no `xlFree` obligation.

Only an explicit `xlcontime-research` build of the example XLL registers
`RUST.ONTIME.SCHEDULE`, `RUST.ONTIME.CALLBACK`, `RUST.ONTIME.CANCEL`, and the
bounded diagnostic helpers. The default minimal XLL retains its pre-spike
registration and export surface. The typed descriptors, context methods,
bootstrap, status-file I/O, and research exports are all feature-gated and are
not approved production catalogue entries.

The callback receives `MacroContext`, checks the active runtime generation,
records process/thread/order/time information, and performs a harmless
preserving `xlAbort` poll. It is not connected to an M17 queue. The research
lifecycle bridge is `unsafe`: its caller must actually be executing inside a
genuine Excel-issued lifecycle callback on that callback thread. Backend linkage
alone cannot manufacture this proof.

## Experimental lifecycle rule

Pending entries retain the exact scheduled serial, command, form, and runtime
generation. Test-mode bootstrap is enabled only by a coordination marker whose
PID matches the current Excel process. With no marker, ordinary XLL open is
unchanged and no experiment runs. With a marker, the diagnostic artifact
records bootstrap attempted/succeeded/failure and unload-hazard state. Because
ordinary runtime initialization has already succeeded, `xlAutoOpen` returns
success even if research bootstrap fails; the harness must read the artifact
and must not infer research success from `RegisterXLL` or `xlAutoOpen`.

Close marks the generation inactive and attempts cancellation while the backend
remains linked. Failed cancellation is detected and ordinary runtime cleanup is
declined. Returning 0 from `xlAutoClose` is **not** known to prevent Excel from
unloading the DLL, so a pending callback remains an unload hazard. The isolated
harness cancels before unload; if that cannot be proved, it records an
unsafe/inconclusive result and terminates only its exact owned Excel process.

This lifecycle behavior is an implementation to validate, not evidence that
Excel honors cancellation or that an `xlAutoClose` return value controls DLL
unloading. The harness must never deliberately rely on that behavior.

## Production acceptance gate

`xlcOnTime` may become the default only after the issue #30 matrix establishes:

- normal current security settings allow the registered-command callback;
- callback context and main-thread legality through more than thread identity;
- two/latest/cancel result contracts are reproducible;
- repeated scheduling remains bounded and coalescible;
- close/unload cancellation is reliable across reload and Excel shutdown;
- XLM macro policy is not a required weakening;
- editing, modal UI, calculation, copy/paste, undo, latency, and idle CPU are
  acceptable.

Until then, no production dispatcher may use this mechanism autonomously.
