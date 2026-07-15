# Callback and Lifecycle Architecture

## Exported callbacks

- `xlAutoOpen`;
- `xlAutoClose`;
- `xlAutoAdd`;
- `xlAutoRemove`;
- `xlAddInManagerInfo12`;
- optional `xlAutoRegister12`;
- `xlAutoFree12`.
- parameterless `excel_api_calculation_canceled` and
  `excel_api_calculation_ended` procedures registered through
  `xlEventRegister` when async functions are present.

## Initialization state machine

```text
Uninitialized
 -> Initializing
 -> Initialized
 -> Closing
 -> Closed/Uninitialized
 -> CleanupRequired (only when Excel cleanup fails)
```

Initialization must be idempotent.

The book documents callback sequences where:

- Add-in Manager info may be requested before normal open initialization;
- add/open can occur without a preceding remove/close;
- close can be called during a shutdown attempt that is later cancelled.

Therefore:

- `xlAutoOpen` checks whether already initialized;
- `xlAddInManagerInfo12` and `xlAutoAdd` may ensure initialization;
- duplicate menus/registrations must be avoided;
- destructive shutdown must be conservative and version-aware.

## `xlAutoClose`

Responsibilities:

- unregister functions where reliable;
- stop/coordinate workers;
- release runtime-owned resources;
- unlink Excel call interface last.

Do not unload call pointers before destructors/releases that need them.

M16 disables and drains the async controller before registration cleanup.
`CalculationCanceled` cooperatively cancels scheduled requests;
`CalculationEnded` removes terminal calculation-scoped records. These event
procedures are panic-contained and make no Excel callbacks.

## `xlAutoFree12`

- the reusable implementation lives in `excel-api` as
  `unsafe extern "system" fn xl_auto_free12(*mut XLOPER12)`;
- the actual exact-name `xlAutoFree12` export lives in `minimal-xll` as a thin
  delegate, avoiding an unconditional global export in the reusable `rlib`;
- null is a defensive no-op; a valid non-null pointer must be a unique handoff
  from the same loaded XLL and must be supplied exactly once;
- the callback is allocation-free on its normal path, panic-contained, and
  makes no Excel calls or project-lock acquisitions;
- it reconstructs and drops exactly one offset-zero `ReturnAllocation`, for
  scalar roots as well as pointer-bearing roots;
- it cannot safely recover from arbitrary invalid pointers or duplicate
  callbacks after memory has already been freed;
- `catch_unwind` contains unwinding panics but cannot catch `panic = "abort"`.

The next lifecycle milestone may generate or relocate the thin export, but it
must retain the verified `void WINAPI (LPXLOPER12)` ABI and delegate to the same
core ownership implementation.

## `DllMain`

No Excel C API calls. Avoid nontrivial initialization under loader lock.

## Experimental xlcOnTime close ordering

Only the explicitly feature-built M17 research XLL tracks exact scheduled
serials and runtime generations. Its close path marks the experiment inactive,
attempts cancellation while the backend is still linked, and declines ordinary
runtime cleanup if any pending call cannot be canceled. Only a successful
experimental cancellation pass proceeds to command unregistration and backend
unlink. A stale-generation callback is a no-op.

This ordering is the invariant under test; it is not proof that current Excel
honors cancellation. There is no authoritative basis for claiming that return
0 from `xlAutoClose` prevents DLL unload. Failed cancellation can therefore
leave an unload hazard, and the isolated harness must terminate its exact owned
Excel process instead of deliberately unloading. Production autonomous
scheduling remains prohibited until the close/unload/reload matrix passes on a
working host.
## M8 implementation

`Runtime` implements `Uninitialized`, `Initializing`, `Initialized`,
`Closing`, and `CleanupRequired`. Duplicate open/add calls are idempotent, partial registration rolls
back in reverse order, close is idempotent, and failed close retains failed IDs
with the backend linked for retry. No lock is held over Excel12v. Thin,
panic-contained exports implement `xlAutoOpen`, `xlAutoClose`, `xlAutoAdd`,
`xlAutoRemove`, `xlAddInManagerInfo12`, and `xlAutoFree12`.

M9B changes only worksheet-function exports. Lifecycle exports remain thin and
handwritten. `Runtime::production` and generated callback scopes share one SDK
callback backend per loaded binary so contexts observe the entry installed by
`xlAutoOpen`/`SetExcel12EntryPt`; lifecycle state and registration IDs remain
owned by the XLL's `Runtime`.

## M12 lifecycle decision

`xlAutoAdd` ensures initialization and `xlAutoRemove` performs the same
idempotent close path as `xlAutoClose`; both exports contain panics. Add-in
Manager information remains a pure metadata query. `xlAutoRegister12` is
intentionally not exported: the explicit descriptor registry always provides
complete type text, and Microsoft warns that registering without it recurses.
The optional callback can be reconsidered only if a documented on-demand
registration use case requires it.

## M16 asynchronous lifecycle

Initialization registers both Microsoft-defined calculation events before
publishing async worksheet registrations. A generated async entry thunk
returns void after copying its handle and all supported inputs. Worker
completion owns its return allocation locally during the synchronous
`xlAsyncReturn` call; it never transfers that allocation to `xlAutoFree12`.
Atomic request state and a controller epoch enforce at-most-once completion
and reject cancellation, duplicate completion, and post-close work.

Event-registration lifetime is longer than one open/close generation.
Microsoft's verified documentation defines `xlEventRegister` and the two event
meanings but no removal operation or repeated-registration semantics. The
production `Runtime`, which is process-static for one loaded XLL binary,
records each successful event registration and never duplicates it on reopen.
If one registration fails, only the missing event is retried. With no active
controller, both exported procedures are safe no-ops.

Each successful open consumes a fresh executor/controller generation. Close
removes and disables it, permanently shuts its executor, and joins accepted
jobs before unregister or backend unlink. If Excel cleanup fails, the runtime
reports `CleanupRequired`; async scheduling stays disabled, initialize is
rejected, and retry-close uses the retained registration IDs. A successful
retry returns to `Uninitialized`, after which the XLL must install a fresh
executor for reopen.
