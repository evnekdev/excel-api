# M18.2 RTD activation and cleanup diagnosis

Date: 2026-07-15 (Australia/Brisbane)

Status: **M18 prototype implemented; Excel activation unresolved**.

## Cleanup result

Every accepted producer owns an RAII lifetime guard. Finite completion, normal
stop, HRESULT failure, and panic each emit one bounded terminal event and
decrement `ACTIVE_PRODUCERS` once. `ServerTerminate` inspects the join result;
a panicked join returns controlled `E_UNEXPECTED` after cleanup continues.
Committed notification calls have equivalent panic-safe accounting.

Callback registration retains `Registered(cookie)`, `Revoking(cookie)`, or
`RevocationFailed(cookie)` until GIT revocation succeeds. Later termination
retries the same cookie. Successful revoke clears it and decrements
`CALLBACK_COOKIES` once. Persistent failure leaves unload blocked.

## Activation evidence

- Windows reported `Microsoft Windows NT 10.0.26200.0`.
- Excel is 64-bit Microsoft 365 version `16.0.20131.20154`.
- The Rust DLL is PE32+ x64. A 64-bit `LoadLibraryEx` probe succeeded.
- Dependencies resolve through Windows/OLE, `VCRUNTIME140.dll`, and Universal
  CRT imports; no debug CRT was reported.
- The DLL is unsigned.
- Read-only Office-key inspection found no `DisableRTD`; it observed
  `VBAWarnings=1`. Absent keys do not rule out endpoint controls.
- HKCU registration, merged 64-bit HKCR, exact server path, `Apartment`, and
  rollback passed in the direct-activation round trip.
- Direct Rust activation reached class factory, object creation, IDispatch
  type-info/name/invoke, heartbeat, and termination. Heartbeat returned 0 in
  the intentionally inactive Created state.
- The test-only `ExcelApi.ControlRtd` was built against the installed Microsoft
  Office Excel PIA, activated directly, and returned heartbeat 1.

The clean Excel comparison did not run. Its preflight found PID 19616, a
verified direct descendant of prior owned Excel PID 7060 in an M18.1 artifact.
Image, command line, and start time were inaccessible. The harness refused to
kill it or launch another Excel test. Cleanup requires administrator-assisted
inspection/termination or reboot.

## Formula and decision matrix

The next clean run tests omitted-server and explicit-empty-server syntax,
lower/upper-case ProgIDs, duplicates, and `FormulaLocal` with Excel's list
separator. Manual entry is a separate observation because COM entry cannot
prove identical interactive behavior.

Outcomes A-E are **unclassified**: neither server was exercised through Excel
on a clean host. A requires the control to fail before `ServerStart`; B-D need
a clean control/Rust comparison; E needs the full lifecycle. M19 remains
blocked.

## Recorded results

- Rust RTD tests: 24 passed, including producer, join, notification, and GIT
  retry injection.
- RTD release build and exact two-export inspection: passed.
- Office 1.9 type-library ABI baseline: passed.
- Rust registration/direct activation/unregistration: passed.
- PIA control build and direct activation: passed.
- Rust/control validation-only and formula helper checks: passed.
- Excel Rust/control comparison: blocked before Excel launch by stale owned-test
  descendant.

No Excel12/Excel12v call, callback context, M17 integration, public production
RTD API, or M19 work was added.
