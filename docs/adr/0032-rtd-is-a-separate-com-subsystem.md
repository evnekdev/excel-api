# ADR-0032: RTD is a separate COM subsystem

## Status

Accepted provisionally for compatibility prototyping. The M18.1 prototype is
implemented; production and live compatibility validation remain blocked.

## Context

M18 must identify a supported streaming architecture without assuming that an
RTD callback is an Excel C API callback. Microsoft exposes RTD through the
`RTD` worksheet function and the COM Automation `IRtdServer`/
`IRTDUpdateEvent` interfaces. The documentation defines method purpose and
Variant-shaped data exchange, but does not grant Excel12/Excel12v legality or
promise a physical callback thread.

The available Excel host still fails plain `Workbooks.Add` before an add-in or
RTD server is loaded. It cannot establish formula connection, refresh,
disconnect, termination, policy, or unload behavior.

## Decision

Use a separate in-process COM RTD server DLL as the target for the first future
compatibility prototype. Keep it in a separate, Windows-only
`excel-api-rtd` package, with a raw `excel-api-rtd-sys` package only if exact
generated bindings require one. Do not add COM dependencies to the default
core crates.

Do not combine the RTD class server and XLL by default. Retain an out-of-process
server as a future isolation and cross-bitness option. Continue to support
async UDFs for bounded one-shot work when streaming is not required.

Treat `UpdateNotify` and `RefreshData` only as RTD data-delivery operations.
They do not create any project Excel C API context and cannot drain the M17
dispatcher. Issue #30 therefore remains open.

Implement the minimal prototype as the unpublished Windows-only
`examples/minimal-rtd-server` package. Its source of truth is the installed
Excel type library version 1.9 reflected through `ITypeLib`/`ITypeInfo`, with
compile-time width/vtable checks. It is limited to the bounded `COUNTER` topic,
a finite/coalescing producer, GIT-marshaled notification, refresh, disconnect,
heartbeat, and termination. It makes no Excel C API calls.

The prototype uses stable non-production identifiers
`ExcelApi.MinimalRtd` and
`{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}`. Registration is explicit,
per-user, idempotent, and reversible. `Apartment` is a test hypothesis, not a
production threading decision. The available host still prevents formula and
lifecycle validation by failing plain `Workbooks.Add`.

## Consequences

- M18 is an architecture/compatibility milestone, not an RTD framework.
- An ordinary XLL cannot advertise RTD support without implementing and
  registering a COM server.
- XLL and RTD ownership, callbacks, signing, registration, and unload remain
  separate.
- No worker retains callback borrows or calls Excel12/Excel12v.
- Production selection remains conditional on a real Excel lifecycle and
  security matrix.
- The cooperative M17 dispatcher gains no autonomous progress mechanism.
- Outcome for M17 is **A: no evidence of general C API capability**. Issue #30
  remains open.

## Revisit conditions

Revisit this ADR after the validation plan passes on a supported 64-bit Excel
host. A production ADR must choose the supported Variant subset, public API,
installer/registration model, signing policy, apartment model, diagnostics,
and stress gates using recorded evidence rather than thread identity alone.

## Sources

The authoritative links are collected in
`RTD_STREAMING_ARCHITECTURE.md`. Key sources are Microsoft's `IRtdServer`,
`IRTDUpdateEvent.UpdateNotify`, RTD formula implementation notes, COM apartment
and marshaling guidance, server registration guidance, and Automation Variant
cleanup contract.
