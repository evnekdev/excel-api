# ADR-0032: RTD is a separate COM subsystem

## Status

Proposed for compatibility prototyping; production and live validation are
blocked.

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

Defer the minimal prototype until a working host and a reviewed raw contract
can verify the Office type-library ABI, Automation ownership, apartment/thread
behavior, reversible registration, policy, and termination. The prototype,
when authorized, is limited to one or two topics, a bounded/coalesced owned
producer, notification, refresh, disconnect, and termination. It makes no
Excel C API calls.

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
