# excel-api

A Rust-first project for building native Microsoft Excel add-ins through the
Excel C API.

The intended long-term outcome is an idiomatic Rust ecosystem for native `.xll`
add-ins: raw C API bindings, safe values and calls, generated worksheet-function
registration, arrays, commands, asynchronous calculation, and optional Windows
integrations.

## Current status

Milestone 8's Windows x64 XLL implementation and automated live Excel smoke
test are complete. M9A generates typed registration metadata, M9B now
generates the five worksheet-function ABI thunks from that same closed type
model, and M10 pins compile-time macro conformance diagnostics with `trybuild`.
The runtime resolves the SDK-defined `MdCallBack12` entry point, uses
callback-scoped Excel-owned result RAII for lifecycle calls, and returns
dynamic values through fresh per-call DLLFree storage. The generated-thunk XLL
passed the same two-process 64-bit Excel load, calculation, MTR, unload, and
reload smoke; details are in [the smoke-test record](docs/manual-tests/m8-excel-smoke-test.md).
M15 adds the opt-in [isolated Excel stress harness](docs/excel-stress-harness.md)
for repeatable smoke and soak validation on self-hosted 64-bit Excel runners.
Its implementation is complete, but live validation is blocked on the current
machine by plain Excel COM workbook creation failing before the XLL is loaded.
M16 now implements bounded native asynchronous UDF scheduling, generated
`>...X` thunks, cancellation events, at-most-once `xlAsyncReturn`, and
permanent per-open executor generations with shutdown-safe draining. Queued
cancellation skips user code, event handlers are registered once per loaded
binary, and failed close cleanup is represented explicitly. Automated coverage passes; real Excel
cancellation/recalculation/unload validation remains pending, and does not
change the blocked M15 live-smoke status.
M17 implements a bounded cooperative dispatcher with owned tickets, sealed
operations, per-open generations, capability-typed drains, and the explicit
`RUST.DISPATCH.PUMP` command. Enqueueing does not wake Excel; autonomous
notification remains open. The isolated `xlcOnTime` research decision remains
inconclusive and is not used by the production dispatcher. Automated coverage
passes, while live pump validation remains blocked by the current host's plain
workbook-creation failure.
The initial stable release is the native Excel 12/XLL core: ABI, owned and
callback-borrowed values, return ownership, registration/macros, typed
contexts/calls, functions/commands/lifecycle, diagnostics, packaging, stress
infrastructure, async UDFs, and cooperative dispatch. See the
[support matrix](SUPPORT_MATRIX.md).

RTD/streaming, general COM/Ribbon, task panes, autonomous notification, and
the `xlcOnTime` probe are optional post-1.0 work. The M18 prototype remains in
the repository as an experimental, Windows-only, unpublished component; it is
not production-supported or part of normal XLL packaging. It has a verified
Office 1.9 ABI audit and direct COM activation, but Excel-formula activation is
unresolved. See [RTD tracking issue #37](https://github.com/evnekdev/excel-api/issues/37).

Build the loadable artifact with:

```powershell
pwsh -File scripts/build-minimal-xll.ps1 -Profile release
```

## Architecture references

- [Overall architecture](ARCHITECTURE.md)
- [Architecture index](ARCHITECTURE_INDEX.md)
- [Excel 12 ABI architecture](ABI_ARCHITECTURE.md)
- [Memory and ownership architecture](MEMORY_OWNERSHIP_ARCHITECTURE.md)
- [String architecture](STRING_ARCHITECTURE.md)
- [Array and reference architecture](ARRAY_REFERENCE_ARCHITECTURE.md)
- [Type conversion architecture](TYPE_CONVERSION_ARCHITECTURE.md)
- [Threading architecture](THREADING_ARCHITECTURE.md)
- [Runtime context architecture](RUNTIME_CONTEXT_ARCHITECTURE.md)
- [Excel call architecture](EXCEL_CALL_ARCHITECTURE.md)
- [Callback and lifecycle architecture](CALLBACK_LIFECYCLE_ARCHITECTURE.md)
- [Registration architecture](REGISTRATION_ARCHITECTURE.md)
- [Error architecture](ERROR_ARCHITECTURE.md)
- [Testing architecture](TESTING_ARCHITECTURE.md)
- [Asynchronous UDF architecture](ASYNC_ARCHITECTURE.md)
- [Main-thread dispatch architecture](MAIN_THREAD_DISPATCH_ARCHITECTURE.md)
- [RTD streaming architecture](RTD_STREAMING_ARCHITECTURE.md)
- [COM architecture boundary](COM_ARCHITECTURE.md)
- [Ribbon UI architecture](RIBBON_UI_ARCHITECTURE.md)
- [Core 1.0 support matrix](SUPPORT_MATRIX.md)
- [Optional integrations roadmap](OPTIONAL_INTEGRATIONS_ROADMAP.md)
- [Implementation roadmap](IMPLEMENTATION_ROADMAP.md)
- [Excel-DNA capability map](EXCELDNA_CAPABILITY_MAP.md)
- [Codex development prompts](prompts-dev/README.md)

These documents are living references and must track implementation changes.

## Workspace

```text
crates/
  excel-api-sys/
  excel-api/
  excel-api-macros/
examples/
  minimal-xll/
  minimal-rtd-server/  # experimental Windows-only, unpublished RTD prototype
docs/
  adr/
  checklists/
  diagrams/
  research/
prompts-dev/
tools/
  Excel2013XLLSDK/
  abi-check/
```

The Microsoft SDK is the authoritative ABI reference. See
[`tools/Excel2013XLLSDK/README.md`](tools/Excel2013XLLSDK/README.md).

## Prompt 01 readiness

Complete [`prompts-dev/01-verify-xloper12-abi.md`](prompts-dev/01-verify-xloper12-abi.md)
before safe wrappers.

```powershell
cargo run --manifest-path tools/abi-check/Cargo.toml
```
