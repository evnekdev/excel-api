# Packaging Architecture

Status: reproducible M14 XLL packaging implemented; M20 crate-publication
rehearsal and release evidence remain open.

Responsibilities:

- build `cdylib`;
- produce `.xll`;
- verify exports;
- generate/link `.def` when required;
- include version resources;
- package x64 artifacts;
- optional code signing;
- emit diagnostics and reproducible metadata.

`scripts/build-minimal-xll.ps1` builds the Windows x64 MSVC `cdylib` and copies
`target/<profile>/minimal_xll.dll` to `minimal_xll.xll`. Rust's unmangled x64
exports avoid x86 decoration and a `.def` file is not required for this slice.
Use `dumpbin /exports` to verify lifecycle and worksheet symbols.

## Optional RTD packaging boundary

The M18.1 RTD prototype is a separate 64-bit in-process COM DLL with
its own ProgID/CLSID, class factory exports, signing identity, registration,
and rollback. It is not copied to `.xll`, does not share XLL exports, and does
not add COM dependencies to the default workspace packages.

An in-process server must match Excel's process bitness. Per-user registration
is the first non-elevated compatibility path; per-machine registration belongs
to an explicit installer. Registration-free COM is not assumed because the
activation manifest belongs to the Excel host. An out-of-process server remains
a future crash-isolation/cross-bitness alternative.

`scripts/build-minimal-rtd.ps1` builds the unpublished
`excel-api-minimal-rtd` package as
`target/release/excel_api_minimal_rtd.dll`. Export inspection requires exactly
`DllGetClassObject` and `DllCanUnloadNow` and rejects XLL lifecycle/backend
symbols. The minimal XLL remains a different artifact with its unchanged 18
production exports. Registration is never a build side effect: the dedicated
register/inspect/unregister scripts modify only the two prototype roots under
`HKCU\Software\Classes` and record `ThreadingModel=Apartment` as a hypothesis
for live compatibility testing.

The RTD prototype, test control server, registration scripts, and research
helpers are excluded from core package publication and ordinary XLL packaging.
They are never built, registered, or installed as a side effect of normal core
build/test/package workflows.

For core 1.0, Cargo default members are `excel-api-sys`, `excel-api`,
`excel-api-macros`, and `minimal-xll`. The RTD prototype remains an explicit
workspace member for compatibility testing but is `publish = false`; the C#
control server and research scripts are not Cargo packages. The three core
package manifests have no COM dependency.

Each core crate package carries crates.io metadata, a package-specific README,
and both Apache-2.0 and MIT license texts. Package contents and staged dry-run
results are release gates in
[`docs/release/core-1.0-release-checklist.md`](docs/release/core-1.0-release-checklist.md).
