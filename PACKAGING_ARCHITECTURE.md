# Packaging Architecture

Status: minimal M8 build implemented; full M14 packaging remains planned.

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
