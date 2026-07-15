# excel-api

`excel-api` is a Rust-first framework for native 64-bit Microsoft Excel XLL
add-ins. It provides the raw Excel 12 ABI, safe callback-borrowed and owned
values, return allocation, registration, procedural macros, typed callback
contexts and calls, lifecycle handling, diagnostics, async UDFs, and a
cooperative dispatcher.

## Status

The core 1.0 release target is the native Excel 12/XLL surface. It is not a
claim that every feature has completed real-Excel validation; see the support
matrix and release checklist for the remaining gates.

| Area | Status |
| --- | --- |
| ABI, values, returns, registration, macros, functions, commands, lifecycle, packaging | Stable target for 1.0 |
| Async UDF full lifecycle, cooperative dispatcher pump, extended stress/soak matrix | Implemented; live validation pending |
| RTD COM prototype, `xlcOnTime` compatibility probe | Experimental |
| Ribbon, general COM, task panes, autonomous notification | Deferred |
| 32-bit Excel, arbitrary background-thread Excel C API calls | Unsupported |

RTD, COM/Ribbon UI, custom task panes, `xlcOnTime`, and autonomous wake are
not part of the initial stable release and do not block core review.

## Quick start

```toml
[dependencies]
excel-api = "0.1"
```

```rust,no_run
use excel_api::prelude::*;

#[excel_function(
    name = "RUST.ADD",
    thunk = "rust_add",
    category = "Rust",
    description = "Adds two numbers.",
    arguments(left = "First addend.", right = "Second addend.")
)]
fn add(left: f64, right: f64) -> f64 {
    left + right
}
```

Use [the minimal XLL example](examples/minimal-xll/) for lifecycle exports and
manual registration. The snippet is `no_run` because an XLL needs a real Excel
host to load and register.

## Choose a crate

| Crate | Use it when… |
| --- | --- |
| [`excel-api`](crates/excel-api/) | Writing an XLL with safe values, registration, contexts, and macros. |
| [`excel-api-macros`](crates/excel-api-macros/) | Inspecting or directly depending on the procedural attributes (normally re-exported). |
| [`excel-api-sys`](crates/excel-api-sys/) | Building a narrowly audited raw Excel 12 FFI boundary. It supplies no ownership safety. |

## Safety model

Excel callback inputs are borrowed and may not escape their callback. Convert
them to owned semantic values before retaining or sending data to a worker.
Dynamic results are planned before being materialized into DLL-owned storage;
`xlAutoFree12` releases that storage. Excel C API results use their own `xlFree`
RAII path. Typed contexts encode callback capability—not merely a thread ID—and
arbitrary Excel C API calls from worker threads are forbidden. The preview
async-UDF subsystem has only the narrowly documented `xlAsyncReturn`
completion exception; it does not create a general Excel callback context.
Dispatcher operations still require a genuine compatible Excel-issued callback.

## Build and package

The supported target is 64-bit Windows Excel. Build and inspect the sample XLL:

```powershell
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/inspect-minimal-xll-exports.ps1
```

The normal XLL build neither contains nor registers the experimental RTD COM
prototype. Signing, publisher trust, and Office policy are deployment concerns;
do not weaken organization-wide security to test an add-in.

## Documentation

- [User guide](docs/guide/README.md)
- [Macro reference](docs/guide/macro-reference.md)
- [Support matrix](SUPPORT_MATRIX.md)
- [Release readiness audit](docs/release/core-1.0-readiness-audit.md)
- [Release checklist](docs/release/core-1.0-release-checklist.md)
- [Architecture index](ARCHITECTURE_INDEX.md) (maintainer reference)

## License

Licensed under either of [Apache License, Version 2.0](LICENSE-APACHE) or
[MIT license](LICENSE-MIT) at your option.
