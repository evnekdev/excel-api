# RC2 live Excel validation

## Status

**Blocked: host plain-workbook preflight failed.** RC2 does not treat this as
an XLL or framework failure because the same isolated Excel instance could not
create a workbook before the XLL was registered. Per the RC2 stop condition,
no worksheet, lifecycle, async, dispatcher, stress, or soak scenarios were
attempted on this host.

## Test artifact

- Tested commit: `077d4ddcd33904dd027ffba06c621936395b19e8`
- Date: 2026-07-15 (Australia/Brisbane)
- Windows: Windows 10 Enterprise, version 25H2, build 26200.8875, x64
- Excel: Microsoft 365 x64, Excel `16.0`, build `20131`
- Office channel: Click-to-Run channel `492350f6-3a01-4f97-b9c0-c7c6ddf67d60`
- Rust: `1.97.0`, target `x86_64-pc-windows-msvc`
- Logical processors: 28
- XLL: unsigned release `minimal_xll.xll`
- XLL SHA-256:
  `CDA9900709211C3B6EAA505713C3D7DB91494E1F84711D2D58EDB94A9953910A`

The generated, uncommitted evidence is retained under
`artifacts/rc2-live-validation/`. It contains only synthetic harness data and
is intentionally excluded from version control.

## Preflight

Command:

```powershell
powershell -File scripts/excel-stress-harness.ps1 -Preflight \
    -OutputDirectory artifacts/rc2-live-validation
```

The harness created an isolated Excel process, recorded its PID, HWND and
start time, and verified that the process exited after cleanup. Its plain COM
workbook creation check failed with `0x800A03EC`:

> Microsoft Excel cannot open or save any more documents because there is not
> enough available memory or disk space.

The post-registration check returned the same error. `RegisterXLL` returned
success, but it is not live-function evidence because the prerequisite plain
workbook check had already failed. The harness classified the result as
`plain-com-failed`.

The preflight worker and owned Excel process both exited; no timeout cleanup
was required and no stale Excel process remained. The collected process samples
are diagnostic evidence only and do not establish a memory trend.

## Matrix

| Area | Status | Result |
|---|---|---|
| Automated release baseline | Passed | Format, strict Clippy, all-feature tests, doctests, no-default core test, MSRV check, ABI check, release XLL build, export inspection, and warning-denied core Rustdoc passed. |
| Plain Excel workbook preflight | Blocked | `Workbooks.Add()` failed before XLL registration with `0x800A03EC`. |
| XLL load and registration | Not run | Registration was observed by the preflight only; no workbook was available to validate loading behavior. |
| Synchronous core, strings, arrays, Q/U, optional values | Not run | Requires a healthy workbook host. |
| Commands, contexts, ownership cleanup, lifecycle and MTR | Not run | Requires a healthy workbook host. |
| Async UDF preview | Not run | Requires a healthy workbook host. |
| Cooperative dispatcher preview | Not run | Requires a healthy workbook host. |
| Stress and soak | Not run | The preflight is a mandatory gate. |
| Signing and trust | Not run | No approved signing material was used. |

## Automated baseline

The following commands passed before the live preflight:

```powershell
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo test -p excel-api --no-default-features
cargo +1.85.0 check --workspace --all-targets --all-features
cargo run --manifest-path tools/abi-check/Cargo.toml
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/inspect-minimal-xll-exports.ps1
$env:RUSTDOCFLAGS = '-D warnings'
cargo doc -p excel-api-sys -p excel-api-macros -p excel-api --features macros --no-deps
powershell -File scripts/excel-stress-harness.ps1 -ValidateOnly
```

The ABI checker reported 145 passing checks. Export inspection reported exactly
18 required production exports.

## Required next action

Repair or replace the Excel automation host, then rerun the plain/post-XLL
preflight from the same commit or its successor. Only a preflight classification
of `passed` permits RC2's workbook matrix to begin. This evidence does not
change the support matrix: async UDFs and the cooperative dispatcher remain
preview with live validation pending, and no optional RTD or COM integration is
a core RC2 blocker.
