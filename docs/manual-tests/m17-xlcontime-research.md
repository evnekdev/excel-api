# M17 xlcOnTime research run

Date: 2026-07-15

## Host

- Windows 64-bit NT 10.00
- Microsoft Excel 16.0, build 20131, 64-bit
- locally built unsigned XLL
- current-user Excel 4.0 macro registry values were not explicitly configured;
  organization policy and Trust Center effective state were not changed

No user name, machine name, PID, local absolute path, or workbook content is
retained here.

## Commands

```powershell
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo run --manifest-path tools/abi-check/Cargo.toml
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/inspect-minimal-xll-exports.ps1
powershell -File scripts/excel-ontime-validation.ps1 -ValidateOnly
powershell -File scripts/excel-ontime-validation.ps1
```

## Results

- Checked-in ABI: `xlfNow = 74`; `xlcOnTime = 32916 = 0x8094`.
- ABI checker: 145 checks passed.
- Release XLL: built successfully; all 21 required exports were present.
- Pure wrapper tests: 2-, 3-, and 4-argument tags/counts, exact command text,
  Boolean/error decoding, raw C API code preservation, and no `xlFree` passed.
- Plain COM `Workbooks.Add`: failed before XLL registration with the known host
  document-creation error.
- Post-XLL workbook test: unavailable because the same host error persisted.
- `RegisterXLL`: returned TRUE.
- Direct COM `Application.Run`: registered test command reported unavailable or
  macros disabled without a workbook.
- `ExecuteExcel4Macro`: returned an error value and did not invoke the command.
- The owned Excel process had not exited within the harness's 15-second
  observation window after the blocked run cleanup; it was no longer present
  on the later process check. This is not unload/cancellation evidence.
- Timed callback, main-thread context, cancellation, latest-time, close/unload,
  recalculation, and user-experience results: not obtained.

Classification: **inconclusive; live validation blocked by the Excel host**.

## Remaining matrix

A working supported host must still run default/XLM-enabled/XLM-disabled
security cases; trusted/ordinary and signed/unsigned cases where permitted;
two/latest/repeated schedules; every cancellation case; calculation/edit/modal
delay; close, unload/reload, and Excel shutdown; copy/paste/undo; latency,
flicker, and idle CPU. No production M17 approval follows from this run.
