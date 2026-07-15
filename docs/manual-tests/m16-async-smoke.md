# M16 asynchronous UDF live validation

Date: 2026-07-15

## Environment

- Windows 64-bit, Excel 16.0 build 20131, 64-bit process
- Release XLL built for Windows x64
- Excel channel was not reported by the available COM diagnostics

No user name, machine-private identifier, generated workbook, or local
absolute path is retained in this record.

## Commands

```powershell
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/inspect-minimal-xll-exports.ps1
powershell -File scripts/excel-stress-harness.ps1 -ValidateOnly
powershell -File scripts/excel-stress-harness.ps1 -Mode Smoke
```

## Results

- Release build: passed.
- Export inspection: passed; the async UDF and both calculation-event
  procedures were present.
- Pure harness validation: passed (15 assertions).
- Real Excel registration: succeeded.
- Real workbook/smoke stage: blocked before function setup. Plain COM
  `Workbooks.Add()` failed with Excel `0x800A03EC`; the post-registration
  workbook attempt failed identically. Classification: `plain-com-failed`.
- The owned Excel process exited and no timeout cleanup was required.

Because the failure occurs in plain Excel before any workbook or async formula
exists, this run provides no live M16 completion, cancellation, recalculation,
or unload evidence. Automated ABI, ownership, cancellation, shutdown, and race
tests pass, but M16 remains implementation-complete with real Excel validation
pending. M15 likewise remains implementation-complete with live validation
blocked; this record does not upgrade either milestone to stable.
