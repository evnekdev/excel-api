# M15.1 Excel COM preflight record

Date: 2026-07-15 (Australia/Brisbane)

Status: **implementation complete; live validation blocked**.

Command:

```powershell
powershell -File scripts/excel-stress-harness.ps1 -Preflight
```

Sanitized result:

- Excel: 64-bit Excel 16.0, build 20131.
- Worker and Excel were resolved as distinct exact PIDs; the coordination
  record included HWND and Excel start time.
- Excel process identity validation passed and the owned process exited after
  cleanup; no broad Excel process selection was used.
- TEMP and TMP existed and were writable; Excel initially had zero workbooks.
- Plain COM `Workbooks.Add()` before XLL registration failed with `0x800A03EC`
  (“cannot open or save any more documents”).
- `RegisterXLL` succeeded.
- `Workbooks.Add()` after registration failed with the same `0x800A03EC`.
- Classification: `plain-com-failed`.
- No Excel crash event was found in the cycle window.

This establishes that the current live-smoke blocker occurs in plain Excel COM
workbook creation before the XLL is loaded. It does not validate the smoke
suite, leak freedom, soak behavior, or the channel matrix. User names, local
paths, workbook data, and generated artifacts are intentionally omitted.
