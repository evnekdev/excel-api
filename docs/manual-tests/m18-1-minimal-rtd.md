# M18.1 minimal RTD compatibility evidence

Date: 2026-07-15

Status: **M18 prototype implemented; Excel activation unresolved**.

## Automated and deployment evidence

- Target: `x86_64-pc-windows-msvc`, 64-bit DLL, Rust stable workspace toolchain.
- Excel: Microsoft 365 64-bit, version 16.0, build 20131; file version
  16.0.20131.20154.
- Prototype: `ExcelApi.MinimalRtd`, CLSID
  `{DC738FE5-30EE-40E8-A8C2-3D16F217C52D}`.
- The installed Excel 1.9 type-library audit passed.
- The release DLL built and exported exactly `DllGetClassObject` and
  `DllCanUnloadNow`; it contained no XLL or Excel12/Excel12v export.
- Reversible HKCU registration, inspection, and unregister passed with the
  absolute build path omitted from this record. Registration included
  `InprocServer32`, `ThreadingModel=Apartment`, ProgID/CLSID mappings, and the
  standard `Programmable` marker.
- Direct COM activation succeeded. An Automation `Heartbeat` dispatch returned
  zero in `Created`, and diagnostics recorded the call without panic.

## Excel validation attempts

Commands:

```powershell
powershell -File scripts/build-minimal-rtd.ps1 -Profile release
powershell -File scripts/inspect-minimal-rtd-exports.ps1
powershell -File scripts/excel-rtd-validation.ps1 -ProcessTimeoutSeconds 30
powershell -File scripts/excel-rtd-validation.ps1 -ProcessTimeoutSeconds 45
```

One isolated run successfully created a plain workbook and completed the
per-user registration round-trip. Both duplicate `COUNTER` formulas remained
at Excel error 2042 (`#N/A`), no `ServerStart` diagnostic was emitted, and no
RTD update was claimed. Workbook close, Excel quit, exact process exit, and
unregister all succeeded.

A repeat run reproduced the host's intermittent plain `Workbooks.Add` hang.
The parent timed out after 45 seconds, read the worker-written Excel PID/HWND/
start time, verified process name and start-time identity, terminated only that
owned Excel process, and rolled back registration. Excel spawned a direct child
`EXCEL.EXE` during cleanup; parent-PID evidence tied it to the owned instance,
but the current user could not terminate that child (`Access denied`). The
harness now records and attempts only such verified direct children. This
remaining process requires host-administrator cleanup or restart before another
live run. No unrelated Excel process was selected.

The successful direct COM activation narrows the unresolved failure to Excel
RTD activation/policy/registration compatibility rather than the base class
factory. It does not prove formula behavior, topic callbacks, `UpdateNotify`,
`RefreshData`, disconnect, termination, or unload under Excel.

M18.2 repeated direct activation with boundary JSONL evidence: the Rust DLL
reached `DllGetClassObject`, `CreateInstance`, IDispatch type-info/name/invoke,
`Heartbeat`, and `ServerTerminate`. The Microsoft-PIA control activated and
returned heartbeat 1. The clean Excel comparison was refused because PID 19616
is a verified direct descendant of prior owned Excel PID 7060, while its image
and start details are inaccessible. The harness did not terminate it and
requires administrator-assisted cleanup or reboot before retrying.

## Remaining working-host matrix

Repeat on a supported host where plain workbook creation is reliable and
record: formula connection, initial and repeated updates, duplicate topics,
disconnect/reconnect, workbook close, `ServerTerminate`, Excel close,
`DllCanUnloadNow`/module unload evidence, registration rollback, callback
threads/apartments, calculation mode, policy/signing, and bounded memory/handle
trends. Do not weaken system-wide Office policy.

No Excel C API call or callback context exists in the prototype. The M17 result
is outcome A: no evidence of general C API capability; issue #30 remains open.
