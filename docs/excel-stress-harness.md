# M15 real-Excel stress harness

Status: **implementation complete; live validation blocked**. The current
machine cannot create a plain Excel COM workbook before the XLL is loaded; see
[the sanitized M15.1 preflight record](manual-tests/m15-1-excel-preflight.md).

Build the release XLL and run the deterministic 64-bit Excel smoke command:

```powershell
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/excel-stress-harness.ps1 -Mode Smoke
```

Smoke uses two lifecycle cycles, 25 full recalculations per cycle, and a
180-second per-cycle timeout. Soak uses 25 cycles, 2,000 recalculations per
cycle (50,000 rebuilds), and a 1,800-second timeout. An explicit
`-ProcessTimeoutSeconds` overrides either default. The effective timeout is
recorded in `summary.json`.

## Ownership and timeout cleanup

The parent starts one worker PowerShell process directly, with redirected logs
and the current environment inherited without enumerating `PATH`/`Path`. The
worker creates `Excel.Application`, reads `Application.Hwnd`, resolves it with
the audited `GetWindowThreadProcessId` P/Invoke, validates `EXCEL.EXE`, and
immediately writes this per-cycle identity:

- worker PID;
- Excel PID and window handle;
- Excel start time;
- Excel version/build.

On timeout the parent kills its directly tracked worker. It kills Excel only
when the coordination file exists and the PID still belongs to `EXCEL.EXE`
with the recorded start time. If the record is missing, invalid, stale, or the
PID was reused, no Excel process is killed. The harness never selects processes
by comparing global before/after Excel process lists.

## Preflight and environment evidence

Run the isolated diagnostic preflight with:

```powershell
powershell -File scripts/excel-stress-harness.ps1 -Preflight
```

It attempts workbook creation first in plain Excel COM and then after
`RegisterXLL`. It records Excel version/build/bitness, process path/session,
user and integrity SID, safe/automation switches when readable, default and
startup paths, TEMP/TMP existence and writeability, disk space, existing
workbook count, registration outcome, and both workbook outcomes. It changes no
Excel security setting.

## Artifacts and trend evidence

Each timestamped directory below `target/excel-stress/` contains ownership,
worker logs, cycle JSON, workbooks when creation succeeds, and `summary.json`.
The exact owned Excel process is sampled after start, registration and workbook
setup, periodically during calculation, before unregister/quit, and after exit.
Samples include working set, private bytes, handles, threads, stage, and time.
Per-cycle first/last/minimum/maximum/delta summaries are trend evidence only;
they do not prove that leaks are absent.

Crash evidence is limited to Application Error or Windows Error Reporting
events whose message names `EXCEL.EXE`, within the cycle window. PID correlation
is recorded where the event exposes it. Denied event-log access is recorded as
explicitly unavailable; unrelated application failures are excluded.

The probes cover every sample function and `RUST.PING.COMMAND`, scalar values,
direct UTF-16 strings, Q arrays, U references, missing/blank input, Excel error
values, a controlled conversion fallback, MTR, registration, and unload/reload.
The sample XLL has no public panic endpoint; Rust thunk tests remain the
authoritative panic-containment evidence.

`-ValidateOnly` runs pure helper assertions without Excel. They cover exact PID
selection and start-time validation, unrelated-process exclusion, missing
coordination, cleanup selection, mode defaults, event filtering, process trend
aggregation, required JSON fields, and preflight classification.

## Runner matrix

| Runner | Required evidence |
|---|---|
| Windows x64, Excel Current Channel | smoke summary with version/build, MTR and process evidence |
| Windows x64, Excel Monthly Enterprise Channel | smoke summary with version/build, MTR and process evidence |
| Windows x64, Excel Semi-Annual Enterprise Channel (when supported) | smoke summary with version/build, MTR and process evidence |

No channel is marked complete yet. Do not weaken macro, Protected View, or
trusted-location policy. After one smoke pass, status may become “M15
smoke-validated; soak/channel matrix pending.”
