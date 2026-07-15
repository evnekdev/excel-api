# ADR-0028: Isolated real-Excel stress harness

## Decision

Use a PowerShell parent/worker harness with one directly tracked PowerShell
worker per lifecycle cycle. After creating `Excel.Application`, the worker maps
`Application.Hwnd` to a PID with `GetWindowThreadProcessId`, validates that PID
as `EXCEL.EXE`, and immediately writes PID, HWND, start time, worker PID, and
Excel version/build to a coordination file.

On timeout the parent terminates its exact worker. It terminates Excel only if
the coordination file exists and the recorded PID still names `EXCEL.EXE` with
the same start time. A missing or stale record never authorizes a global Excel
process scan or kill. The direct process model avoids a nested `Start-Job`
host; its untouched `ProcessStartInfo.EnvironmentVariables` inherits the shell
environment without enumerating conflicting `PATH`/`Path` keys.

## Consequences

This validates registration, unload/reload, MTR calculation, and observable
workbook outputs without requiring a desktop Excel installation on ordinary
GitHub-hosted CI. The harness records version/build, timing, process memory and
handle/thread samples and descriptive trends, Excel-only Windows crash-event
evidence where readable, worker logs, and workbooks. Trends are evidence, not
proof that leaks are absent. A production XLL must not expose a test-only panic function, so real
panic injection is not claimed; Rust thunk tests continue to establish panic
containment and the harness tests a real controlled error fallback.

Preflight creates a plain COM workbook before registering the XLL, then repeats
after registration. On the current machine both fail with Excel `0x800A03EC`,
which localizes the blocker before XLL loading. M15 therefore remains
implementation complete with live validation blocked.
