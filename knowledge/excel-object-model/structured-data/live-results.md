# Structured-data live result

Environment ID: `post-prompt13-visible-fresh-process`.

The controlled test began with zero `EXCEL.EXE` processes, created one visible
Excel process, closed its workbook without saving, called `Application.Quit`,
and finished with zero `EXCEL.EXE` processes.

The passing test is `excel-com/tests/structured_data_live.rs`. It records the
three environment-specific observations retained in the research note:

- InsertRowRange was absent on a populated table without the UI insert row.
- A table AutoFilter did not imply worksheet AutoFilterMode.
- ListObject Sort.SetRange was rejected because the sort object was table-bound.

After that pass, retrying the new and historical live tests from a zero-process
state failed at unchanged `Workbooks.Add` with Excel exception `0x800A03EC`.
The non-Rust PowerShell Automation binder also failed to open the repository's
read-only `IntlMap.xlsx` fixture in that state. No process was terminated; each
transient Excel process exited after a passive wait.

The user then rebooted. The `2026-07-22T20:16:42+10:00` cold boot started with
zero `EXCEL.EXE` processes. At `2026-07-22T20:21:04+10:00`, the independent
Microsoft-style raw control (`excel-com-microsoft-sample rust-child`, baseline
profile) again reached `Workbooks.Add` and received `0x800A03EC`; it invoked
`Quit` normally and left zero Excel processes. The isolated minimal high-level
reproduction, the full high-level control, and the generic `windows-sys`
`IDispatch` control then reproduced the same `DISP_E_EXCEPTION` / `0x800A03EC`
result. Each control began with zero processes, invoked `Quit`, and its
transient Excel process exited during a bounded passive wait. The native C++
and C-ABI-shim controls were not run because CMake configuration stalled before
compilation or any Excel launch; the already completed independent controls
were sufficient to stop the live suite. This establishes that the remaining
live matrix is blocked by the current Excel host even after a clean reboot,
rather than by a stale process or only the high-level Rust wrapper.

After the user manually opened and closed Excel, the structured-data live test
was retried from zero `EXCEL.EXE` processes. It again failed at the unchanged
`Workbooks.Add` invocation with `0x800A03EC`, then naturally left zero Excel
processes. Opening Excel interactively is therefore not an observed recovery
for this automation-host condition.
