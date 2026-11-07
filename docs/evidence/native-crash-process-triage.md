# #137: crash correlation and safe x64 resumption

Windows Application Error event 1000 at 2026-11-07 16:22:09 local identifies
faulting PID `0x715C` = 29020, the test-owned Excel process in REF-STALE. It records
EXCEL.EXE 16.0.20527.20150, access violation `0xc0000005`, offset `0xc38038`.
The corresponding WER report ID is 473c96e8-25fe-4a7c-a846-de0e925080bc. Thus the
earlier RPC loss is correlated with an actual host crash, not ordinary rejection.
It does not by itself establish whether the Rust wrapper, an expired Excel
identity or an installation problem caused the access violation.

The later OFFICE_MODULE_VERSION_MISMATCH bucket belongs to that same report.
ClickToRun VersionToReport and ClientVersionToReport, EXCEL.EXE's file version,
and the faulting-module version all report 16.0.20527.20150. The bucket does not
name a mismatched module. Access to the protected WER archive was denied by
Windows, so module-by-module or dump inspection has not been completed. No ACL,
Office configuration, installer or repair operation was attempted. There is not
enough evidence to prescribe an Office repair.

Remaining PID 19684 has parent PID 29020 and creation time 16:22:08. Repeated CIM
observations show zero threads, zero handles, 61440-byte working set, no window.
Read-only OpenProcess queries are denied. This is consistent with a retained
crash/snapshot process rather than a live workbook session, but the exact kernel
mechanism is not proven. It has not been attached to or terminated.

The harness can safely coexist with pre-existing processes: it snapshots PIDs,
requests DispatchEx, and requires its HWND PID to be absent from the baseline
and to equal the sole new PID. An existing/reused instance or concurrent new PID
fails closed. Only that verified new PID is monitored for exit; only its scratch
workbook/application is closed. Regression tests cover adoption of an old PID,
ambiguous process creation, zero PID, and a valid new PID alongside an old one.

Native resumption `20261107T064010Z-37fc793a` recorded existing PID 19684 and
independently verified new PID 33220 for REG-ADD (PASS), and 32156 for EVENTS
(BLOCKED: no callback observed after posted F9 / CalculateNow). Both owned
applications received Quit and exited; PID 19684 remained unchanged.

HELP-WIZARD was then attempted in isolated sessions. The Function Arguments
dialog was identified by title, class and owned process ID, inspected through
native child-window text and Windows UI Automation, and closed by a message
sent only to that dialog. The provider exposed the dialog title and argument
value 17, but not the expected function/argument descriptions. The result is
BLOCKED, not failed metadata registration. No screenshot, OCR, arbitrary mouse
coordinate or global keyboard input was used.

#137 remains open for causal analysis. The stale crash probe remains disabled
in ordinary sweeps. Event-delivery and Wizard acceptance remain incomplete;
cancellation races and same-process unload/retry have not been established.
x86 remains separate, with no execution or acceptance claim. No CI was run.

## Resumed sweep

Run `20261107T064311Z-f7e3b899` used the unchanged acceptance-extended.xll,
SHA-256 `f02edb333a1cd35901e502b0211fbfbd91a089b857be40741c9ece49b1cb53f0`.
The runner was held unchanged for the duration and its SHA-256 is recorded in
`native-resumed-results.json`. The XLL build provenance remains the earlier
extended fixture build, not a claim of a newly rebuilt final revision.

Results: **14 PASS, 0 FAIL, 4 BLOCKED, 0 NOT RUN** among selected catalogue entries.
Blocked: stale crash reproduction (disabled), Function Wizard descriptions
(not exposed by the tested accessibility methods), event delivery (not observed),
autonomous wake (unimplemented). All other basic probes, including protected-cell
rejection, passed. This does not complete cancellation races or same-process
unload/retry. PID 19684 still had zero threads/handles after the sweep; it was
never attached to, closed or terminated.

Validation: 13 Python tests PASS; cargo check --workspace --all-targets
--all-features PASS; cargo test --workspace --all-features PASS; cargo fmt --all
-- --check PASS; git diff --check PASS. Workspace strict Clippy FAIL on unchanged
RTD fetch_update deprecations. No CI, Office modification or x86 execution.
