# Native acceptance attempt, 2026-11-07

Tracking: #126, #85, [registration investigation #136](https://github.com/evnekdev/excel-api/issues/136). This is an incomplete acceptance attempt, not native acceptance. No CI, x86 Excel execution, or visual automation was used.

## Environment and builds

- Source under test: fetched master `26360e10cde0de5cb1d2d0f57121a3d2e6822ea0`; release `minimal-xll`, built using `powershell -NoProfile -File scripts/build-minimal-xll.ps1 -Profile release`.
- XLL SHA-256: `6bc798bf218e40e79ab973dfafd46be02e95bee48c20773d92d061211be0c393`; PE machine x64. The build command ran from that checkout; the binary does not independently attest to its source revision.
- Windows `10.0.26300`, Excel PE x64, Excel `16.0` build `20527.0`, EXCEL.EXE file version `16.0.20527.20150`; Python 3.11.7 x64 with pywin32; Rust `1.99.1 (58cf8f9c3 2026-10-13)`.
- Comparison: a release XLL built from a `git archive` of earlier merge `d5ef74b0e4ccd097d72cea42e47eaa2e1ba0d70e`, SHA-256 `658042579b2de1814efd35f9f7a5548ed6700b1127de6119e8651f79b94a2b44`, gave the same `RUST.ADD` result on this host. This does not identify the cause and does not exonerate either build.

## Observation

Several isolated sessions returned success from `Application.RegisterXLL`. In a disposable workbook the built-in `=SUM(2,3)` returned 5, while `=RUST.ADD(2,3)` returned Excel error 2029 (`#NAME?`); `RegisteredFunctions` inspection found no `RUST.ADD` row. This reproduced with both XLLs. `RUST.ECHO` produced the same Excel error. A separate command attempt returned Excel's ambiguous message that `RUST.PING.COMMAND` was unavailable or macros were disabled. Root cause remains unknown: host configuration/security and XLL initialization are both possible. No Office security policy was changed. The local release XLL had no `Zone.Identifier` stream.

The first COM activation stalled before it returned an application handle, and its Python probe was stopped; the Excel process subsequently exited. In a later consolidated attempt a worker timed out after logging that it requested Quit; its identified Excel process subsequently exited. No Excel process was force-terminated. Another earlier probe exposed a harness defect in restoring an untouched calculation setting; the harness was corrected to restore only changed settings and always attempt Quit. A console encoding failure after the consolidated JSON and summary were written was also corrected. The timeout's observation is not counted as a successful feature assertion.

## Consolidated run

Run `20261107T053948Z-527cd23a`: 15 selected catalogue entries, **0 PASS, 2 FAIL, 5 BLOCKED, 8 NOT RUN**. This reflects that particular run; a later separate command probe remained BLOCKED. A failed assertion is an observed host result, not a determination of product causation.

| ID | Expected | Observed | Status |
|---|---|---|---|
| REG-ADD | 5 | #NAME? | FAIL |
| VALUE-TEXT | Unicode round trip | #NAME? | FAIL |
| VALUE-OPTION | missing | Worker timed out during teardown; partial log showed #NAME? | BLOCKED |
| REF-KIND | reference argument | Previous Excel process still present | BLOCKED |
| COMMAND | ping command runs | Previous Excel process still present; separate retry: unavailable or macros disabled | BLOCKED |
| QUEUE | queue then explicit pump | Previous Excel process still present | BLOCKED |
| ASYNC | double completes | Previous Excel process still present | BLOCKED |
| REF-READ | xlCoerce reads | Fixture not implemented | NOT RUN |
| REF-WRITE | xlSet writes/clears | Fixture not implemented | NOT RUN |
| DOC | document information | Fixture not implemented | NOT RUN |
| CALC | normal command recalculation | Fixture not implemented | NOT RUN |
| HELP | HelpTopic metadata | Fixture not implemented | NOT RUN |
| EVENTS | calculation events | Fixture not implemented | NOT RUN |
| CLEANUP | unload and registration cleanup | Fixture not implemented | NOT RUN |
| WAKE | autonomous wake | Not implemented (#30/#101) | NOT RUN |

The JSON report and individual logs for this run were generated under local `target/native-acceptance/20261107T053948Z-527cd23a/`; the comparison run was `20261107T054325Z-3fda203f`. These ignored local files are not a permanent hosted artifact. This note preserves the essential results. The runner is deliberately a partial portable foundation; scenario IDs do not imply the unimplemented coverage exists. x86 XLL support, ABI correctness and native execution remain unverified under #85.
