# Async, dispatch, events and lifecycle batch

Baseline: `d5ef74b0e4ccd097d72cea42e47eaa2e1ba0d70e`, fetched master,
including PR #130 and Sol correction `064fddc3`. No open PRs at orientation.
Branch: `codex/async-dispatch-lifecycle`. Inspection date: 2026-11-07.
Native sub-issues: #131 under #89, #132 under #99, #133 under #101, #134 under
#102. Native acceptance remains #126; all these issues and #78 stay open.

Implementation commits:
- [23d61920fa3b74207a6ace432efe9ca124bc037e](https://github.com/evnekdev/excel-api/commit/23d61920fa3b74207a6ace432efe9ca124bc037e): lifecycle cleanup and event validation (#131/#134).
- [035ce479f066b4d27697c34fb42646994ecc89aa](https://github.com/evnekdev/excel-api/commit/035ce479f066b4d27697c34fb42646994ecc89aa): async, event diagnostics and typed queue (#132/#133/#134).

## Source contracts and decisions

- [Microsoft native async](https://learn.microsoft.com/en-us/office/client-developer/excel/asynchronous-user-defined-functions): void/X registration, deep-copy inputs, xlAsyncReturn-only worker callback. Generated ABI/executor retained.
- [xlEventRegister](https://learn.microsoft.com/en-us/office/client-developer/excel/xleventregister): two event IDs, procedure string, positive integer acceptance. No documented removal operation or host generation tag.
- [Handling events](https://learn.microsoft.com/en-us/office/client-developer/excel/handling-events): cancellation followed by end; **not raised during programmatic recalculation**. Not a universal calculation-completion barrier (#84).
- [xlfUnregister Form 1](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfunregister-form-1): Boolean acceptance, separate hidden-name cleanup.
- [xlfSetName](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfsetname): omitted definition deletes; Boolean/error result. Documented hidden-name deletion limitation remains a host risk.
- [Calling into Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/calling-into-excel-from-the-dll-or-xll): transport success is not operation acceptance; Abort/Uncalced requires returning control except xlFree; thread identity alone is not permission.
- [Excel-DNA NativeAsyncTaskUtil](https://github.com/Excel-DNA/ExcelDna/blob/f484749260d5cc43c35f7aea5f7a0b0f8f74a78f/Source/ExcelDna.Integration/NativeAsyncTaskUtil.cs): calculation cancellation-token lifecycle; not proof of RTD-backed Run or PyroApp parity (#82).
- [Excel-DNA synchronization](https://github.com/Excel-DNA/ExcelDna/blob/f484749260d5cc43c35f7aea5f7a0b0f8f74a78f/Source/ExcelDna.Integration/ExcelSynchronizationContext.cs): posted request -> COMRunMacro/Application.Run -> registered SyncMacro. The message signals work; Excel must grant the callback.

**Wake/crate decision:** no autonomous adapter or excel-api-host crate.
Application.Run is a source-backed candidate, but a production Office resource,
retry, security and unload protocol is not established here. OnTime research stays
inconclusive. Timers/RTD/async returns cannot manufacture MacroContext. The existing H3 preferred Application.Run topology is preserved; no roadmap
reversal or competing adapter is proposed. Cooperative
work remains useful without guaranteed latency. #30/#101 remain open.

## Implementation and regression evidence

| Owner | Delivered behavior | New regressions |
| --- | --- | --- |
| #131 | Read accepted registration ID before release; independent name/ID cleanup obligations; shared rollback/close loop; preserve Boolean/tag/error/release evidence; defer cleanup on Abort | Function/command release failure, FALSE/error/nil cleanup, no repeat after accepted cleanup, partial rollback, compound release error, abort stops cleanup |
| #132 | Cancellation through preparation until Returning commitment; bounded public async diagnostics; emit existing diagnostic on completion failure | Barrier-controlled cancel after allocation; late/duplicate completion; accepted vs backend/conversion failure counters |
| #133 | Owned SetReference/ClearReference/DocumentInfo/CalculateNow reuse typed calls; command-origin selection; CommandAccepted; stop/requeue on Abort/Uncalced | All weaker origins dispatch nothing; SDK IDs/counts; cancel/shutdown; Boolean rejection; invalid metadata; abort/uncalced stop batch |
| #134 | Validate documented event metadata; active-generation event counters and async-only cancellation | Invalid ID/empty/NUL before dispatch; inactive old controller; cancel/end counts and request retirement |

14 new tests supplement existing races, panic handling, event registration failure,
duplicate/reopen behavior, queue expiry/capacity/reentrancy and macro tests. Existing
typed-call tests verify positional values and ownership below the new queue layer.
Mocks establish Rust behavior, not native acceptance.

No ABI layouts, exports, dependencies, vendor files, catalogue IDs or historical
assessments changed. Genuine callback preconditions, thread confinement, deep-copy
and release boundaries remain. New enum variants/report field require exhaustive
matches/struct literals to be extended; minimal-XLL's display match is updated.

## Limits and uncertainty

- Workers/custom executors must cooperate and join accepted work on shutdown.
  No new executor, forced thread termination or abandoned-executor recovery.
- Cancellation after Returning commits cannot retract a permitted callback.
  Close still joins accepted jobs before unlinking.
- Diagnostics are observational active-generation counters, not durable receipts
  or an atomic snapshot of all concurrent state.
- Parameterless host events contain no epoch. Captured old Rust controllers are
  isolated, but a delayed host event after reopen cannot be authenticated as old.
  No invented unregistration or unload safety claim.
- Events neither drain nor cancel unrelated dispatcher work. Programmatic recalc
  may emit no events; no CalculationState or transaction barrier is introduced.
- Result release is attempted once. Failure can leave an Excel allocation leaked;
  reporting it does not repair it. A transport failure/malformed registration with
  no usable ID remains ambiguous; no ID or automatic registration retry is invented.
- Runtime's existing public lifecycle methods still require a genuine lifecycle
  callback by contract. No broader public constructor redesign occurred.
- Consumers must honor return_to_excel_required, return rather than redrain, and
  respect queued sheet identity/lifetime. Writes remain undo-destructive.

## Verification

Windows x64, installed dependencies in offline mode. Final commands:

| Command | Result |
| --- | --- |
| `cargo fmt --all` | PASS; formatting applied; newline-only churn removed |
| `cargo check --workspace --all-targets --all-features --offline` | PASS, exit 0, unchanged example warnings |
| `cargo test --workspace --all-features --offline` | PASS, exit 0; 372 passed, 0 failed, 30 live-host tests ignored; includes 173 core units, metadata, trybuild and Rustdoc |
| `cargo clippy -p excel-api-sys -p excel-api -p excel-api-macros --all-targets --all-features --offline -- -D warnings` | PASS, exit 0 |
| `cargo clippy --workspace --all-targets --all-features --offline -- -D warnings` | FAIL, exit 101; unchanged RTD fetch_update deprecations at server.rs 599/608/728; no suppression/unrelated fixes |
| `git diff --check` | PASS |
| SDK ABI checker | NOT RUN; no raw IDs/layout/export signatures changed |
| Catalogue validator | NOT RUN; catalogue/source/inventory tools unchanged |
| CI | Unintentionally triggered by branch push; see correction below |
| Native Excel, x86 | NOT RUN by task boundary |

An intermediate compile exposed minimal-XLL's exhaustive DispatchResult match;
fixed under #133 before the final passing run. Workspace Clippy failure was
independently reproduced, not assumed from historical reports.

## Native handover (#126)

All scenarios **NOT RUN**, Excel x64 only. Later Python automation must use scratch
workbooks, no screen capture, source/build hashes, Office version/bitness, bounded
logs, and restore settings. x86 remains #85 on another machine.

| Scenario | Observations and cleanup |
| --- | --- |
| ASYNC | void/X registration, multiple functions and value kinds; cancel before/during work; late completion/reopen; no stale publication/crash/leak; join workers before unload |
| QUEUE | Worker enqueues; no effect before genuine command pump; correct explicit sheet; weaker contexts cannot mutate; capacity/expiry/cancel/reentrancy/shutdown; return on Abort/Uncalced |
| EVENTS | Both exports once; user cancel/end counts and async effects; separately record absent programmatic events; close/reopen/unload callback lifetime |
| CLEANUP | Mixed function/command partial initialization; rejected name deletion/unregister; release diagnostics where controllable; retry only unresolved obligations; no stale registrations on reload. Forced mock failures are not presumed reproducible in Excel |
| Existing #124/#127/#128/#129 | Retain load/unload/reload, shortcuts, REF-READ, REF-WRITE, DOC, CALC and HELP from native-c-api-batch.md |
| Autonomous wake | BLOCKED/design unresolved #30/#101; not a runnable production scenario |

Ready for independent source review, not native accepted. Keep leaves and parents
open pending their review/host requirements. No merge or next implementation phase.

## Push-triggered CI correction

The first push unintentionally triggered the repository's unconditional branch-push
workflow: [run 42313476068](https://github.com/evnekdev/excel-api/actions/runs/42313476068).
The agent omitted a skip marker despite the no-CI instruction. It was already
completed when discovered; it was not manually dispatched or rerun. Earlier issue
comments saying CI NOT RUN are corrected by this record and follow-up comments.
MSRV failed on existing let-chain syntax in the unchanged object-model inventory
validator (lines 41, 414, 420). The validate job passed formatting then failed
Clippy on the unchanged OnTime fetch_update deprecation (line 90); later test,
ABI and packaging steps were skipped. No native Excel execution occurred.
This evidence-only corrective commit uses [skip ci], with no workflow changes.
