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

## Independent Sol review (2026-11-07)

Starting fetched master: `d5ef74b0e4ccd097d72cea42e47eaa2e1ba0d70e`.
Fetched Astra tip: `065efa907523224fce15a7004b2aaaa60a5cdeab`, four commits
ahead / zero behind master, no existing PR, and no part of this batch already
merged. Review branch: `codex/review-async-dispatch-lifecycle`.
Corrections and seven additional regressions:
[`4f7e7eff3b8fd2e1866a7b1ce702f4ac78f50490`](https://github.com/evnekdev/excel-api/commit/4f7e7eff3b8fd2e1866a7b1ce702f4ac78f50490).
Astra's two implementation and two evidence commits remain ancestors, without
rewrite. The containing evidence commit identifies this review record; final
PR/merge SHA and remote-tree verification are recorded on #131-#134.

### Findings, corrections and regression evidence

| Severity / owner | Finding at Astra tip | Correction and evidence |
| --- | --- | --- |
| High / #131, #133 | `ExcelCallError::requires_return_to_excel` recognized only direct ExcelFailure. Abort/Uncalced from xlFree was nested in release/conversion errors, so rollback, close or a drain could issue another prohibited callback. | Inspect all release-bearing error forms, including compound conversion/release and rejected cleanup. Keep raw bit-combined codes. Two runtime regressions failed before correction: close continued after release Abort, and registration-release Abort rolled back to Uninitialized instead of retaining the accepted ID for a later callback. Both now pass. Predicate and queue tests also cover combined codes, non-stop failures, accepted/rejected cleanup, no duplicate release or accepted side effect. Microsoft [Calling into Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/calling-into-excel-from-the-dll-or-xll) requires returning control after either status; the fact that the failing operation was xlFree does not permit a subsequent unregister/write. |
| Medium / #133 | `drain` protected the running request but not its selected tail on unwind. Remaining SELECTED requests were absent from the queue and could remain pending until explicit cancellation/close. | Private SelectedBatchGuard uses the existing requeue CAS on normal early return and unwind. `unwind_requeues_the_unstarted_selected_tail` failed before correction (pending 0 versus 2) and passes after it. The running guard still publishes one failure; the untouched tail executes later. No queue or notification subsystem was replaced. |
| Regression coverage / #132 | The existing post-materialization cancellation test did not demonstrate shutdown joining a worker already holding an allocated result. | A barrier-observed wrapper around the real ThreadPoolExecutor pauses that worker before Returning, commits controller shutdown/cancellation, proves close has not finished, then releases the worker. No callback occurs; close joins it, capacity reaches zero, and the pool is Closed. Async and write tests which materialize results share the existing allocation-counter test lock; production allocation code is unchanged. |
| Independent oracle / #134 | The event expectation used the implementation constants directly. | The registration/reopen regression now expects literal canceled=2, ended=1 from pinned XLCALL.H:317-318. The SDK, constants and event ABI remain unchanged. |

The new queue regression stops after a command result's xlFree Abort/Uncalced,
preserves the failed ticket's error, returns selected work in order ahead of work
enqueued during the drain, verifies capacity accounting, cancels requeued work,
and proves no subsequent Excel callback executes. Cancellation and shutdown
permutations plus repeated requeue verify that terminal requests are neither
duplicated nor resurrected. These supplement the existing barrier-controlled
Selected/Running shutdown, generation, reentrancy, expiry and panic tests.

### Independent lifecycle and concurrency assessment

Registration results are inspected before release. A valid numeric ID survives
release failure exactly as IEEE-754 bits and enters the cleanup obligation list;
neither an ID nor successful host state is invented for malformed/ambiguous
results. Name deletion and unregister have independent pending flags. Boolean
TRUE clears only its accepted obligation, even if release fails; FALSE/error or
malformed feedback retains the unresolved obligation. Both rollback and close
use the same loop. CleanupRequired prevents reopen; a later close retries only
unresolved work. An error with all side effects resolved can require one later
close to acknowledge it, without repeating those calls. The new nested-status
handling stops that loop even when the status originated in result release.

Async admission/acceptance and shutdown share the controller mutex; workers
cannot start under that lock. Running->Completing accepts one completion attempt.
Cancellation's CAS from Completing competes with Completing->Returning; only
the latter commits publication, after preparation. A cancellation which wins
suppresses publication. Preparation failure also retires without publication.
Shutdown deactivates the old controller, requests cancellation and joins accepted
executor work before backend unlink. A return whose permission check precedes
shutdown can finish while close waits, but cannot outlive the executor join.
The retired bit and identity-checked registry removal bound retirement to once.
No controller/global mutex is held across Excel calls or executor joins.

Diagnostics count accepted async return calls, preparation/backend failures,
retained requests and notifications separately. Relaxed counters plus the request
mutex provide an observational snapshot, not a simultaneous transaction or durable
cell-delivery receipt. Task failures mapped to an Excel error can still count as
accepted returns. Controllers own their counters; old captured Rust controllers
do not affect replacements. Workers use only xlAsyncReturn, retain Rust-owned
result storage across that synchronous call, and never transfer it with DLLFree
or free the opaque handle. Existing void/X metadata and generated thunks are
unchanged. Cooperative executor contracts and indefinitely blocked user tasks
remain limits, not newly solved problems.

Dispatcher selection/start commit under its mutex, checking active generation
and request identity. Registry accounting includes queued/selected/running work.
Cancellation can win before Running; shutdown deactivates before retirement;
requeue operates on that same controller and CASes only SELECTED members. The new
batch guard cannot redirect work to another generation. Command requirements
permit only genuine command origin; MacroSheet, worksheet, MTR and lifecycle
cannot run those operations. Context-neutral work remains permitted. The sealed
catalogue owns arguments and reuses PR #130 typed calls; it accepts no arbitrary
Excel IDs or context closures. A drain's return_to_excel_required is caller
guidance to return, not permission to immediately redrain or a global poison gate.

Event registration matches [xlEventRegister](https://learn.microsoft.com/en-us/office/client-developer/excel/xleventregister):
procedure string, integer event, two arguments and positive integer acceptance.
No invented removal operation is used. Registered procedures remain parameterless,
panic-contained and make no Excel calls or dispatcher cancellation. Inactive Rust
controllers ignore new notifications; counters and requests remain generation
owned. A host callback delivered after reopen has no epoch and cannot be proved
stale: preserve that native #126 requirement. Microsoft [Handling Events](https://learn.microsoft.com/en-us/office/client-developer/excel/handling-events)
says programmatic recalculation does not raise these notifications; no universal
calculation-completion barrier is claimed.

Hidden-name deletion's documented host issue, unusable registration IDs, ambiguous
transport failure, unrecoverable allocations after failed release and xlAutoClose
unload control remain explicit limits. The existing public lifecycle methods
still rely on their genuine-callback contract; no broad constructor redesign was
made. Native cleanup and unload success remain unverified.

### Complete diff, architecture and compatibility

Reviewed all ten original changed files and surrounding runtime/ownership/context
code: async_udf, dispatcher, excel_call, runtime, dispatch_demo, three architecture
documents, SUPPORT_MATRIX and this evidence. Review corrections add only focused
tests in the existing call-test module. Public API snapshots from starting master
and review were independently generated with cargo-public-api, in separate named
worktrees, and compared. No public item/signature was removed or changed. Added
AsyncDiagnostics/function, cleanup error variants, command operations/results,
MacroSheet and report field are visible in that diff. **Enum additions break
exhaustive downstream matches, and the report field breaks complete struct literals**;
they are not described as universally source-compatible. Minimal-XLL's display
match handles CommandAccepted, and the full workspace compiles/tests. No release
compatibility guarantee for every external consumer is inferred.

The same three core crates remain sufficient. No host adapter, COM dependency,
dependency/feature change, export/ABI change, duplicate runtime, cancellation
registry or autonomous wake was added. PR #130 capability safety, confinement
and ownership protections remain intact. H3/Application.Run remains the preferred
future adapter topology, with #30/#101 unresolved; this batch implements no adapter.

### Independent verification and CI investigation

Windows x64 MSVC, stable rustc 1.99.1; declared core MSRV checked with installed
Rust 1.85.0. Initial sandbox execution denial for that older compiler was resolved
by scoped execution permission; the completed results below are actual runs.

| Exact command / inspection | Independent result |
| --- | --- |
| `cargo fmt --all` | PASS; formatting applied |
| `cargo check --workspace --all-targets --all-features --offline` | PASS, exit 0; unchanged optional example warnings |
| `cargo test -p excel-api --all-features --offline release_abort -- --nocapture` before correction | FAIL, both new lifecycle regressions; corrected rerun without nocapture PASS |
| `cargo test -p excel-api --all-features --offline unwind_requeues_the_unstarted_selected_tail -- --nocapture` before correction | FAIL, selected tail stranded; corrected rerun PASS |
| `cargo test -p excel-api --all-features --offline` | PASS: 180 core unit, 7 metadata, 36 doctest entries |
| `cargo test --workspace --all-features --offline` | PASS: 379 entries, 0 failures, 30 live-host tests ignored; includes macro/trybuild/minimal-XLL tests |
| `cargo clippy -p excel-api-sys -p excel-api -p excel-api-macros --all-targets --all-features --offline -- -D warnings` | PASS |
| `cargo clippy --workspace --all-targets --all-features --offline -- -D warnings` | FAIL, exit 101: unchanged OnTime:90 and RTD server.rs:599/608/728 deprecated fetch_update |
| Same workspace Clippy command in baseline worktree at d5ef74b | FAIL, same OnTime:90; stopped before RTD diagnostics |
| `cargo clippy -p excel-api-minimal-rtd --all-targets --all-features --offline -- -D warnings` in baseline worktree | FAIL, same three RTD deprecations; failing files byte-identical to review |
| `cargo +1.85.0 check --workspace --all-targets --all-features --offline` in baseline worktree | FAIL, exit 101: unchanged inventory validate.rs let chains at 41/414/420, matching CI |
| `cargo +1.85.0 check -p excel-api-sys -p excel-api -p excel-api-macros --all-targets --all-features --offline` | PASS; no core MSRV regression found |
| `cargo public-api -p excel-api --all-features --simplified --color never` in both named worktrees, text diff | PASS generation/comparison; documented additions only, no detached checkout |
| `RUSTDOCFLAGS="-D warnings -D rustdoc::broken_intra_doc_links" cargo doc -p excel-api-sys -p excel-api -p excel-api-macros --all-features --no-deps --offline` | PASS; environment set in PowerShell |
| `git diff --check` and protected-path diff against d5ef74b | PASS; Cargo manifests/lockfile, macros/sys, SDK, ABI-checker and canonical/historical evidence untouched |
| SDK ABI checker / C API catalogue validator | NOT RUN as requested: all their inputs/dependencies unchanged; this is not a new ABI/catalogue PASS |
| New CI execution, native Excel, x86, installers/releases | NOT RUN |

The earlier unintended [CI run 42313476068](https://github.com/evnekdev/excel-api/actions/runs/42313476068)
**did occur**, on Astra evidence tip 9b8facb4. Independently fetched both job logs:
MSRV job 128220323298 failed Rust 1.85 with E0658 at inventory validate.rs:41/414/420;
validate job 128220323597 passed formatting and failed Clippy at OnTime:90.
Later CI tests, ABI, export/build/package steps were skipped. The available logs
support baseline code/toolchain failures, not a CI test-suite PASS. No native Excel
execution is evidenced. Local workspace tests are a separate independently
passing run. No CI rerun, workflow edit or unrelated MSRV/deprecation repair was
performed; review and merge commits use [skip ci]. The unresolved workspace-wide
MSRV failure remains accurately disclosed despite core MSRV passing.

Source-level acceptance is satisfied after these corrections, with explicit
non-blocking baseline toolchain failures. Native acceptance is not waived.
Native hierarchy is verified as #78 -> #89/#99/#101/#102 -> #131/#132/#133/#134;
all remain OPEN. #126 retains combined x64 load/unload/reload, command/shortcut,
reference/document/recalc/help, async, queue, event, partial cleanup/retry,
shutdown and stale-host-event scenarios. Add release-origin Abort/Uncalced and
selected-tail recovery where host behavior can be exercised; mock-only injection
must remain labeled synthetic. Autonomous wake remains BLOCKED under #30/#101;
x86 remains separate #85/#126. The next session must use authorized Python
application automation and isolated scratch workbooks; none ran here.
