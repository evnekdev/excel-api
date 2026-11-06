# Native references, typed calls and registration metadata

Implementation batch on `codex/native-references-calls-help`, starting at
`7bc669a2ebe56582cea2bee40cb6b8ac737be595` (fetched master, PR #125).
Owners: [#127](https://github.com/evnekdev/excel-api/issues/127) under #93,
[#128](https://github.com/evnekdev/excel-api/issues/128) under #94, and
[#129](https://github.com/evnekdev/excel-api/issues/129) under #95.
Implementation is locally verified, **not native accepted**. This document's
containing revision and linked implementation commit identify the new evidence.

Independent review and corrections are recorded in the final section below.
Its findings supersede the original no-result `xlSet` contract and the original
safe Rust capability-helper/command-thunk signatures. Historical verification
and the pinned 1,132-record assessment are retained.

## Sources and boundaries

Pinned header: `tools/Excel2013XLLSDK/INCLUDE/XLCALL.H`, SDK 15.0,
revision `6c8354d6536e18b58559f9cc3134f25a9e58d478`, blob
`4eb50e8fd3aea40d283181b7dcc74871ba88d854`. No vendor payload is changed.
Microsoft contracts consulted on 2026-11-07:

- [xlCoerce](https://learn.microsoft.com/en-us/office/client-developer/excel/xlcoerce): omit destination type to read scalar/nil or rectangular array values.
- [xlSet](https://learn.microsoft.com/en-us/office/client-developer/excel/xlset): contiguous destination, constant values, array broadcasting/padding/truncation, omitted value clears, command-only and undo-destructive.
- [Calling into Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/calling-into-excel-from-the-dll-or-xll): callback permissions, exact return codes, command Boolean results, and return-to-Excel obligation after abort/uncalculated dependency.
- [MTR](https://learn.microsoft.com/en-us/office/client-developer/excel/multithreaded-recalculation-in-excel): reference coercion cannot obtain uncalculated cells in MTR; xlSet is not a worksheet operation.
- [Excel recalculation](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-recalculation): xlcCalculateNow performs normal calculation across open workbooks; no forced full-rebuild contract.
- [Microsoft macro reference, compiled and hosted by Excel-DNA](https://excel-dna.net/assets/files/excel-c-api-excel-4-macro-reference-7df97f7e896ff61d528847f9124b1927.pdf), PDF pp. 46, 196â€“206: zero-argument CALCULATE.NOW and GET.DOCUMENT selector meanings. Historical primary documentation, not modern host observations.
- [xlfRegister Form 1](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1): HelpTopic is position 9; file/HTTPS resource and decimal unsigned context ID. Macro-sheet registration grants information-function permission, not command permission.
- [Known XLL issues](https://learn.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development): historical off-active-sheet array-formula xlSet issue was fixed in Excel 2007. Other Excel restrictions can still reject a target.

Excel 12 representation requires Excel 2007 or later; this project targets Windows
x64. These tests do not establish support on every Excel release, x86 or Mac.
No calculation-state query, COM fallback, formula-writing API, wake scheduler or
unrestricted callback API is introduced.

## Delivered contracts

| API | Identifier / behavior | Permission and ownership |
| --- | --- | --- |
| `ReferenceArea::cell/new`, geometry getters | inclusive zero-based coordinates, validated against 1,048,576 rows Ã— 16,384 columns | Pure Rust, no host access; invalid/reversed coordinates rejected |
| `OwnedExcelReference::current_sheet/on_sheet/from_borrowed` | SRef or one-sheet Ref geometry, 1..65,535 areas | Owns coordinates only; sheet ID must remain valid in Excel; zero explicit ID rejected |
| `read_reference(reference, limits)` | xlCoerce = 16386, one argument, omitted destination | Worksheet/MTR/macro; single rectangle; preserves errors and blanks, copies under limits and releases result once |
| `set_reference` / `clear_reference` | xlSet = 16387; reference,value / reference only; Boolean TRUE required | Actual command origin only; local arguments without DLLFree transfer; returned root checked and released once |
| `document_info(selector, name)` | xlfGetDocument = 188; integer selector, optional counted name | Command or macro-sheet-equivalent context; copy and xlFree; checked text/Boolean or Excel error |
| `calculate_now()` | xlcCalculateNow = 32799, no arguments | Actual command origin only; Boolean TRUE required, FALSE/error rejected; result owner released |
| `HelpTopic`, `FunctionHelp`, `initialize_with_help` | xlfRegister = 149, HelpTopic in position 9 | Lifecycle metadata, stable counted strings through callback, existing ID/rollback/cleanup paths |

Document selectors are closed: Name (1), Directory (2), HasChanges (4), ReadOnly
(5), ActiveSheetName (76), ActiveWorkbookName (88). The last two reject a supplied
name. Otherwise `None` means active document and a supplied name identifies an
open workbook/sheet. An unsaved directory may return Excel #N/A. No name is
implicitly the formula caller, and no volatility modifier is added automatically.

Writes accept scalar number/integer/Boolean/error/text and flat arrays. Text stays
literal. Empty clears a whole rectangle; nil array elements clear corresponding
cells. Missing arguments, zero-dimensional arrays, oversized strings and multiple
areas fail before dispatch. Array fit is Excel's documented behavior, not an exact
shape requirement; callers needing exact fit should compare rows/columns first.
Writes destroy undo history and are not transactional. Excel can reject stale
sheet IDs, protected cells or invalid array-formula targets. No silent local
geometry normalization or partial multi-area write loop occurs.

`MacroContext` also occurs in `#[excel_function(macro_type)]`. The generator now
retains that worksheet origin and the new command operations reject it before
dispatch. `#[excel_command]` grants command origin through the existing thunk
boundary. The extra constructor is doc-hidden and unsafe with a genuine command
callback precondition; it is not a general call dispatcher. Existing public
registration structs, exported names and C layouts are unchanged.
Callback tokens and CallbackScope now carry a private !Send/!Sync marker. This
intentionally prevents previously compilable, prohibited cross-thread token use;
MTR does not authorize creating a new Rust thread to invoke Excel. New error enum
variants may require downstream exhaustive matches to be extended.

## Use

```rust
use excel_api::{MacroContext, OwnedExcelReference, ReferenceArea, ExcelValue,
                ExcelCallError, ConversionLimits, DocumentInfo};

fn write_from_command(context: &MacroContext<'_>) -> Result<(), ExcelCallError> {
    let area = ReferenceArea::new(0, 0, 0, 1).map_err(ExcelCallError::Reference)?;
    let target = OwnedExcelReference::current_sheet(area);
    context.set_reference(&target, ExcelValue::Number(5.0))?;
    let _copied = context.read_reference(&target, &ConversionLimits::default())?;
    let _active = context.document_info(DocumentInfo::ActiveWorkbookName, None)?;
    context.calculate_now()
}
```

Do not use this example on user data: it overwrites A1:B1 on the current sheet.
Reference descriptors do not retain a workbook or prevent sheet deletion.
For callback-supplied U references, decode using the existing borrowed API and
copy only the geometry with `from_borrowed` when a separate descriptor is needed.
Neither this copy nor read_reference supports returning a reference to Excel.

```rust
use excel_api::{excel_function, AddInDescriptor, Runtime};

#[excel_function(name = "DEMO.HELP", thunk = "demo_help",
                 help_topic = "https://example.test/help.chm!0")]
fn demo() -> f64 { 1.0 }

static ADD_IN: AddInDescriptor = AddInDescriptor::new(
    "Demo", "Help example", &[__EXCEL_FUNCTION_METADATA_DEMO]);

// In the existing xlAutoOpen/lifecycle path:
fn initialize(runtime: &Runtime) -> Result<(), excel_api::LifecycleError> {
    runtime.initialize_with_help(&ADD_IN, &[__EXCEL_FUNCTION_HELP_DEMO])?;
    Ok(())
}
```

Without a HelpTopic attribute/supplement, existing declarations and initialize
calls are unchanged. Manual metadata uses `FunctionHelp::new("export",
HelpTopic::new("guide.chm", 42))`. Empty/malformed help is rejected; omit it to
request no help. Resource existence and Help UI behavior are host acceptance.

## Deferred capabilities and evidence

Multi-area descriptors are implemented, but read/write operations require a single
rectangle: no implicit union/coercion order or partial-update semantics are assumed
(#127). Owned reference returns, FP12 registration, address parsing and reference
arithmetic remain outside this batch. #128 does not implement full recalculation,
dependency rebuilding, calculation-state queries or a background wake mechanism
(#101). Help viewers and comprehensive command-help/shortcut authoring remain
under #95/#96. These are not placeholder APIs.

The 1,132-entry canonical catalogue and the prior Terra/Astra assessment remain
unchanged. Relevant source identities are CAPI-000001 (XLOPER12), CAPI-000003
(XLREF12), CAPI-000004 (XLMREF12), CAPI-000009 (xlCoerce), CAPI-000010 (xlSet),
CAPI-000334 (xlfGetDocument), CAPI-000758 (xlcCalculateNow) and CAPI-000011
(xlfRegister). The old assessment
is still evidence at its pinned revision, not this implementation's current
coverage. This scoped record adds revision-bound implementation/test evidence;
general current-assessment/freshness tooling remains #118. No editable field is
promoted to authenticated source completeness or native-host acceptance.

## Deferred native acceptance under #126

All scenarios are **NOT RUN / DEFERRED**, Windows Excel x64 only. Prepare a scratch
acceptance XLL using the existing minimal-XLL lifecycle/backend and these APIs;
no new runner or native execution is part of this batch. The eventual #126 Python
automation must record source/build hashes and Office versions, without screen
capture. x86 readiness/execution remains #85/#126 on a separate machine.

| Scenario / owner | Setup and observable checks | Failure / cleanup |
| --- | --- | --- |
| REF-READ / #127 | Scratch saved workbook, cells with number/text/Boolean/error/blank and 2Ã—2 range; UDF receives U reference, reads via worksheet/MTR contexts | Wrong shape/value/error, illegal dependency access, leak or crash fails; return to Excel on Uncalced/Abort; unload scratch XLL |
| REF-WRITE / #127 | Explicit command on scratch workbook only; set scalar and mixed array, verify literal text, broadcasts/padding, clear; explicit sheet ID differs from active sheet | Unexpected target, partial-array rejection mishandling, stale contents or leaks fail; discard workbook without saving, no reliance on undo |
| DOC / #128 | Two open workbooks with different active/caller sheets; named/active queries, saved/unsaved path, read-only and changed state | Conflated identity or missing #N/A fails; close scratch workbooks without changing user files |
| CALC / #128 | Scratch dirty formulas in manual mode, command invokes calculate_now | Required formulas remain stale or macro-sheet worksheet mutation succeeds: fail; restore previous calculation mode |
| HELP / #129 | XLL with present/absent file and HTTPS help supplements, supplied test help resource | Verify Function Wizard category/help/descriptions/argument help, register IDs, unload/reload and no stale registrations; inaccessible resource recorded separately from registration failure |

Local verification results and exact commit links are recorded below and on the
owning issues. Native failures must stay separate from deterministic mock outcomes.

## Local verification (2026-11-07)

Implementation commit: [`0574d59171775e18d2ec5911b9d602b1cb565435`](https://github.com/evnekdev/excel-api/commit/0574d59171775e18d2ec5911b9d602b1cb565435).
All commands run from the repository root on Windows x64. No native Excel call
was executed. Test presence and mock execution are not host acceptance.

| Exact command | Observed result |
| --- | --- |
| `cargo fmt --all` | PASS; formatting applied, line-ending-only churn removed |
| `cargo check --workspace --all-targets --all-features --offline` | PASS, exit 0; existing deprecated `fetch_update` warnings in OnTime/RTD examples |
| `cargo test -p excel-api-sys -p excel-api -p excel-api-macros -p minimal-xll --all-features --offline` | PASS, exit 0; 223 test entries: API 155, macro metadata 7, macros 7 + 1 UI harness (5 fixtures), sys 7, minimal-XLL 15, API doctests 30, sys doctest 1 |
| `cargo clippy -p excel-api-sys -p excel-api -p excel-api-macros --all-targets --all-features --offline -- -D warnings` | PASS, exit 0 |
| `cargo clippy --workspace --all-targets --all-features --offline -- -D warnings` | FAIL; existing deprecated `fetch_update` in `examples/minimal-rtd-server/src/windows_server/server.rs` lines 599, 608, 728; no lint suppression or unrelated fix |
| `cargo run --manifest-path tools/abi-check/Cargo.toml --offline` | PASS, 147 x64 MSVC SDK comparisons, including the two added identifiers and reference layouts |
| `python tools/excel-c-api-inventory/validate.py check` | PASS, 1,132 canonical records, anchors and historical identities |
| `python tools/excel-c-api-inventory/validate_rust_draft.py` | Expected stale result, exit 1: audited source/catalogue differs from its pinned revision; historical assessment unchanged |
| `git diff --check` | PASS |
| `cargo test --workspace --all-features` | NOT RUN; focused all-feature suite covers affected crates and minimal-XLL, avoiding unrelated COM/RTD/research execution |
| CI, native Excel, x86 execution | NOT RUN by task boundary; #126/#85 |

The 11 dedicated call tests cover positional arguments and tags, copied result
values, malformed roots, bounded copying, callback/transport failures, release
failures, macro-sheet rejection, and HelpTopic position preservation. Three
reference tests cover bounds, area counts and marshaling. Runtime tests cover
help mapping validation before linking, mixed registration and cleanup/rollback.
Macro metadata tests exercise generated command versus macro-sheet origins; UI
negative cases reject malformed/overflowing/duplicate help. Four compile-fail
context examples enforce thread-confinement constraints.

Remaining review limitation: workspace-wide Clippy is not green because of the
unchanged example warnings above. New production packages pass strict Clippy.
All new host scenarios remain open under the native child issues and #126.

## Independent Sol review (2026-11-07)

Starting remote master: `7bc669a2ebe56582cea2bee40cb6b8ac737be595`.
Fetched Astra branch: `4ba4b3910e67754eba23d3415c6a399779453381`, two commits
ahead and zero behind; no existing PR and no implementation already merged.
Original implementation `0574d59171775e18d2ec5911b9d602b1cb565435` and evidence
`4ba4b391` remain ancestors, preserving attribution. Review corrections:
[`064fddc34b18de845adad1f1b41b11223213218c`](https://github.com/evnekdev/excel-api/commit/064fddc34b18de845adad1f1b41b11223213218c),
branch `codex/review-native-references-calls-help`. The containing evidence
commit records the reviewed source without changing its production behavior.
PR and final merge SHA are recorded on #127/#128/#129 after remote verification.

### Findings and dispositions

| Finding at Astra tip | Source and evidence | Correction / disposition |
| --- | --- | --- |
| `xlSet` discarded its result, promoting transport success to write success | `excel_call.rs:XL_SET`, `MacroContext::set_reference/clear_reference`; Microsoft [xlSet](https://learn.microsoft.com/en-us/office/client-developer/excel/xlset) calls it command-equivalent, while [Calling into Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/calling-into-excel-from-the-dll-or-xll) distinguishes transport success from Boolean command success | Request the root; require TRUE, reject FALSE/error/unexpected values, release once using existing owner. The xlSet page's NULL-result sample intentionally discards feedback; it does not establish that transport success guarantees a write. Its result interpretation also agrees with pinned [Excel-DNA SetValue](https://github.com/Excel-DNA/ExcelDna/blob/f484749260d5cc43c35f7aea5f7a0b0f8f74a78f/Source/ExcelDna.Integration/ExcelReference.cs), supplemental evidence. Exact modern host rejection cases remain #126. |
| Safe helpers could manufacture worksheet/MTR/macro permissions; a safe generated command export could grant command origin | `thunk.rs:CallbackScope::with_worksheet_context/with_thread_safe_context/with_macro_context`, macro crate `expand_excel_command`; Microsoft call restrictions and [registration flags](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1) require the actual authorized callback class/thread | Capability-granting helpers and generated command exports now require unsafe caller contracts. Generated worksheet exports already were unsafe and now document the registration permission precondition explicitly. Ordinary annotated functions, metadata, export names, parameters, result widths and calling convention are preserved. Hidden Rust helper/thunk callers must add justified unsafe boundaries; this is a deliberate safety compatibility tightening. |
| Resource validation accepted an empty HTTPS authority and control/whitespace malformed resources | `registration.rs:HelpTopic::validate`, macro crate `expand_excel_function`; Microsoft HelpTopic syntax requires a file or HTTPS resource | Reject empty authority, URL whitespace/backslashes and resource control characters in both paths. File paths with spaces remain accepted; decimal u32 and counted UTF-16 bounds remain intact. This is bounded syntax checking, not a complete URL validator or resource availability guarantee. |
| New write tests interfered with process-global allocation counters | First independent focused run failed `return_alloc::tests::local_drop_and_repeated_construction_release_exactly_once`: live roots `(2,2,1)` versus `(1,2,1)`, 158 passed / 1 failed | Share the existing allocation-test lock with batch tests that materialize arguments; production allocation behavior is unchanged. Normal parallel focused and workspace reruns pass. |
| Selector test derived its expected number from the same implementation | `excel_call_batch_tests.rs:typed_document_selectors_validate_positions_results_and_names` | Use independent literal oracle pairs 1/2/4/5/76/88 from the historical Microsoft reference. |

### Reference ABI, ownership and callback audit

The pinned header declares WORD count and `XLREF12 reftbl[1]` with an explicit
trailing-count comment (`XLCALL.H:66-87`). On the supported Windows x64 ABI,
the table offset is 4, area size 16, and both structures align to 4. The owned
allocation has exactly `4 + 16*n` bytes for `1 <= n <= 65535`, at most
1,048,564 bytes; its `1 + 4*n` u32 element calculation fits usize. Compile-time
assertions now bind the Rust allocation to those layout facts, independently
compared with the SDK by the x64 checker.

Both the header pointer and every area pointer derive from one pointer to the
complete allocation. No Rust reference or one-element slice to `reftbl` is used
to write trailing areas. Count and padding are initialized, coordinates are
validated before raw construction, and the stable box remains owned through the
synchronous callback. Moving the owner does not move its backing storage. The
root is passed only after reaching its call-local location. This lifetime and
provenance argument is separate from size/alignment test success.

Maximum-count regression decodes the last area, copies all geometry, drops the
original raw storage/views, then builds and decodes an independent argument.
Negative tests cover count zero, invalid SRef count, reversed trailing geometry,
zero explicit sheet ID, bounds, empty/oversized area sets. Rust-owned descriptors
remain Send + Sync; callback tokens/views and CallbackScope remain neither.
Separate Send and Sync compile-fail cases now test each constraint independently.
No x86 ABI or host compatibility is inferred from this x64 evidence.

SRef means the current callback sheet, not a captured active/front sheet.
An explicit sheet ID is an Excel identity whose continued validity cannot be
proved locally. Neither descriptor form owns values, retains Excel pointers,
holds a workbook open, or supports owned reference returns.

`read_reference` sends one argument to xlCoerce=16386, omits target type, rejects
multi-area I/O, preserves nil/errors and copies scalars/rectangles under limits.
The existing owner releases one successful root on copy success or failure.
Transport failure creates no result-release obligation; Microsoft specifies an
error result on callback failure. `xlretUncalced`/`xlretAbort` remain exact failures;
callers must return to Excel before any further C API call except xlFree.
Worksheet/MTR use does not authorize uncalculated dependency access.

Writes/clearing and calculate_now require command origin before dispatch.
Macro-sheet UDFs can query document information but cannot write or recalculate.
Generated thunks select their matching origin; the unsafe construction boundary
prevents ordinary safe Rust from manufacturing a stronger context on another
thread. This is a callback permission correction, not scheduling infrastructure.
Temporary reference/value arguments remain alive through dispatch and result
copy/release. No DLLFree transfer or xlFree of Rust arguments occurs. A reported
error can follow a side effect (including result-release failure); do not assume
rollback, retry automatically, or rely on undo history.

### Typed calls, registration and complete changed-file review

The selected GET.DOCUMENT values and name semantics match the historical
Microsoft macro reference, PDF pages 196-197 and 205-206 (one-based):
1 is document name, 2 directory (#N/A when unsaved), 4 changed, 5 read-only,
76 active sheet, 88 active workbook. Omitted name means active workbook;
selectors 76/88 deliberately reject a supplied name in this bounded wrapper.
Boolean/text/error results are checked and copied/released. These historical
contracts are not observations of every modern Excel version.
CALCULATE.NOW's zero-argument form is PDF page 46: normal calculation across
open workbooks, not forced full dependency rebuilding. SDK IDs 188 and
31|xlCommand=32799 agree with the independent C/Rust checker.

HelpTopic remains registration argument 9, between shortcut and function help.
Missing help, descriptions, argument help, macro flags, command category and
registration ordering remain intact. Existing public registration struct
literals and macro syntax compile. Duplicate/unknown/ambiguous supplements fail
before backend linking; registration IDs, rollback and close/reopen use the
existing runtime paths. Macro negative fixtures now check all four capability
grants plus safe command-thunk invocation, alongside existing attribute,
signature, symbol-collision and HelpTopic diagnostics. Six macro UI fixtures
pass with snapshots verified normally after initial snapshot generation.

Every production file in Astra's diff was inspected: reference geometry/storage;
call descriptors, arguments, return codes and release; context/thunk origin and
thread confinement; runtime supplement validation and lifecycle; registration
metadata; return argument accessor; public exports; sys IDs; macro parsing and
generation. ABI-checker changes are only the two numeric IDs. Architecture,
support matrix, catalogue README, all modified/new tests and evidence were also
reviewed. No generated SDK payload, raw ABI struct, Cargo dependency/lockfile,
canonical identity, historical assessment, unrelated RTD code or second evidence
database changed. The minimal-XLL delta only updates hidden command-thunk Rust
function-pointer safety assertions and its isolated test call.

Pre-existing lifecycle limitations were recorded separately on
[#89](https://github.com/evnekdev/excel-api/issues/89#issuecomment-6698155884),
related #92/#96/#124: unregister/delete-name currently check transport and release
but not Boolean/error acceptance; successful registration followed by xlFree
failure can lose the newly created ID before Runtime records it; lifecycle entry
methods retain their documented genuine-callback precondition. These exist on
starting master and were not changed or represented as solved by HelpTopic.
They are separately owned baseline lifecycle work, not regressions in this batch;
complete host-cleanup guarantees remain unclaimed. No lifecycle-order redesign
or scheduling work was included.

### Independently executed verification

Windows x64 MSVC, rustc 1.99.1. Commands use the review tree; baseline commands
use a separate named worktree at starting master, without detached HEAD.

| Exact command | Observed result |
| --- | --- |
| `cargo fmt --all` | PASS; applied, no formatting-only drift |
| `cargo check --workspace --all-targets --all-features --offline` | PASS; unchanged RTD/OnTime deprecation warnings; a temporary new test unused-mut warning was removed |
| `cargo test -p excel-api-sys -p excel-api -p excel-api-macros -p minimal-xll --all-features --offline` | Initial FAIL due to allocation-counter test race above; after correction PASS 233 entries: core 159, metadata 7, macros 7 + UI harness 1, sys 7, minimal-XLL 15, API doctests 36, sys doctest 1 |
| `cargo test --workspace --all-features --offline` | PASS, 358 entries; 30 live host tests ignored, none executed |
| `cargo test -p excel-api-macros --test trybuild --offline` | PASS, six fixtures with ordinary snapshot verification (no overwrite) |
| `cargo clippy -p excel-api-sys -p excel-api -p excel-api-macros --all-targets --all-features --offline -- -D warnings` | PASS after correcting safety-comment placement in three new test assertions; initial lint failure preserved in review history |
| `cargo clippy --workspace --all-targets --all-features --offline -- -D warnings` | FAIL, exit 101: unchanged RTD `server.rs:599/608/728` and optional OnTime `ontime_experiment.rs:90` deprecated fetch_update on this toolchain |
| Same workspace Clippy command at master `7bc669a2` | FAIL, exit 101: same OnTime deprecation; build stopped before RTD diagnostics |
| `cargo clippy -p excel-api-minimal-rtd --all-targets --all-features --offline -- -D warnings` at master `7bc669a2` | FAIL, exit 101: the same three RTD deprecations; both failing files byte-identical to review tree. Explicit non-blocking baseline disposition; no suppression or unrelated fix |
| `RUSTDOCFLAGS="-D warnings -D rustdoc::broken_intra_doc_links" cargo doc -p excel-api-sys -p excel-api -p excel-api-macros --all-features --no-deps --offline` | PASS (environment set with PowerShell) |
| `cargo run --manifest-path tools/abi-check/Cargo.toml --offline` | PASS, all 147 x64 SDK comparisons |
| `python tools/excel-c-api-inventory/validate.py check` | PASS, 1,132 records, pinned declarations and reachable ID history |
| `python tools/excel-c-api-inventory/validate_rust_draft.py` | FAIL / STALE, exit 1: source differs from historical audited revision; correctly not claimed current, #118 remains owner |
| `git diff --check` and protected-path diff against starting master | PASS; historical assessment, canonical catalogue, SDK and raw structs unchanged |
| CI, native Excel, x86, installers and release publication | NOT RUN by explicit boundary |

Source-level acceptance is satisfied after the corrections, with the documented
pre-existing Clippy/lifecycle dispositions. No new unresolved reference ABI,
lifetime, callback permission, write-result or HelpTopic blocker remains.
Native acceptance is **NOT RUN / DEFERRED**, not waived or inferred from mocks.
#127/#128/#129, subsystem parents #93/#94/#95 and programme #78 remain OPEN.
Native GitHub parent links now verify #78 -> #93/#94/#95 -> #127/#128/#129.
#126 retains REF-READ, REF-WRITE, DOC, CALC, HELP and #124 command lifecycle
scenarios for the later consolidated Python/scratch-workbook x64 sweep with the
async/scheduling/events batch. x86 remains separate under #85/#126.
