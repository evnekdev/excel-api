# Native references, typed calls and registration metadata

Implementation batch on `codex/native-references-calls-help`, starting at
`7bc669a2ebe56582cea2bee40cb6b8ac737be595` (fetched master, PR #125).
Owners: [#127](https://github.com/evnekdev/excel-api/issues/127) under #93,
[#128](https://github.com/evnekdev/excel-api/issues/128) under #94, and
[#129](https://github.com/evnekdev/excel-api/issues/129) under #95.
Implementation is locally verified, **not native accepted**. This document's
containing revision and linked implementation commit identify the new evidence.

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
| `set_reference` / `clear_reference` | xlSet = 16387; reference,value / reference only | Actual command origin only; local Rust storage reused without DLLFree transfer; no result root or xlFree call |
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
