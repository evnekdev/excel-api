# Independent technical review of the C API to Rust inventory

Review date: 2026-11-07. Repository: `evnekdev/excel-api`. Tracking: [#122](https://github.com/evnekdev/excel-api/issues/122), child of #90; programme #78.

The reviewed remote branch `codex/edna02-source-rust-audit` was exactly
`1c1bfdf419ca619cfd97dd0e0d097f1c4d3593a2` after fetch. Audited Rust and current
remote master were `172b306fae9b8496d5ad79b626a9fa0b1890cfd4`. The worktree was
clean. Corrections are on `codex/edna02-rust-inventory-review`, preserving all
four Terra commits. This report's containing commit identifies the corrected
assessment revision; the Rust revision remains unchanged.

**Decision: ready for review and merge as a bounded implementation inventory.**
There is a newly demonstrated production command-registration defect, recorded
below, not corrected in this investigation. It prevents claiming working command
registration; it does not prevent merging accurate evidence about it. No production,
SDK, canonical catalogue or other-repository content was modified.

## Method and independence

This is a source-backed challenge and corrective review, not a blind second
assessor or human certification. The coordinating conversation contained Terra's
earlier work and the reviewer read Terra's report during orientation. Source
inspection then independently traced the raw declarations, actual methods,
generator output templates, callers and test assertions before accepting their
relationships. Neither the existing validator's expected tables nor a passing
test suite alone was treated as proof of implementation.

All 1,132 mappings were mechanically reconciled. **157 distinct identities were
investigated at declaration or functional-family level**: all 93 positive raw
representations (73 constants, 13 same-name declarations, seven differently named
declarations), 13 legacy identities, and the 51-name LPenHelper family. This
includes all prior positive functional classifications and all 78 prior test-linked
records. Semantic investigation used shared mechanisms and representative error
paths, not an exhaustive proof of each possible input or race. The remaining 975
worksheet/command raw identities were investigated as a source family and through
the generic dispatch boundary, not individually established as equivalent or absent.

The bounded source search included all 23 Rust files in the three core `src`
trees, their test trees, the minimal XLL, and C/Rust ABI checker. Direct review
covered all sys definition files and re-exports; the only sys generation-like
inclusion is README documentation. The proc macro emits registration metadata,
ABI thunks and conversions, not an unseen catalogue of xlf/xlc identifiers.
`thunk::RawXloper12` and `RawXchar` aliases were traced through generator uses.
The private backend resolves only MdCallBack12 for Excel invocation. There was
no alternate LPenHelper loader, type generator or UI-query implementation in
these inspected paths. This justifies a bounded family absence observation,
not a claim about every external crate or possible numeric call.

## Corrections and classification delta

86 JSONL records changed; **83 records have changed axis classifications**:
51 raw, 78 functional and five test-state changes (134 axis changes, with overlap).
The remaining three records change evidence or configuration only. No host state
changed. The small assessment-v2 extension retains version-1 fields, source IDs,
historical evidence and the four independent axes.

| Finding | Affected identities | Disposition |
| --- | --- | --- |
| Windows/x64 test incorrectly labeled default | CAPI-000001 XLOPER12 | Indented cfg attribute was missed by exact-line lookup. Correct configuration is `windows-x86_64`. Reproduced rejection of the correct label before fixing strip handling. |
| Return evidence could point to a comment/import | Demonstrated using CAPI-000049 xlbitDLLFree | Before fix, pointing to return_alloc.rs:5 passed validation. Both comment and import negatives now fail; supported materialization statements are matched precisely. Existing genuine citations were retained. |
| Build/dependency drift was not checked | All assessments dependent on compiled evidence | Added manifest, lockfile, build-script, Cargo configuration and toolchain paths. Five synthetic changed-path cases failed their rejection expectations before correction. No production file mutation was needed. |
| Descriptor presence understated implemented methods | 13 descriptor-linked records | Twelve now have typed-operation observations; xlfRegister is partial. Implementation anchors include generated context methods and cleanup, with research gates preserved. |
| Excel12v functionality was unassessed | CAPI-000004 | Private `SdkExcel12vBackend::excel12v_raw` forwards through MdCallBack12 with its parameter order. This is a generic backend, not a linked Excel12v export or public arbitrary safe-call API. |
| Command registration fails before dispatch | CAPI-000011 xlfRegister | Worksheet registration exists. Command builder supplies nine roots against descriptor minimum ten; mark `typed-registration-partial`. Owner #96 (EDNA-08), related #95 (shared registration). |
| Legacy applicability remained implicit | 13 XLREF/XLMREF/FP/XLOPER names, tags, pointer aliases, Excel4/Excel4v | Functional state now explicitly outside the Excel12 target, supported by sys/constants.rs:136 and sys/lib.rs support boundary. No demand to implement legacy ABI. |
| LPenHelper relationships remained unassessed | CAPI-000079, CAPI-000100–000101, CAPI-000104–000151 (51) | No raw declaration or operation observed in the inspected core: LPenHelper, two selectors, 44 edit/document/hit-test values, two aggregate typedefs and two tags. Separate from generic Excel12v selectors; applicability is a consumer decision. |
| Existing behavior tests were overlooked | CAPI-000003, 000006, 000039, 000041, 000046 | XLREF12, array, text, multi-reference and single-reference assertions were inspected and rerun. Replace ABI-only test state with the stronger named behavioral test reference; retain ABI evidence. Total test-linked IDs stays 78. |
| Generated aliases omitted from evidence | XCHAR and LPXLOPER12 | Added exact RawXchar/RawXloper12 alias references and checked generator use; no inherited blanket capability or host claim. |

## Positive findings challenged

**Numbers.** All 73 matching constants were checked against their actual sys
definitions and source catalogue. Sixty-six also have independent C/Rust numeric
comparisons rerun by the ABI tool. The other seven are xlGetInst, xlGetHwnd,
xlEnableXLMsgs, xlDisableXLMsgs, xlDefineBinaryName, xlGetBinaryName and
xlRunningOnCluster: OR with xlSpecial=16384 yields respectively 16391, 16392,
16394, 16395, 16396, 16397, 16402. Equality does not establish callability. The
XLTYPE/XLBIT/XLRET compatibility spellings and XlError discriminants reuse these
same declarations; they are not additional CAPI coverage achievements.

**Types and signatures.** The five scalar typedefs, four Excel12 aggregates,
three canonical pointer typedefs and one cluster callback typedef account for
13 same-name raw declarations. Four C struct tags resolve to differently named
Rust aggregates; Excel12, Excel12v and XLCallVer resolve to function-pointer
aliases. Struct bodies were inspected against source relationships and the
checker covers sizes, alignments and selected member offsets on Windows x64.
XLOPER12's named Rust substructures/unions represent the C anonymous members;
one-element tails do not confer allocation bounds. BOOL/RW/COL are i32, XCHAR
u16 and IDSHEET usize; source widths and target representation remain distinct
from ownership and function calling contracts. The C prototype static assertions
compare C types with C SDK declarations, not Rust function pointer compatibility.
No full ABI-equivalence status was manufactured, and no x86 claim is made.

**Calls and generated methods.** At the pinned Rust revision, excel_call.rs:400–505
contains a crate-private backend. CallCapability::call (:524) checks permission,
argument count and linkage, invokes the backend, and attaches result release
policy on success. The public descriptors alone are not callable public APIs.
The actual methods are recorded per row, with important examples:

| Mechanism | Source and scope |
| --- | --- |
| Sheet information | `sheet_context!`, excel_call.rs:1074–1134 emits active_sheet_id/current_sheet_name for worksheet, thread-safe and macro contexts. Active/front and current sheet semantics are deliberately distinct. |
| Caller | `caller_context!`, :1159–1178 emits caller for worksheet/macro only; returns ExcelOwnedValue. |
| Coercion | `coerce_context!`, :1183 onward uses a borrowed ExcelOwnedValue plus a four-value CoerceTarget. This is not arbitrary range construction, cell writing or all Excel coercion targets. |
| Cancellation | `abort_context!`, :1134 onward; cancellation_requested :587 handles optional retain/clear Boolean and checks result type. Not a calculation-complete query. |
| Lifecycle | get_module_name/register/unregister/delete_defined_name/register_event at :948–1070; runtime.rs drives registration, rollback and shutdown. |
| Research | Now/OnTime methods are cfg-gated by xlcontime-research; 132 feature-enabled unit tests pass, but the feature does not establish supported autonomous wake. |

All 1,001 selector relationships remain relationships only. Function modifiers
and numeric collisions are not aliases of functional support. xlSet has a raw
constant but no typed mutation operation in the inspected core; xlfGetDocument
has neither a named raw constant nor typed operation there. The generator does
not supply these missing methods. Use #93/#94 for consumer-required operations,
not one wrapper per historical number.

**Ownership, values and tests.** Borrowed decoder branches and eight error enum
conversions were inspected; the scalar/error test iterates all eight codes, rather
than incidentally mentioning them. Borrowed views preserve callback lifetimes;
owned value conversion copies strings/flat arrays, explicitly rejects references
without coercion and nested arrays. These are safety/representation boundaries,
not automatic defects. Flow and big-data decoding are explicitly rejected.
Raw flow/big-data fields still exist and have layout checks.

ExcelOwnedValue::attempt_release (excel_owned.rs:253) marks the obligation consumed
before calling xlFree, catches backend panics and does not retry ambiguous failure.
into_owned_value (:220) attempts release even when conversion fails, preserving
both errors. The conversion-and-release failure test at :475 and panic/once tests
provide local evidence, not a guarantee against every host leak. DLL returns take
a different path: return_alloc.rs:83 marks only the root DLLFree, hands off one
allocation; xl_auto_free12 :136 and reclaim_excel_return :168 reclaim that tree.
Arbitrary pointers or duplicate host frees violate the documented unsafe contract.
xlbitXLFree's transfer-token API does not expose a completed pass-through return
handoff; do not confuse that token with DLLFree completion. Reuse these owners
rather than adding a second free protocol.

The nine type checks, 66 constant checks and two C-only prototype associations
were reconciled with the checker. All original 39 named test associations were
challenged at their assertion/shared-fixture contract; five additional behavioral
associations now make 44. This still says only that the named contract was exercised.
The 26 rerun doctests include callback lifetime, non-Send/non-escaping values,
context restrictions and unsafe-return misuse examples; they are shared mechanism
evidence and were not automatically assigned to all CAPI numbers.

**Registration defect and scope.** The command builder in excel_call.rs:867–907
contains eight string entries plus inserted numeric macroType=2, hence nine roots.
register_command (:1023) sends them through XLF_REGISTER (:72), whose minimum is
ten; CallCapability::call rejects them at :538 before backend invocation. runtime.rs:388
uses that path. This deterministic source-level finding is checked by the narrow
argument-accounting regression; no Excel experiment is needed to prove this local
rejection. Runtime test fixtures at :766–805 contain worksheet/async functions and
no command list, explaining why the passing suite does not cover the defect.
The future fix must validate correct positional category/shortcut/help arguments,
not simply lower the minimum. HelpTopic is absent in metadata and emitted as an
empty registration slot. FP12 has raw layout and K% constant but neither enum nor
macro authoring variant; general XLOPER12 arrays are an existing alternative, not
direct FP12 registration. #95/#96 own these distinct contracts.

**Async, scheduling and diagnostics.** The async macro emits a void thunk with a
trailing handle, performs owned conversions before queuing, and uses the async
controller's generation/cancellation state. return_to_excel (async_udf.rs:753)
calls xlAsyncReturn and separately checks callback status and returned Boolean.
The rerun tests include late generations, cancellation and rejected return codes.
These do not prove live Excel completion. The dispatcher drains only through
callback contexts; enqueue does not wake Excel. Diagnostics preserve raw failures
through DiagnosticEvent rather than providing operation success. RTD, Ribbon,
Office COM and CLR are separate protocol boundaries; no CAPI relationship was
invented for them. Existing EDNA-01–14 table remains applicable, with command
registration now explicitly partial. The 57 records' 39 canonical and 44 candidate
EDNA links were reviewed against subsystem ownership; candidate links remain
conceptual relationships, never implementation/host acceptance.

## Coverage and remaining uncertainty

The [inventory report](rust-assessment-draft.md#per-kind-independent-evidence-counts)
contains reproducible denominators for every source kind and four axes. Raw
observations include bounded absence and legacy findings, not only implementations.

| Axis | Current disposition | Unassessed |
| --- | --- | ---: |
| Raw | 73 numeric matches; 20 nonnumeric/related declarations with ABI bounded; 13 legacy and 51 LPenHelper absence observations | 975 (573 worksheet/XLM, 402 commands) |
| Functional | 12 typed paths + 1 partial registration; 1 private backend; 45 other bounded mechanisms; 13 legacy exclusions; 51 LPenHelper absence observations | 1,009 |
| Automated tests | 44 named Rust test links, 33 further ABI-only links, 1 generated-thunk compile check | 1,054 |
| Native host | No current-revision per-CAPI result verified | 1,132 |

The 1,009 functional unknowns comprise 574 worksheet/XLM numbers, 402 commands,
10 special selectors, five flow codes, four function modifiers, four other
constants, four struct tags, three pointer typedefs, two prototypes and one
callback typedef. Raw/type observations do not automatically assign functional
status to aliases, tags or pointer declarations. No total implementation percentage
is meaningful across these categories.

Historical M8 smoke records Excel 16.0 build 20131 results, but no pinned artifact
hash/current-code attribution sufficient for these rows. M16 records registration
then workbook setup failure before async execution. Neither was promoted to current
host PASS. #118 owns scoped freshness and structured host evidence. GitHub #117
was accepted and closed in its [final acceptance record](https://github.com/evnekdev/excel-api/issues/117#issuecomment-6676074196);
the earlier committed provisional report predates that decision. Authenticated
machine completeness remains unavailable; this review does not change reconcile
exit 3. #120's rights question remains open and separate from factual source tracing.

## Verification performed in this review

Windows `x86_64-pc-windows-msvc`, rustc/cargo 1.99.1, Python 3.11.7. Rust source
is unchanged from the audited revision. No vendor data was copied or edited.

| Command | Result |
| --- | --- |
| `python -m unittest discover -s tools/excel-c-api-inventory -p "test*.py" -v` | PASS, 83 tests (75 retained, eight focused additions), exit 0 |
| `python tools/excel-c-api-inventory/validate.py check` | PASS, 1,132 catalogue records and reachable history, exit 0 |
| `python tools/excel-c-api-inventory/review_source.py --require-subfamilies` | PASS, zero differences, exit 0 |
| `python tools/excel-c-api-inventory/validate_rust_draft.py` | PASS, 1,132 mappings, exit 0 |
| `python tools/excel-c-api-inventory/summarize_rust_draft.py` | PASS; exact report table checked by test, exit 0 |
| `cargo test -p excel-api-sys --lib` | PASS, 7 tests |
| `cargo test -p excel-api --lib` | PASS, 130 tests |
| `cargo test -p excel-api --lib --features xlcontime-research` | PASS, 132 tests; overlaps default suite |
| `cargo test -p excel-api-macros --tests` | PASS, 7 unit tests + 1 trybuild harness with four fixtures |
| `cargo test -p excel-api --test macro_metadata` | PASS, 5 tests |
| `cargo test -p excel-api --doc` | PASS, 26 doctests |
| `cargo run --manifest-path tools/abi-check/Cargo.toml` | PASS, all 145 x64 checks |
| `cargo check -p excel-api-sys -p excel-api -p excel-api-macros --all-targets --all-features` | PASS |
| `cargo clippy -p excel-api-sys -p excel-api -p excel-api-macros --all-targets --all-features -- -D warnings` | PASS |
| `cargo fmt --all -- --check` | PASS |
| `git diff --check` | PASS |

Before fixes, the platform-label and comment-reference regressions failed, and
all five build-configuration mutation subcases failed their expected rejection.
After fixes they pass. New negatives reject wrong typed-operation anchors/version,
false legacy/LPenHelper classifications, attribution of the generic backend to an
arbitrary selector, mismatched generated aliases and unsupported command-gap
claims. The command defect itself remains present; passing the audit regression
means the inventory accurately reports it, not that command registration succeeds.

NOT RUN: broad workspace check/test/clippy (unrelated excel-com and research
tools), CI, native Excel automation, x86 checks and installer/release work. Focused
three-crate checks substitute for the broad ordinary-code baseline because this
change is assessment/Python/documentation only. No audit-tool success can
authenticate that an editable historical test-run string is true. External build
environment/global Cargo settings and full per-operation runtime legality remain
NOT VERIFIED. No additional Rust tests or production changes were introduced.

## Gap/reuse matrix and next work

| Priority / owner | Finding type | Reuse and bounded next acceptance |
| --- | --- | --- |
| 1: #96, related #95 | Demonstrated command registration defect | Correct CallArguments::command_registration position/count contract; exercise a command-bearing AddInDescriptor through the mock backend, assert actual argument values and successful dispatch/cleanup. Then separately verify in Excel. Do not just reduce the minimum. |
| 2: #85/#90 | ABI verification gap | Reuse tools/abi-check for x86 and function-pointer compatibility; keep x64 layout evidence and named type relationships. No legacy port solely for coverage. |
| 3: #93/#94 | Missing bounded safe operations | Reuse CallCapability, ExcelOwnedValue and context macros for consumer-required xlSet/document paths, references and recalculation. Specify allowed callback, arguments, result tags and cleanup before adding descriptors. Negative tests: illegal context, malformed result, backend failure, release-on-error. |
| 4: #95/#96 | Metadata/authoring gap | Extend existing registration and generator for required HelpTopic; prove slot order and generated metadata/ABI. FP12 exposure only after a demonstrated consumer need. |
| 5: #91/#92 | Deliberate boundaries plus verification gaps | Reuse borrowed decoder, FromExcel, ReturnPlan and distinct Excel/DLL owners. Preserve rejection of unsupported nested/reference conversions unless a reviewed design justifies them. Test malformed payloads, limits, panic and exactly-once cleanup. |
| 6: #99/#102/#101 | Runtime/architecture uncertainty | Retain native-async generation logic and cooperative queue; establish real host cancellation/unload and a legal wake mechanism separately. Do not substitute RTD or thread identity for callback authority. |
| Cross-cutting: #118 | Evidence/freshness gap | Record version/bitness/artifact/scenario-bound host outcomes and relevant dependency drift. No blanket pass across identifiers sharing a mechanism. |

The next smallest useful engineering task is the command-registration correction
under #96, coordinated with #95. This audit does not open duplicate implementation
issues or authorize its implementation. PyroApp requirements should choose the
next safe calls and help metadata, while its product acceptance remains in its
own repository.

## Closure recommendation

#122's bounded audit deliverables are satisfied: one row per canonical identity,
reference validation, explicit unknowns, source-backed corrections, reproducible
tests and gap/reuse owners. Recommend closure when the review branch is accepted;
there is no requirement to implement the remaining 975 raw/1,009 functional
unknowns to close an investigation. Leave #122 open for that disposition and
leave #90/#78 open. The unmerged review is not production parity, legal clearance,
authenticated completeness or host certification. No merge is performed here.
