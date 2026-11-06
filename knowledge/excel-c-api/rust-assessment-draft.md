# Source-to-Rust audit: preliminary search leads

Tracking: #122 (parent #90). Audited source revision for these search leads:
`172b306fae9b8496d5ad79b626a9fa0b1890cfd4`. The canonical declarations,
anchors and EDNA links remain in `symbols.jsonl`; rows here reference IDs only.

**This is an incomplete draft, not the requested implementation inventory.**
Every one of the 1,132 identities has four independent assessment states.
Seventy-three raw numeric constants have been evaluated against the source
catalogue and match; thirteen same-name raw declarations and seven related
declarations are present but their ABI equivalence remains unassessed. Thirteen
legacy names have no matching core declaration in the bounded search; that is
not proof that all possible alternatives are absent. The other 1,026 raw states
remain unassessed.
Thirteen function identifiers have a source-traced call descriptor; ten type
flags have a decoded branch, two have explicitly unsupported decode branches,
two ownership flags are masked during decoding, eight error codes map through
raw and safe error enums, two event IDs are registered for async functions,
two aggregate types have borrowed wrappers, `FP12` has no registration type
in the inspected enums, eleven return codes share a generic bit-flag wrapper,
and `xlAsyncReturn` has an async return operation. The other 1,080 functional
states remain unassessed. Thirty-nine records reference tests that passed in
this run, 1,093 test states are unassessed, and all 1,132 host
states are unassessed.
Descriptor presence does not establish complete operation support. `xlfNow`
and `xlcOnTime` descriptors are feature-gated research, not default support.
Across 23 Rust source files in the three core crate `src` trees, 86 rows have
exact-name textual public-definition candidates, all in the three
`excel-api-sys/src/{constants,types,functions}.rs` files; 1,046 do not. These are search
leads, not validated ABI equivalence or evidence that no other representation
exists. No percentages or missing-implementation conclusions follow from the
counts alone. Named
function-pointer equivalents and nonmatching names need manual investigation.
The broader `core_name_occurrences` lexical scan, including comments and
strings as well as code, found no exact spelling for 1,039 names (573 worksheet
numbers, 402 commands and 64 other names). A zero count does not exclude a
different Rust spelling, a numeric literal or a generic callback mechanism.
The descriptors are a limited functional trace. A test pass does not establish
complete behavior or native Excel acceptance. Seven additional source names
have explicitly related differently named Rust types (not ABI-equivalence claims).
Eighteen rows also link to return materialization separately from input
decoding; those links do not assert ownership correctness.
Nine rows carry links to size checks and 66 constants to numeric checks
exercised by the Windows x64 C/Rust ABI checker; the extent of the checks
differs by type and is not a complete C semantics or ownership proof.

Validate coverage, states and candidate line locations in the pinned Git
revision with `python tools/excel-c-api-inventory/validate_rust_draft.py`.
The validator independently evaluates the supported Rust numeric expressions and
compares the result with the pinned catalogue. Numeric equality alone says
nothing about call safety, C aggregate ABI, or runtime behavior. This narrow
validator checks descriptor names, associated constants and source locations;
it rejects other claims rather than validating full future schema,
functional references or implementation evidence. Further work is required
before this can be accepted as the requested audit.

### Draft mapping fields

Each JSONL row has exactly one `capi_id`, resolved against the pinned canonical
catalogue. `raw_state`, `functional_state`, `test_state` and `host_state` are
separate: no state is inferred from another. Empty `exact_name_candidates`
means the bounded textual public-definition search found none, not that no
equivalent exists. `core_name_occurrences` counts exact lexical tokens across
the core source files, including comments and strings; it is not a usage claim.
`related_rust` identifies a carefully reviewed different-name
declaration. `functional_evidence` gives one bounded code observation and its
configuration; `related_evidence` records additional return-materialization
uses. `test_evidence` records the named test, configuration and actual command;
`abi_evidence` identifies a checker statement, target and run command.
Host results remain unassessed. Paths and line numbers refer to the pinned Rust
revision above, and source anchors and value relationships are joined by CAPI ID
from the canonical catalogue, not duplicated as a second source registry.

The checker rejects missing, duplicate or unknown IDs; unexpected fields;
unsupported statuses; bad exact-name and descriptor search results; stale
paths/lines; mismatched numeric, enum, target, feature or test references; and
claims outside the small encoded evidence vocabulary. It also rejects changed
audited paths or a header not equal to the pinned blob after CRLF normalization.
It cannot prove that a documented test actually ran, enumerate generated Rust,
verify general Rust semantics, prove host execution or authenticate an assessor.

Raw statuses distinguish numeric equality, a same-name declaration, a related
different-name declaration and a legacy public declaration not observed in the
bounded search; declaration statuses leave ABI equivalence unassessed. Functional
statuses describe a named descriptor, a decoder branch, an explicitly rejected
decoder branch, an ownership mask, error conversion, event registration, borrowed
wrapper, async return operation, or the absence of FP12 in the inspected
registration enums. Return-code records can share a generic bit-flag wrapper;
that does not assert individual handling of each error. These are partial observations with source references, not
blanket safe-function claims. Test status records a named local pass or unknown;
host status remains unknown for the pinned source. A status never follows merely
from equality of unrelated numeric values.

### Raw source observations by category

Counts are exact-name/value observations, not a single implementation
percentage. `unassessed` includes possibly equivalent declarations under other
names and does not mean missing.

| Source category | Total | Numeric matches | Nonnumeric declarations, ABI unassessed | Legacy name not observed | Unassessed raw |
| --- | ---: | ---: | ---: | ---: | ---: |
| Type flags | 12 | 12 | 0 | 0 | 0 |
| Ownership flags | 2 | 2 | 0 | 0 | 0 |
| Error, flow, return, event and modifier codes | 30 | 30 | 0 | 0 | 0 |
| Special function identifiers | 20 | 18 | 0 | 0 | 2 |
| Worksheet/XLM identifiers | 579 | 6 | 0 | 0 | 573 |
| Macro commands | 403 | 1 | 0 | 0 | 402 |
| Other constants | 48 | 4 | 0 | 0 | 44 |
| Scalar typedefs | 5 | 0 | 5 | 0 | 0 |
| Aggregate typedefs | 10 | 0 | 4 | 4 | 2 |
| Struct tags | 10 | 0 | 4 related | 4 | 2 |
| Pointer typedefs | 6 | 0 | 3 | 3 | 0 |
| Function-pointer typedef | 1 | 0 | 1 | 0 | 0 |
| Function prototypes | 6 | 0 | 3 related | 2 | 1 |

The separate `related_rust` field points to seven different-name declarations,
including `Excel12Fn`/`Excel12vFn`/`XLCallVerFn` and four struct-tag counterparts.
Those observations do not yet establish full ABI or linkage equivalence.

The `descriptor_observation` field records a bounded scan of named descriptor
definitions in `excel_call.rs`: 5 of 579 worksheet/XLM identifiers, 1 of 403
commands and 7 of 20 special identifiers have named descriptors there. Of the
remaining 574, 402 and 13 respectively, absence from this one catalogue is not
proof of no generic raw dispatch, alternate API, host support or applicability.
The `xlcOnTime` and `xlfNow` descriptors require the research feature.

For the special-function family, the sole exact-name token in the scanned core
source is the raw declaration for `xlSet`, `xlStack`, `xlGetInst`, `xlGetHwnd`,
`xlEnableXLMsgs`, `xlDisableXLMsgs`, `xlDefineBinaryName`, `xlGetBinaryName`,
`xlRunningOnCluster` and `xlGetInstPtr`; none has a named descriptor in
`excel_call.rs`. This establishes a bounded declared-only observation, not that
an equivalent operation cannot exist under another name or through raw dispatch.

### Per-kind independent evidence counts

An observation includes a restricted or explicitly unsupported form; it is not a claim of complete implementation. A passing test is only the linked test, not full test coverage. Host assessed remains zero for this pinned revision. Reproduce this table with `python tools/excel-c-api-inventory/summarize_rust_draft.py`.

| Source kind | IDs | Raw observations / unknown | Function observations / unknown | Test passed / unknown | Host assessed / unknown |
| --- | ---: | ---: | ---: | ---: | ---: |
| error-code | 8 | 8 / 0 | 8 / 0 | 8 / 0 | 0 / 8 |
| event-id | 2 | 2 / 0 | 2 / 0 | 2 / 0 | 0 / 2 |
| flow-code | 5 | 5 / 0 | 0 / 5 | 0 / 5 | 0 / 5 |
| function-modifier | 4 | 4 / 0 | 0 / 4 | 0 / 4 | 0 / 4 |
| function-pointer-typedef | 1 | 1 / 0 | 0 / 1 | 0 / 1 | 0 / 1 |
| function-prototype | 6 | 5 / 1 | 0 / 6 | 0 / 6 | 0 / 6 |
| macro-command-id | 403 | 1 / 402 | 1 / 402 | 1 / 402 | 0 / 403 |
| other-constant | 48 | 4 / 44 | 0 / 48 | 0 / 48 | 0 / 48 |
| ownership-flag | 2 | 2 / 0 | 2 / 0 | 2 / 0 | 0 / 2 |
| pointer-typedef | 6 | 6 / 0 | 0 / 6 | 0 / 6 | 0 / 6 |
| return-code | 11 | 11 / 0 | 11 / 0 | 3 / 8 | 0 / 11 |
| scalar-typedef | 5 | 5 / 0 | 0 / 5 | 5 / 0 | 0 / 5 |
| special-function-id | 20 | 18 / 2 | 8 / 12 | 5 / 15 | 0 / 20 |
| struct | 10 | 8 / 2 | 3 / 7 | 3 / 7 | 0 / 10 |
| struct-tag | 10 | 8 / 2 | 0 / 10 | 0 / 10 | 0 / 10 |
| type-flag | 12 | 12 / 0 | 12 / 0 | 8 / 4 | 0 / 12 |
| worksheet-function-id | 579 | 6 / 573 | 5 / 574 | 2 / 577 | 0 / 579 |

## Source trace so far

- `crates/excel-api-sys/src/functions.rs` defines `Excel12Fn`, `Excel12vFn`,
  `XLCallVerFn` and `PXL_HPC_ASYNC_CALLBACK` function pointer types. These are
  not themselves proof of direct exported functions or full ABI compatibility.
- `crates/excel-api/src/excel_call.rs` defines context/permission-aware call
  descriptors for a bounded set including `xlGetName`, `xlfRegister`,
  `xlfUnregister`, `xlFree` and `xlAbort`. Constants alone do not imply a
  safely supported operation. `context.rs` separates callback capabilities.
- `crates/excel-api-sys/src/types.rs` contains `XLOPER12`-family raw types and
  layout tests; `tools/abi-check` contains C/Rust layout and signature checks.
  Test source presence does not establish a passing result on a target host.
- Registration, macros, ownership, conversions, async, lifecycle, tests,
  examples and host results require the remaining detailed source trace. Do
  not infer absence, completeness, callback legality or Excel acceptance from
  these preliminary findings.

The source inventory was independently reviewed by an agent; an authenticated
machine completeness attestation remains unavailable. SDK redistribution and
legacy header provenance are separately tracked by #120. Ribbon, RTD, Office
COM and CLR hosting are protocol boundaries and should not be assigned invented
CAPI identities. No native Excel automation was run.

## Executed Rust tests

On Windows x86-64 with rustc 1.99.1 / cargo 1.99.1, at the pinned Rust source
revision, `cargo test -p excel-api-sys --lib` passed 7/7 and
`cargo test -p excel-api --lib` passed 130/130. These are local unit tests,
including mocked callbacks; they are not CI or live Excel. The 39 linked rows
reference individual tests contained in those successful binaries. `cargo test
-p <crate> --lib -- --list` verified that every named test was present. The
XLOPER12 layout test is conditional on Windows x86-64. The validator checks
the named source and documented run, but cannot independently authenticate a
historical run or prove a host result.

`cargo test -p excel-api-macros --tests` passed seven macro unit tests and one
trybuild harness (one pass and three expected compile-fail fixtures).
`cargo test -p excel-api --test macro_metadata` passed five integration tests.
These exercises cover the generator and metadata pipeline, not every registration
scenario or Excel host acceptance.

`cargo run --manifest-path tools/abi-check/Cargo.toml` passed all 145 checks on
`x86_64-pc-windows-msvc`, including selected sizes, alignment, field offsets,
constants and C prototype checks. This is not x86 or live Excel verification.
The SDK header in the Windows checkout had SHA-256
`2e0a6907ff3694b9c5ea5b7891086f595d3250d7dce209616b481987705b66cf` due to
1,480 CRLF line endings; normalizing CRLF to LF produces byte-for-byte the
pinned Git blob SHA-256
`d69f860ca5f7be9779f69423276808845ebb2f9648e38aadc99908277a710c74`.
The validator checks this normalization and detects changes to the audited
Rust/checker paths. Recorded test execution is documentary evidence, not an
authenticated result artifact.

## Further concrete observations and gaps

- The borrowed decoder explicitly rejects `xltypeFlow` and `xltypeBigData`
  (`crates/excel-api/src/borrowed.rs:672` and `:721`); an existing test checks
  those rejections. The raw numbers exist, but safe decoding is unsupported.
- `ExcelCallDescriptor` and `CallCapability::call` bound context, argument
  count and result release (`crates/excel-api/src/excel_call.rs:41`, `:524`).
  `xlSet` has a raw number but no descriptor in that catalogue. This is a
  scoped source observation, not proof no other interface exists.
- `FunctionRegistration` (`crates/excel-api/src/registration.rs:104`) carries
  name, signature, category, description, argument names/help and flags; its
  structure does not contain a HelpTopic field. Macro generation begins at
  `crates/excel-api-macros/src/lib.rs:87` and builds registration metadata and
  thunks. End-to-end registration equivalence remains to be evaluated.
- `FP12` (CAPI-000002) exists as a raw structure and is covered by the x64
  size/offset checker. The inspected `ExcelArgumentType` and `ExcelReturnType`
  enums (`registration.rs:12` and `:48`) contain no FP12 type-code variant;
  no direct `FP12` reference was found in the higher-level crate or macro
  source. General `XLOPER12` arrays are a different, existing mechanism.
  Need for direct FP12 exposure is a consumer decision, not assumed parity.
- Async completion calls the generic backend with `xlAsyncReturn` at
  `crates/excel-api/src/async_udf.rs:753`; the dispatcher is cooperative and
  enqueue does not wake Excel (`crates/excel-api/src/dispatcher.rs:1`).
  Live cancellation, close/unload and scheduling acceptance remain separate.
- Historical `docs/release/core-1.0-readiness-audit.md` identifies baseline
  `7a6b22db0b9a33e3f067fb3649faec431b941812` and reports 145 ABI checks passed
  at that earlier revision. It is not a run or proof for the current baseline.
- `docs/manual-tests/m8-excel-smoke-test.md` at audit revision `172b306`
  (last changed at `e91856fe33e237847559bfe906feaf5d3d204daf`) records an earlier automated x64
  Excel 16.0 build 20131 smoke: registration, scalar/text/array/reference
  examples and recalculation, with interactive cases still pending. Its artifact
  revision/hash is not established by this draft, so no current per-CAPI host
  PASS is assigned. This historical result is not the absence of host evidence.
- `examples/minimal-xll/src/lib.rs` uses generated scalar, text, array,
  reference, optional, async and command examples. The checked-in M8 metadata
  snapshot describes selected generated registrations; it is a test fixture,
  not evidence that every example was loaded by Excel in this audit.
- The three core crate source trees contain no exact-name declaration for
  `XLREF` (CAPI-000022), `XLMREF` (CAPI-000025), `FP` (CAPI-000031),
  `XLOPER` (CAPI-000035), their legacy pointer typedefs, `Excel4`/`Excel4v`
  (CAPI-000077/000078), `LPenHelper` (CAPI-000079), `FMLAINFO` (CAPI-000148)
  or `MOUSEINFO` (CAPI-000150) in the bounded textual
  search; `excel-api-sys/src/constants.rs:136` also states legacy XLOPER and
  Excel4 are unsupported. No code-generation include was found in the core
  `src` trees; the only `include_str!` occurrences load crate README text. This
  identifies a legacy-ABI coverage gap relative to the source inventory, not a
  requirement to support legacy Excel for the current Excel 12 product target.
- The 44 miscellaneous constants without same-name public declarations are
  the four edit modes, five document types and 35 mouse hit-test codes. The
  remaining four miscellaneous constants—the three cluster return codes and
  `xlUDF`—do have matching raw constants. The first three families relate to
  the LPenHelper-side source declarations; this observation does not prove a
  need for those legacy UI APIs in the Excel 12 target.

## Subsystem connections for continued review

The following links identify code to inspect; they are not claims that every
subsystem or CAPI row has been implemented or host-tested.

| Subsystem | Observed code and boundary |
| --- | --- |
| EDNA-01 | `runtime.rs`: lifecycle state and registration; example exports in `examples/minimal-xll/src/lib.rs`. |
| EDNA-02 | `excel-api-sys/src/{types,constants,functions}.rs`: raw Excel 12 declarations; target and ABI evidence remain separate. |
| EDNA-03 | `borrowed.rs`, `value.rs`, `convert.rs`: callback-borrowed and owned values and conversions; flow/big-data decode explicitly rejected. |
| EDNA-04 | `excel_owned.rs`, `return_plan.rs`, `return_alloc.rs`: distinct Excel-owned and DLL-owned cleanup paths. |
| EDNA-05 | `borrowed.rs:483` and `excel_call.rs:1189`: reference views and coercion; full range mutation remains to be reviewed. |
| EDNA-06 | `excel_call.rs`: bounded descriptors and generic private backend, not one safe wrapper per function number. |
| EDNA-07 | `registration.rs:104`, macro crate `lib.rs:87`: metadata and generated function thunks; HelpTopic field absent in inspected descriptor. |
| EDNA-08 | `registration.rs:126` and macro crate `lib.rs:877`: command metadata and generator. |
| EDNA-09 | `convert.rs` and registration/macro machinery: conversion/authoring mechanisms; parity of advanced pipelines remains unassessed. |
| EDNA-10 | `context.rs`, `excel_call.rs`: callback capability types and permissions; legality per operation and live threaded behavior are separate. |
| EDNA-11 | `async_udf.rs:753`: async return through backend; live cancellation and lifecycle remain unassessed. |
| EDNA-12 | RTD is a distinct COM protocol; `examples/minimal-rtd-server` is an experimental example, not a CAPI symbol. |
| EDNA-13 | `dispatcher.rs`: callback-drained cooperative queue; no autonomous wake. |
| EDNA-14 | `excel_call.rs:116`, `runtime.rs`: event registration and lifecycle; complete event/host behavior remains unassessed. |

## Preliminary gap/reuse disposition

This section is bounded to inspected code; it is not a declaration-by-declaration
absence claim for the 1,046 unassessed raw identities.

| Area | Source observation | Needed evidence or follow-up |
| --- | --- | --- |
| Raw ABI | Excel 12 structures and function-pointer signatures are present; exact-name public constants are narrow. | Compare legacy forms, Windows x86/x64 layouts and external SDK type contracts; do not infer ABI from spelling. Reuse `tools/abi-check` and #85. |
| Conversions | Callback decoder explicitly rejects flow and big-data; scalar/text/array/reference branches exist. | Establish which unsupported tags consumers need; preserve ownership and lifetime guarantees, and test malformed and nested values. |
| Calls and references | Private generic callback backend supports narrow permission-aware descriptors; `xlCoerce` is exposed. `xlSet` (CAPI-000010) has a raw ID but no descriptor in this file; `xlfGetDocument` (CAPI-000334) has no same-name public declaration or descriptor in the scanned scope. | Trace alternate forms and #93/#94 before concluding missing APIs; add only justified operations with callback, argument and cleanup contracts. |
| Registration | Metadata/thunk generation and registration exist; inspected `FunctionRegistration` has no HelpTopic member. | Compare #95 and metadata call encoding; avoid duplicating generators or treating an absent field alone as host behavior. |
| Ownership and lifecycle | Callback borrowing, Excel-owned cleanup and DLL-owned returns have separate mechanisms. | Trace failure/rollback and exactly-once tests and independently verify host paths; no leak-free claim follows from the raw structs. |
| Threading and async | Capability-restricted calls, async return and cooperative dispatcher exist. | Test callback legality, cancellation, stale generations and unload in the appropriate host; autonomous wake is not provided by the cooperative queue. |
| Tests and host | 39 rows link local passing unit tests; no native host run is claimed here. Historical ABI report concerns an earlier revision. | Review remaining tests and pinned artifacts per operation; preserve not-run/unknown and use #118 for structured host evidence. |

Recommended dependency order once the remaining trace is complete: (1) establish
target ABI and source equivalence using the existing checker; (2) reuse the
generic backend, ownership and conversion paths; (3) address proven reference,
registration and safe-call gaps in their existing capability issues; (4) verify
errors, callback legality and lifecycle; (5) run separately authorized native
integration; (6) validate PyroApp-specific requirements. No per-function-number
wrapper or new implementation issue is implied by this search inventory.
