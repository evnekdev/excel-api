# Source-to-Rust audit: preliminary search leads

Tracking: #122 (parent #90). Audited source revision for these search leads:
`172b306fae9b8496d5ad79b626a9fa0b1890cfd4`. The canonical declarations,
anchors and EDNA links remain in `symbols.jsonl`; rows here reference IDs only.

**This is an incomplete draft, not the requested implementation inventory.**
Every one of the 1,132 identities has four independent assessment states.
Seventy-three raw numeric constants have been evaluated against the source
catalogue and match; thirteen more raw declarations are present but their ABI
equivalence remains unassessed. The other 1,046 raw states remain unassessed.
Thirteen function identifiers have a source-traced call descriptor; ten type
flags have a decoded branch, two have explicitly unsupported decode branches,
two ownership flags are masked during decoding, eight error codes map through
raw and safe error enums, and `xlAsyncReturn` has an async return operation.
The other 1,096 functional
states remain unassessed. Thirty-two records reference tests that passed in
this run, 1,100 test states are unassessed, and all 1,132 host
states are unassessed.
Descriptor presence does not establish complete operation support. `xlfNow`
and `xlcOnTime` descriptors are feature-gated research, not default support.
Across 23 Rust source files in the three core crate `src` trees, 86 rows have
exact-name textual public-definition candidates, all in the three
`excel-api-sys/src/{constants,types,functions}.rs` files; 1,046 do not. These are search
leads, not validated ABI equivalence or evidence that no other representation
exists. No percentages or missing-implementation conclusions follow. Named
function-pointer equivalents and nonmatching names need manual investigation.
The descriptors are a limited functional trace. A test pass does not establish
complete behavior or native Excel acceptance. Seven additional source names
have explicitly related differently named Rust types (not ABI-equivalence claims).

Validate coverage, states and candidate line locations in the pinned Git
revision with `python tools/excel-c-api-inventory/validate_rust_draft.py`.
The validator independently evaluates the supported Rust numeric expressions and
compares the result with the pinned catalogue. Numeric equality alone says
nothing about call safety, C aggregate ABI, or runtime behavior. This narrow
validator checks descriptor names, associated constants and source locations;
it rejects other claims rather than validating full future schema,
functional references or implementation evidence. Further work is required
before this can be accepted as the requested audit.

### Raw source observations by category

Counts are exact-name/value observations, not a single implementation
percentage. `unassessed` includes possibly equivalent declarations under other
names and does not mean missing.

| Source category | Total | Numeric matches | Same-name nonnumeric declarations, ABI unassessed | Unassessed raw |
| --- | ---: | ---: | ---: | ---: |
| Type flags | 12 | 12 | 0 | 0 |
| Ownership flags | 2 | 2 | 0 | 0 |
| Error, flow, return, event and modifier codes | 30 | 30 | 0 | 0 |
| Special function identifiers | 20 | 18 | 0 | 2 |
| Worksheet/XLM identifiers | 579 | 6 | 0 | 573 |
| Macro commands | 403 | 1 | 0 | 402 |
| Other constants | 48 | 4 | 0 | 44 |
| Scalar typedefs | 5 | 0 | 5 | 0 |
| Aggregate typedefs | 10 | 0 | 4 | 6 |
| Struct tags | 10 | 0 | 0 | 10 |
| Pointer typedefs | 6 | 0 | 3 | 3 |
| Function-pointer typedef | 1 | 0 | 1 | 0 |
| Function prototypes | 6 | 0 | 0 | 6 |

The separate `related_rust` field points to seven different-name declarations,
including `Excel12Fn`/`Excel12vFn`/`XLCallVerFn` and four struct-tag counterparts.
Those observations do not yet establish full ABI or linkage equivalence.

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
CAPI identities. No native Excel or C-versus-Rust ABI checker was run.

## Executed Rust tests

On Windows x86-64 with rustc 1.99.1 / cargo 1.99.1, at the pinned Rust source
revision, `cargo test -p excel-api-sys --lib` passed 7/7 and
`cargo test -p excel-api --lib` passed 130/130. These are local unit tests,
including mocked callbacks; they are not CI or live Excel. The 32 linked rows
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
- Async completion calls the generic backend with `xlAsyncReturn` at
  `crates/excel-api/src/async_udf.rs:748`; the dispatcher is cooperative and
  enqueue does not wake Excel (`crates/excel-api/src/dispatcher.rs:1`).
  Live cancellation, close/unload and scheduling acceptance remain separate.
- Historical `docs/release/core-1.0-readiness-audit.md` identifies baseline
  `7a6b22db0b9a33e3f067fb3649faec431b941812` and reports 145 ABI checks passed
  at that earlier revision. It is not a run or proof for the current baseline.
- `docs/manual-tests/m8-excel-smoke-test.md` records an earlier automated x64
  Excel 16.0 build 20131 smoke: registration, scalar/text/array/reference
  examples and recalculation, with interactive cases still pending. Its artifact
  revision/hash is not established by this draft, so no current per-CAPI host
  PASS is assigned. This historical result is not the absence of host evidence.

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
| EDNA-11 | `async_udf.rs:748`: async return through backend; live cancellation and lifecycle remain unassessed. |
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
| Calls and references | Private generic callback backend supports narrow permission-aware descriptors; `xlCoerce` is exposed. `xlSet` has a raw ID but no descriptor in this file. | Trace all call sites and #93/#94 before concluding missing APIs; add only justified operations with callback, argument and cleanup contracts. |
| Registration | Metadata/thunk generation and registration exist; inspected `FunctionRegistration` has no HelpTopic member. | Compare #95 and metadata call encoding; avoid duplicating generators or treating an absent field alone as host behavior. |
| Ownership and lifecycle | Callback borrowing, Excel-owned cleanup and DLL-owned returns have separate mechanisms. | Trace failure/rollback and exactly-once tests and independently verify host paths; no leak-free claim follows from the raw structs. |
| Threading and async | Capability-restricted calls, async return and cooperative dispatcher exist. | Test callback legality, cancellation, stale generations and unload in the appropriate host; autonomous wake is not provided by the cooperative queue. |
| Tests and host | 32 rows link local passing unit tests; no native host run is claimed here. Historical ABI report concerns an earlier revision. | Review remaining tests and pinned artifacts per operation; preserve not-run/unknown and use #118 for structured host evidence. |

Recommended dependency order once the remaining trace is complete: (1) establish
target ABI and source equivalence using the existing checker; (2) reuse the
generic backend, ownership and conversion paths; (3) address proven reference,
registration and safe-call gaps in their existing capability issues; (4) verify
errors, callback legality and lifecycle; (5) run separately authorized native
integration; (6) validate PyroApp-specific requirements. No per-function-number
wrapper or new implementation issue is implied by this search inventory.
