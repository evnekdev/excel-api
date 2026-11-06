# Source-to-Rust audit: preliminary search leads

Tracking: #122 (parent #90). Audited source revision for these search leads:
`172b306fae9b8496d5ad79b626a9fa0b1890cfd4`. The canonical declarations,
anchors and EDNA links remain in `symbols.jsonl`; rows here reference IDs only.

**This is an incomplete draft, not the requested implementation inventory.**
Every one of the 1,132 identities has four independent `unassessed` states.
Eighty-six rows have exact-name public-definition candidates from the three
`excel-api-sys/src/{constants,types,functions}.rs` files. These are search
leads, not validated ABI equivalence or evidence that no other representation
exists. No percentages or missing-implementation conclusions follow. Named
function-pointer equivalents and nonmatching names need manual investigation.
No safe API, test or host evidence has been assessed in the rows.

Validate coverage, states and candidate line locations in the pinned Git
revision with `python tools/excel-c-api-inventory/validate_rust_draft.py`.
This narrow validator rejects claims rather than validating full future schema,
functional references or implementation evidence. Further work is required
before this can be accepted as the requested audit.

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
CAPI identities. No native Excel, Rust or ABI test was run for this draft.
