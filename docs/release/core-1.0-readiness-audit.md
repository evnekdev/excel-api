# Core 1.0 Readiness Audit

Date: 2026-07-15  
Baseline: `7a6b22db0b9a33e3f067fb3649faec431b941812`

## Recommendation

**Ready after the listed blockers.** The audit found no demonstrated ABI,
ownership, or concurrency defect in the core implementation, but documentation,
live-validation, experimental-feature packaging, and staged-publication gates
remain before a stable 1.0 release can be represented truthfully.

## Audited scope

The release candidates are `excel-api-sys`, `excel-api`, and
`excel-api-macros`. The RTD prototype, Ribbon, general COM UI, custom task
panes, autonomous notification, and `xlcOnTime` research remain experimental or
deferred. They are not reasons to implement new code before core 1.0.

The RTD prototype remains an explicit, Windows-only, unpublished workspace
member so `--workspace --all-features` can exercise it. It is not a default
member, core dependency, Cargo publication candidate, or XLL package input.

## Findings corrected in M20

- Replaced Rust let chains that did not compile on the declared Rust 1.85 MSRV;
  the full workspace now checks with Rust 1.85.0.
- Added a dedicated MSRV CI job and a core rustdoc broken-link gate.
- Fixed two broken `AbortCheckMode` intra-doc links.
- Removed the safe public re-export of internal calculation-event helpers; the
  registered, panic-contained event procedures remain unchanged.
- Added `Display`/`Error` contracts for public async executor/completion errors.
- Replaced stringly `ExcelCallError` conversion/release messages with typed,
  source-preserving error variants.
- Added package metadata, package-specific READMEs, and Apache-2.0/MIT license
  texts to every publication candidate.
- Gated the macro-only integration test so `excel-api --no-default-features`
  remains a valid tested package configuration.
- Corrected stale M12, M14, M16, and M20 architecture/roadmap statuses.

## Public API audit

The three core source trees contain 553 public declaration lines by the audit
inventory. Modules, re-exports, generated-thunk support, context capabilities,
owned/borrowed values, errors, registration metadata, async execution, and
dispatcher tickets were reviewed for scope leakage.

One accidental surface was reduced: `calculation_canceled` and
`calculation_ended` are implementation helpers, not application APIs. The
hidden `xlcontime-research` feature remains a publication-boundary blocker:
although it is excluded from default builds and support promises, it still
appears in the publishable `excel-api` manifest and enables hidden research
items. Relocate that probe outside the published core crate or obtain an
explicit maintainer decision to ship it as an unstable, unsupported feature.

Rustdoc with `missing_docs` found 541 undocumented public items in `excel-api`,
175 in `excel-api-sys`, and none in `excel-api-macros`. Broken intra-doc links
are now zero. Completing and then enforcing public documentation is a 1.0
blocker; this audit does not hide the debt with crate-wide allowances.

## Unsafe, ABI, and ownership audit

Unsafe code occurs in 11 core source files (200 source lines containing the
term `unsafe`). Workspace lints deny unsafe operations in unsafe functions,
missing `# Safety` documentation, and undocumented unsafe blocks. Clippy with
all targets/features passes. The reviewed unsafe regions are confined to raw
callback decoding, union access, Excel12v calls, owned-result release, return
materialization, generated thunks, async handle copying, and FFI exports. No
concrete unsoundness was identified.

The native ABI checker passed all 145 checks, including sizes, alignments,
offsets, unions, constants, callback signatures, type text, and Excel limits.
The release XLL contains all 18 required exports.

Borrowed values remain callback-scoped and non-Send/non-Sync; owned semantic
values and return plans contain no callback pointers. Excel-owned auxiliary
storage has one `xlFree` obligation, XLL-owned returns have one DLLFree/
`xlAutoFree12` path, and dispatcher/async generations own queued work until one
idempotent retirement path releases capacity.

## Concurrency and panic audit

Deterministic tests cover executor shutdown races, dispatcher selection and
shutdown races, cancellation, panic retirement, stale generations, nested
drain suppression, bounded queues, and no completion after unlink. Production
locks are not held across Excel calls, executor joins, operation bodies, or
dispatcher waits. Callback wait rejection and bounded work queues remain
explicit.

This audit did not run ThreadSanitizer, Loom, Miri across Excel FFI, or a live
Excel lifecycle matrix. Those tools are not claimed as evidence.

## Error audit

Public operational error types now implement `Debug`, `Display`, and
`std::error::Error` where applicable. `ExcelError` and raw `XlError` remain
semantic worksheet error values rather than Rust operation failures. Excel C
API return codes, registration errors, runtime shutdown, dispatcher errors,
conversion errors, and release errors remain distinct typed paths.

## Documentation and examples

All 125 tracked Markdown files have valid local Markdown link targets and the scan
found no mojibake. Architecture status contradictions found by the audit were
corrected. The minimal XLL still demonstrates the recommended registration,
function, command, async, and cooperative-pump surfaces; experimental RTD and
`xlcOnTime` paths remain labelled and separately selected.

## Packaging and publication

All three package lists exclude the RTD prototype, control server, scripts,
workspace architecture documents, and COM dependencies. Each now includes its
own README and both license texts. `excel-api-sys` and `excel-api-macros`
publish dry-runs pass. `excel-api` cannot complete a crates.io-aware dry-run
until `excel-api-macros` 0.1.0 exists in the registry; the required staged order
is sys/macros first, then `excel-api`. Nothing was published.

## Remaining validation gaps

- Full real-Excel async UDF cancellation/recalculation/close/reopen/unload.
- Cooperative dispatcher pump, bounded batches, cancellation, and stale
  generation behavior in real Excel.
- Stress/soak/channel matrix on a host where plain `Workbooks.Add` succeeds.
- Complete public API documentation and an enforced missing-docs gate.
- A successful post-dependency-publication `excel-api` dry-run.
- Optional API-semver snapshot tooling and non-Windows core CI coverage.

The exact release decisions are tracked in
[`core-1.0-release-checklist.md`](core-1.0-release-checklist.md).
