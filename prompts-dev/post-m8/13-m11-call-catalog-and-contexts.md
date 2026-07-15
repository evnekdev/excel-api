# M11 — Typed Excel call catalogue and richer execution contexts

## Universal project rules

- Work from the latest `origin/master` in `evnekdev/excel-api`.
- Confirm the worktree is clean before editing and create exactly one feature branch for this milestone.
- Read `README.md`, `ARCHITECTURE.md`, `ARCHITECTURE_INDEX.md`, `IMPLEMENTATION_ROADMAP.md`, all relevant architecture files, ADRs/checklists, and the existing implementation before editing.
- Treat checked-in `XLCALL.H` and official Microsoft documentation as authoritative. Excel-DNA and the book are secondary references.
- Keep raw ABI in `excel-api-sys`, policy/runtime in `excel-api`, macro expansion in `excel-api-macros`, and end-to-end examples in `examples/`.
- Never let a panic cross an FFI boundary. Never guess an Excel code, ownership rule, callback rule, or calling convention.
- Keep unsafe code centralized, minimal, and documented with validity, ownership, lifetime, and threading invariants.
- Do not hold project locks while calling Excel.
- Push the branch and open a **draft pull request**. Do not merge it. Stop after reporting the draft PR and wait for human review.

## Required baseline and completion validation

Run before editing and before opening the draft PR:

```text
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo run --manifest-path tools/abi-check/Cargo.toml
```

Run additional Windows, XLL, macro, and integration checks required below. Record exact commands and outcomes; never claim a check that was not run.

## Purpose
Expand the narrow lifecycle call layer into an auditable typed catalogue exposed only through legal callback capabilities.

## Branch
`feature/m11-call-catalog-contexts`

## Required work
- Define call metadata: function ID, legal contexts, MTR safety, argument rules, result-root need, ownership/release policy, and return-code interpretation.
- Keep arbitrary integer calls unsafe and internal.
- Add selected calls only: `xlCoerce`, verified sheet/caller helpers, abort/cancellation polling through `xlAbort`, with exact context and thread-safety rules verified from Microsoft documentation, and prerequisites justified by examples.
- Expose operations only through Worksheet, ThreadSafe, Lifecycle, or Macro contexts.
- Return `ExcelOwnedValue<'call>` for Excel-owned results and preserve exactly-once `xlFree`.
- Add explicit Q/value and U/reference public wrappers so macros cannot blur semantics.
- Name current-versus-active APIs precisely.
- Add mock-backend conformance tests for every call descriptor and ownership rule.

## Acceptance
Catalogue metadata is the source of truth; illegal context calls are unavailable or rejected before Excel; no owner escapes its callback capability.

## Calculation-state scope

Application calculation state (`Done`/`Calculating`/`Pending`) is not currently
exposed by the project's Excel C API catalogue. Microsoft documents
`Application.CalculationState` through VBA/COM, not through a verified
`Excel12`/`Excel12v` function ID. It is deferred unless an authoritative C API
contract is found. Do not add a placeholder calculation-state API or use an
XLM function unless its exact selector, argument contract, return type,
ownership, and legal callback contexts are authoritatively verified.
