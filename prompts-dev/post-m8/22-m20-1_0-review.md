# M20 — 1.0 readiness and stabilization review

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

## Branch
`release/m20-1.0-review`

## Purpose
Perform a release-blocking audit; do not add unrelated features.

## Required work
- Inventory public APIs; remove accidental exposure; classify stable/experimental surfaces; review naming, lifetimes, ownership, traits, and error extensibility; add API/semver diff tooling.
- Audit every unsafe block/export, ABI check, ownership state machine, panic boundary, context, threading rule, unload path, and deferred XLFree issue.
- Run Miri where applicable, fuzzing/sanitizers/static analysis, extended stress/soak, and supported Rust/Windows/Excel matrices.
- Complete public documentation, examples, migration/compatibility notes, support/security/contribution/release policies.
- Finalize crate split, features, dependencies, licenses, metadata, version synchronization, and crates.io dry runs.
- Produce a 1.0 report: blockers, accepted risks, deferred features, compatibility matrix, safety result, API freeze decision, and release checklist.

## Acceptance
Open a draft review PR. Do not tag, publish crates, or create a final release without explicit human approval.
