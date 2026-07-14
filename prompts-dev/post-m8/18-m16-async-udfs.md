# M16 — Asynchronous worksheet functions

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

## Gate
Start only after M9-M15 are merged and stable. Re-read `ASYNC_ARCHITECTURE.md` and official async XLL documentation; create/update ADRs before coding if assumptions are incomplete.

## Branch
`feature/m16-async-udfs`

## Required work
- Verify async registration codes, handle ABI, completion API, cancellation, recalculation, workbook close, and unload behavior.
- Define an owned async handle and state machine.
- Generated async thunks must copy all callback inputs to owned Rust values before returning and retain no callback context or Excel-owned result.
- Schedule bounded work; completion marshals owned results through the official API and is at-most-once, panic-contained, and disabled after unlink.
- Provide runtime-neutral executor abstraction or optional adapters; do not force one runtime unnecessarily.
- Add cancellation, queue bounds, shutdown, race tests, macro diagnostics, and real Excel async smoke tests.

## Acceptance
No borrowed pointer/capability escapes; completion is at most once and never after shutdown; real Excel cancellation/recalc/unload tests pass.
