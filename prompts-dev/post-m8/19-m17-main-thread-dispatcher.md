# M17 — Main-thread dispatcher

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
`feature/m17-main-thread-dispatcher`

## Purpose
Allow background work to request main-thread-only Excel operations without calling Excel from worker threads.

## Required work
- Select an officially supported wake-up mechanism and document it in an ADR.
- Define owned request/response messages, bounded queueing, cancellation, timeout, shutdown, and backpressure.
- Main-thread drain creates a legal capability and executes only whitelisted operations.
- Never hold queue/runtime locks while calling Excel; prevent reentrancy deadlocks and recursive dispatch.
- Unload must cancel/drain before unlink and leave no stranded waiters.
- Integrate diagnostics and stress tests; feature-gate if appropriate.

## Acceptance
Worker threads never call Excel directly; requests execute under a main-thread capability; shutdown has no post-unlink calls or stranded requests.
