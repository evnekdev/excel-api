# M13 — Diagnostics, observability, and failure reporting

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
Add bounded production diagnostics without violating Drop, AutoFree, callback, threading, or loader-lock constraints.

## Branch
`feature/m13-diagnostics`

## Required work
- Define structured events for runtime, registration, Excel calls, thunk failures, return failures, release failures, and debug ownership invariants.
- Implement a bounded non-panicking sink with no allocation/formatting in guarded critical paths, no Excel calls, no recursion, and shutdown safety.
- Support optional debugger output, bounded memory ring, carefully scoped file sink, and user sink with reentrancy contract.
- Add correlation IDs without treating raw pointers as stable identities; avoid workbook/user data by default.
- Define severity and stable diagnostic codes separately from formatting.
- Test saturation, sink panic, concurrency, shutdown, recursion prevention, and exact Excel-code preservation.

## Acceptance
Diagnostics improve supportability but cannot change ownership behavior or become a crash source.
