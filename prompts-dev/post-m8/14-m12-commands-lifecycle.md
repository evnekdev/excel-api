# M12 — Commands and lifecycle completeness

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
Complete lifecycle behavior and implement Excel commands separately from worksheet functions.

## Branch
`feature/m12-commands-lifecycle`

## Required work
- Implement `#[excel_command]` metadata/thunks with command return semantics and `MacroContext`.
- Add a minimal verified command example; commands must be unavailable from thread-safe UDF contexts.
- Complete AutoOpen/Close/Add/Remove/AddInManagerInfo; decide whether AutoRegister12 is justified.
- Test duplicate/unusual callback ordering, shutdown cancellation, partial unregister retry, and backend unlink ordering.
- Separate worksheet and command registration records; rollback partial registration in reverse order; hold no lock across Excel calls.
- Validate Add-in Manager metadata and real Excel load/remove/reload/command behavior.

## Acceptance
Commands and worksheet functions have distinct APIs/contexts; lifecycle stays consistent and leak-free under all tested sequences.
