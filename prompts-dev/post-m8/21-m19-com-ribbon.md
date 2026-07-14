# M19 — Optional COM and Ribbon integration

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
`feature/m19-com-ribbon`

## Architecture gate
Create an ADR selecting optional crate boundaries, COM apartment model, registration/deployment, and Ribbon resource strategy. Core crates must remain usable without COM/Ribbon.

## Required work
- Separate COM interfaces/class factory, apartment/threading, registration, Ribbon XML/callbacks, and XLL integration.
- No Excel calls from DllMain; explicit COM init/shutdown.
- Ribbon callbacks use legal main-thread capability; callback-scoped owners do not retain COM objects.
- Support optional packaging/signing and preferably per-user or registration-free deployment where feasible.
- Add unit/mocks and real Excel UI load/unload tests.

## Acceptance
Optional UI loads/unloads cleanly without weakening core XLL safety or making COM dependencies mandatory.
