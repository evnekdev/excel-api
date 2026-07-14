# M9A — Procedural macro metadata and typed signature generation

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
Implement `#[excel_function]` as a metadata generator while leaving the ordinary Rust function and manual ABI thunk unchanged.

## Branch
`feature/m9a-macro-metadata`

## Required work
- Read `PROC_MACRO_ARCHITECTURE.md`, registration architecture, M8 fixtures, macro crate, and all manual M8 procedures.
- Define and document a stable attribute syntax for Excel name, category, description, flags, and argument help.
- Build a closed Rust-type mapping to the existing typed registration signature.
- Initially support scalar inputs/outputs; explicit Q/value and U/reference wrappers; explicit counted/NUL UTF-16 input wrappers; optional values; context injection; owned text/value/result outputs; and supported `Result<T,E>`.
- Reject generics, methods, async, impl Trait, borrowed returns, variadics, destructuring, ambiguous Q/U, and unsupported direct dynamic-string returns.
- Generate deterministic hidden metadata: `FunctionRegistration`, `FunctionSignature`, names/help/flags, Excel name, and future thunk symbol association.
- Do not generate or export an FFI thunk yet.
- Compare generated metadata and type text exactly to M8 golden fixtures. Add expansion snapshots and compile-pass examples.
- Detect symbol collisions deterministically or generate collision-resistant symbols.

## Acceptance
All M8 functions can be expressed with macro metadata and exactly match fixtures; no thunk or unsafe callback code is generated.
