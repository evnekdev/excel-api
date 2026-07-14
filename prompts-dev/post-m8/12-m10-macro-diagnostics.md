# M10 — Compile-time diagnostics and macro conformance

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
Make unsupported or inconsistent signatures fail at compile time with actionable diagnostics.

## Branch
`feature/m10-macro-diagnostics`

## Required work
Use `trybuild` or equivalent. Add compile-fail tests for unsupported inputs/outputs, borrowed returns, generics, methods, receivers, async, variadics, ambiguous Q/U, direct dynamic string ABI, incompatible flags, forbidden contexts, unjustified cluster safety, duplicate names/symbols, metadata mismatch, invalid attributes, unsupported Result errors, and invalid string parameter forms.

Diagnostics must point to useful spans, explain the rule, and suggest a supported replacement. Add compile-pass tests for every supported signature and expansion snapshots. Prove metadata and thunk signatures cannot diverge. Document generated-symbol semver behavior.

## Acceptance
Common mistakes fail before Excel runs; no runtime-only check remains where compile-time validation is possible.
