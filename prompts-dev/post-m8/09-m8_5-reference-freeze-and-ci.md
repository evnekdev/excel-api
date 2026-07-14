# M8.5 — Freeze the handwritten reference implementation and add CI

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
Freeze the successful M8 manual XLL as the normative conformance target before procedural macros generate ABI glue.

## Branch
`chore/m8-reference-freeze-ci`

## Required work
1. Audit manual thunks, descriptors, type text, exports, runtime transitions, rollback, `xlFree`, DLLFree, and AutoFree. Correct stale roadmap wording, especially raw XLFree remaining deferred.
2. Add golden fixtures for every M8 function: procedure symbol, Excel name, exact type text, flags, argument metadata, return strategy, and error mapping. Add lifecycle/export fixtures and tests proving handwritten code matches them.
3. Add Windows x64 GitHub Actions for formatting, Clippy, tests, doctests, ABI checker, release XLL build, and export inspection. Upload the XLL and export report. Do not require live Excel in ordinary CI.
4. Update the M8 manual test record: Function Wizard, Add-in Manager, spill behavior, missing versus empty, Q versus U, and any direct string probes. Mark unperformed UI tests pending.
5. Add a short policy document defining the manual implementation as the M9 oracle: ABI signature, metadata, conversion, error mapping, ownership, panic boundary, and lifecycle.

## Non-goals
No macros, runtime redesign, expanded call catalogue, signing pipeline, or raw XLFree optimization.

## Acceptance
Golden fixtures cover all current procedures; CI is reproducible; documentation separates automated, interactive, and pending validation.
