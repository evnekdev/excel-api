# M15 — Automated Excel integration and stress harness

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
Build a repeatable real-Excel suite for crashes, leaks, registration regressions, MTR bugs, and unload/reload failures.

## Branch
`feature/m15-excel-stress-harness`

## Required work
- Use PowerShell/COM or a dedicated harness to launch isolated Excel processes, load/unload/reload the XLL, invoke all functions/commands, and verify results.
- Cover scalars, UTF-16, arrays, missing/nil, Q/U, errors, panic fallbacks, and direct string inputs.
- Toggle MTR and run concurrent formulas; repeat recalc and workbook cycles.
- Capture exit status, event/crash evidence, diagnostics, workbook outputs, timing, handle and memory trends.
- Add hard timeouts, process cleanup, and artifacts. Separate quick PR smoke from long soak mode.
- Define Excel version/build matrix for self-hosted/manual runners without weakening global security settings.

## Acceptance
One command runs deterministic smoke against installed 64-bit Excel; soak mode performs thousands of recalculations and lifecycle cycles with diagnostic artifacts on failure.
