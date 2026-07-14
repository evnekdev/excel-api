# M9B — Generate worksheet-function ABI thunks

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
Generate the exact M8 thunk pipeline from M9A metadata, delegating all policy to audited runtime helpers.

## Branch
`feature/m9b-generated-thunks`

## Required work
- Generate exact exported ABI signatures implied by typed registration metadata.
- Decode inputs through existing borrowed parsers; inject only legal contexts; call the ordinary Rust function; convert, plan, materialize, and DLLFree-handoff outputs.
- Catch panics, use controlled Excel error fallbacks, and create one fresh dynamic root per call.
- Support scalar signatures, Q/value-only, U/reference-preserving, missing-aware inputs, dynamic XLOPER12 text/array returns, and validated direct UTF-16 inputs.
- Do not generate dynamic direct simple-string returns.
- Ensure registration type text and thunk signature share one internal type model.
- Generate deterministic unique x64 export names and verify exports.
- Compare generated and handwritten M8 behavior: tags, values, errors, ownership bits, Q/U behavior, symbols, and panic boundaries.
- Keep handwritten thunks until parity is proven, then migrate the minimal example while retaining fixtures.

## Acceptance
The minimal XLL builds and runs with generated thunks; automated parity passes; real Excel smoke is rerun or explicitly pending.
