# M18 — RTD / streaming data

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
`feature/m18-rtd-streaming`

## Architecture gate
First re-evaluate native XLL RTD, COM RTD, or another supported mechanism and create an ADR selecting crate boundaries, threading model, and deployment.

## Required work
- Define topic identity, subscription, update ownership, disconnect, workbook close, and shutdown.
- Use owned values across threads and marshal updates only through legal Excel/COM contexts.
- Bound/coalesce update queues; prevent callbacks after disconnect/unload.
- Keep RTD ownership separate from DLLFree/AutoFree assumptions.
- Integrate diagnostics, packaging, and stress tests, including reconnect and multiple subscribers.

## Acceptance
Multiple topics/subscribers update and disconnect safely; real Excel close/reconnect/unload stress passes.
