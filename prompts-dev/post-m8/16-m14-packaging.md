# M14 — Reproducible XLL packaging, metadata, signing hooks, and releases

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
Create reproducible distributable Windows x64 XLL packages with externalized signing credentials.

## Branch
`feature/m14-packaging`

## Required work
- Deterministic artifact naming, version resources, and product/file/crate version consistency.
- Manifest with crate versions, git SHA, Rust/target/features, SDK provenance, hashes, and export report.
- Verify mandatory exports and package notices, install/uninstall, and smoke instructions.
- Add optional Authenticode hooks using externally supplied certificate/thumbprint and timestamp; never commit secrets; verify signatures.
- Build unsigned CI artifacts and optional protected signed release artifacts.
- Generate checksums and validate package contents.
- Decide SDK redistribution according to its license.
- Add tag workflow that validates clean semver/tag alignment but does not publish without explicit approval.

## Acceptance
A fresh documented Windows environment can reproduce the unsigned package; signed output is externally configured and verifiable.
