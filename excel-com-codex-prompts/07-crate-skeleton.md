# Codex Prompt 07 — Create the `excel-com` Crate Skeleton

## Objective

Create the independently versioned `excel-com` crate with platform boundaries, module structure, documentation, CI compatibility, and compile-time safety markers.

Do not yet implement full `IDispatch::Invoke`.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
feature/excel-com-07-skeleton
```

## Required preparation

Work from the latest `master` containing approved Excel COM architecture. Read all Excel COM ADRs and architecture documents before editing.

## Workspace changes

Add:

```text
crates/excel-com/
examples/excel-com-quickstart/
```

Add them as workspace members. Do not add them to `default-members` unless the approved ADR requires it. Existing core crates must not depend on `excel-com`.

## Versioning and metadata

`excel-com` must be independently versioned. Do not use `version.workspace = true` unless explicitly approved.

Set package name, description, repository, docs URL, README, MIT/Apache-2.0 license, keywords, categories, Rust version, and Windows docs.rs target. Set `publish = false` for examples.

## Dependencies

Use approved versions of `windows` and `windows-core` only where required. Enable the minimum necessary Windows features.

## Module structure

Create a documented structure similar to:

```text
src/
  lib.rs
  apartment.rs
  dispatch.rs
  error.rs
  value.rs
  array.rs
  excel/
    mod.rs
    application.rs
    workbooks.rs
    workbook.rs
    worksheets.rs
    worksheet.rs
    range.rs
```

Avoid public placeholder APIs that will immediately be removed.

## Platform behavior

On Windows, expose intended entry points. On non-Windows, either compile with an explicit unsupported-platform error API or cleanly target-gate according to the ADR. Do not pretend Automation works outside Windows.

## Initial public surface

Implement only stable foundations:

```rust
pub type Result<T> = core::result::Result<T, Error>;
```

Create error variants for unsupported platform, COM initialization, changed apartment mode, and generic HRESULT-backed failure. Implement `Display` and `std::error::Error`.

Create the initial thread-bound apartment owner according to ADR-0036. Make thread affinity explicit through type structure and compile-time tests. Do not add `Send` or `Sync` unsafely.

## Documentation

Add crate README, crate-level Rustdoc, safety model, platform requirements, support boundary, created-versus-attached lifecycle preview, and an explicit statement that 0.1 is incomplete.

The quickstart may compile without meaningful Excel operations at this stage.

## Tests

Add tests for non-Windows boundary where CI permits, apartment initialization result mapping, compile-time `!Send`/`!Sync`, error formatting, and balanced cleanup through testable abstractions where feasible.

Do not require Excel for normal CI.

## Validation

Run:

```text
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo package --manifest-path crates/excel-com/Cargo.toml --allow-dirty
```

Report exact package dry-run behavior.

## Acceptance

- crate isolated from XLL core;
- clean platform gates;
- explicit apartment ownership;
- COM pointers cannot accidentally become cross-thread safe;
- accurate incomplete-state documentation;
- core workspace checks still pass.

Commit, push, and open a draft PR.

Do not merge. Stop after reporting the draft PR.
