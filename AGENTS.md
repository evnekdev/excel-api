# Repository engineering guidance

## Mandatory issue-first hierarchy

Follow [the shared issue-first and hierarchical-issue policy](https://github.com/evnekdev/engineering-architecture/blob/main/policies/agent-workflow.md): before any tracked file, documentation, test or configuration change, create or identify a **scoped issue in this repository**. For work spanning other repositories (including `pyroapprs`), establish and cross-link a separate issue in each affected repository **before** editing either.

For the [Excel-DNA capability roadmap](docs/architecture/excel-dna-capability-roadmap.md), keep the larger outcome traceable using GitHub **parent/sub-issues**: programme or milestone → subsystem/capability → independently testable implementation task (e.g. a specific Excel C API enum). Use as many levels as useful, not one issue per code line. Prefer native GitHub sub-issue links, including cross-repository links where appropriate; if tooling cannot create them, record reciprocal parent/required-child links and note that native association is outstanding. Use unrelated or blocked-by links for dependencies; a mere issue reference is **not** a child.

**Close leaf issues on verified completion; close parent issues only bottom-up when every required child is resolved recursively *and* the parent's own acceptance criteria have passed.** Closed-as-not-planned work does not satisfy an unmodified requirement. Do not prematurely close an initiative from one implementation PR or assume GitHub automatically closes parents. Record outstanding children, verification evidence and closure decisions. The architecture-planning issue [#76](https://github.com/evnekdev/excel-api/issues/76) is not itself the permanent implementation epic; use a separately scoped programme parent when implementation is organized.

## Tooling

Use the most focused available tool for the task: `rg` for text and symbol
searches; `rust-analyzer` when an available interface supports the required
semantic operation; `cargo check`, `cargo test`, `cargo nextest run`, `cargo
clippy`, and `cargo fmt` for Rust validation; `sg` for previewed structural
search and repeated syntax-aware rewrites; and `git diff` / `git diff --check`
for review. Use `cargo expand`, `cargo semver-checks`, `cargo public-api`, and
`cargo machete` when their respective macro, compatibility, public-surface, or
dependency questions are in scope.

## Refactoring

Before a nontrivial refactor, inspect `git status` and use `rg` to identify
definitions and references. Prefer semantic or AST-aware edits; preview `sg`
matches before applying repeated transformations. Do not modify generated files
unless the generator is updated, and regenerate them through that generator.
Preserve serialized names, FFI symbols, and public compatibility unless the
task explicitly authorizes a breaking change. Review the complete diff after
automated edits.

## Rust validation

For ordinary changes, run the following before completion and report anything
that could not be run:

```text
cargo fmt --all -- --check
cargo check --workspace --all-targets --all-features
cargo test --workspace --all-features
cargo clippy --workspace --all-targets --all-features -- -D warnings
git diff --check
```

Use the repository's relevant inventory, documentation, and integration
validation commands in addition to this baseline when the change affects those
areas.
