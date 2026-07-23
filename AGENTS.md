# Repository engineering guidance

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
