# Apply this package

This is an **overlay**, not a replacement repository.

Keep the complete existing repository, including `.git`, crates, examples,
architecture files, and the Microsoft SDK. Extract this ZIP into the repository
root and choose **replace/overwrite on conflict**.

Then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\apply-prompt01-readiness.ps1
```

Review and validate:

```powershell
git status
git diff
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo run --manifest-path tools/abi-check/Cargo.toml
```

The cleanup script removes only transfer-only duplicates:

- `PACKAGE_README.md`
- `MANIFEST.md`
- root-level `BOOK_REVISION_NOTES.md`

The research notes remain under `docs/research/`.
