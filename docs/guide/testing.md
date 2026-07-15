# Testing

```powershell
cargo test --workspace --all-features
cargo test --workspace --doc
```

Run deterministic repository checks first:

```powershell
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo run --manifest-path tools/abi-check/Cargo.toml
```

Then build and inspect the XLL. Real Excel testing remains essential for
loading, registration, calculation, callback legality, teardown, and process
isolation. The stress harness is evidence for trends and cleanup behavior, not
proof that an add-in cannot leak or crash.
