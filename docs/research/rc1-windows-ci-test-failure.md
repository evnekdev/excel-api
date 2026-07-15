# RC1 Windows CI doctest failure

## Failure evidence

- Workflow run: `29410171941`
- Failed job: `87335108519` (`Windows XLL` / `validate`)
- Failed step: `Tests`
- Command: `cargo test --workspace --all-features`
- Exit code: `1`

The workspace test command runs doctests. Two `excel-api` crate-landing-page
doctests failed: `crates/excel-api/src/lib.rs` lines 82 and 165. Both reported
`E0601: main function not found in crate rust_out`. All ordinary unit,
integration, trybuild, RTD-prototype, sys, and minimal-XLL tests in that job
passed; this was not an RTD, COM, concurrent-test, timeout, or assertion
failure.

## Root cause and correction

PR #41 feature-gated the macro item in the landing-page examples. The examples
also carried a feature-gated hidden `main` fallback. Rustdoc recognized the
lexical `main` while constructing its wrapper, but the doctest crate did not
enable the package feature cfg in the same way, so no actual entry point
remained. The result was deterministic under the workspace all-features run.

The fix retains the feature gate on the optional macro item and makes the
hidden empty `main` unconditional. It changes only doctest scaffolding and the
crate README; it does not alter macro behavior or any production API.

## Reproduction and regression evidence

Before the correction, local `cargo test --workspace --all-features` reproduced
the same two `E0601` failures. The focused regression command is:

```powershell
cargo test -p excel-api --doc
```

The complete workspace command is also repeated after the correction. The
ordinary test set uses its normal concurrency; no global serialization, sleep,
or timeout adjustment was added.

## Scope

Production runtime behavior, ownership, ABI, feature boundaries, and the
minimal XLL export surface are unchanged. The repair PR's Windows XLL workflow
run `29414822477` passed: MSRV and every validate step completed successfully,
including `Tests`, no-default-features tests, doctests, warning-denied published
Rustdoc, ABI checks, release XLL build, exact export inspection, and artifact
upload. The export inspection retained exactly 18 required production exports.
