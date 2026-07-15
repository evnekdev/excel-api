# RC4 final publication review

Date: 2026-07-15
Audited commit: `4f0a0ffc28882e4c2ad52cc9622a67a609fa5f23`

## Release decision

**NOT READY**

The core build, test, ABI, package-list, and CI evidence is green. Publication
is blocked by contradictory release-facing documentation. RC4 is audit-only, so
this branch records the blockers rather than fixing them.

## Blockers

| Severity | Evidence | Resolution |
|---|---|---|
| Release blocker | [`INITIAL_IMPLEMENTATION_NOTES.md`](../../INITIAL_IMPLEMENTATION_NOTES.md) says the workspace has placeholder macros/lifecycle exports, unfinished registration, and incomplete ABI work. | Remove it from the release-facing surface or replace it with an accurate, explicitly historical record. |
| Release blocker | [`docs/release/core-1.0-readiness-audit.md`](../release/core-1.0-readiness-audit.md) still reports 541 missing-doc items and pending live validation despite the accepted RC baseline. | Reconcile this audit and its checklist with accepted RC1–RC3 evidence, retaining any real remaining debt. |

The following is an expected staged-publication constraint, not an attempted
publish: `excel-api` cannot complete a registry-aware package or dry run until
`excel-api-macros 0.1.0` is published and indexed on crates.io.

## Repository, API, and security review

- The audit branch starts at clean `origin/master`, with no local commits ahead
  of master and no merge conflicts.
- No active release-source `dbg!`, `println!`, TODO, FIXME, or XXX diagnostic
  code was found. Located `panic!` and `unreachable!` expressions are test
  injection/assertions or closed-enum exhaustiveness assertions.
- No tracked credential, private key, certificate, API key, or absolute user
  path was found. A byte-level UTF-8 check confirmed that the root README is
  correctly encoded; its apparent mojibake was terminal rendering only.
- `git fsck --no-reflogs` found only local dangling objects; no connectivity
  corruption affects the audited branch.
- RTD/COM/Ribbon prototypes are absent from core package file lists and core
  dependency graphs.

The publishable crates are `excel-api-sys`, `excel-api-macros`, and
`excel-api`, all version `0.1.0` with `rust-version = "1.85"`.

| Crate | Frozen release surface |
|---|---|
| `excel-api-sys` | Raw Excel 12 ABI aliases, constants, C-layout types, errors, and callback signatures. |
| `excel-api-macros` | `#[excel_function]` and `#[excel_command]` with closed signature validation. |
| `excel-api` | Values, conversions, returns, registration, contexts/calls, runtime, diagnostics, plus preview async and dispatch. |

The default `macros` feature is documented. `xlcontime-research` is experimental
and excluded from docs.rs. Async UDFs and the dispatcher remain preview; RTD,
COM/Ribbon, task panes, and autonomous notification remain excluded.

## Unsafe, ABI, ownership, and concurrency review

Workspace lints deny unsafe operations in unsafe functions, missing safety
documentation, and undocumented unsafe blocks. Unsafe code is confined to raw
ABI decoding, union/tag access, Excel12v binding, FFI entry points, return
allocation, and Excel-owned-result release. Its contracts cover pointer
validity, callback capability/lifetime, tag agreement, allocation origin, and
cleanup.

- Callback-borrowed values do not outlive callbacks.
- Owned semantic values and `ReturnPlan` contain no callback pointers.
- DLL-owned returns are released once through `xlAutoFree12`.
- Excel-owned results have one `xlFree` obligation.
- Async/dispatcher requests use generation-scoped idempotent retirement.

Deterministic tests cover cancellation, shutdown, panic containment, bounded
queues, stale generations, and no completion after unlink.

## Validation and CI

Passed:

```powershell
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo doc --workspace --all-features --no-deps
cargo run --manifest-path tools/abi-check/Cargo.toml
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/inspect-minimal-xll-exports.ps1
```

The ABI checker passed 145 checks against the checked `XLCALL.H` contract. The
release XLL inspection reported exactly 18 required production exports.

The latest observed Windows XLL workflow,
[`29416273643`](https://github.com/evnekdev/excel-api/actions/runs/29416273643),
passed MSRV, format, Clippy, all-feature tests, no-default core tests,
doctests, published-core Rustdoc, ABI, release build, export inspection, and
artifact upload. The tag-only package workflow was not run: RC4 creates no tag.

## Package and crates.io review

All manifests contain description, repository, documentation URL, dual license,
README, keywords, categories, and MSRV metadata. docs.rs targets
`x86_64-pc-windows-msvc`; `excel-api` documents `macros` without research.

| Crate | Package list | Package/dry run |
|---|---|---|
| `excel-api-sys` | README, licenses, manifest, and source only | Passed; dry-run upload was correctly aborted. |
| `excel-api-macros` | README, licenses, manifest, source, and trybuild fixtures only | Passed; dry-run upload was correctly aborted. |
| `excel-api` | README, licenses, manifest, source, and metadata test only | Registry-aware package and dry run blocked until `excel-api-macros 0.1.0` is indexed. |

No core package includes RTD/COM prototypes, scripts, research evidence, local
paths, or generated RC artifacts. The successful archives were inspected from
`target/package`.

## Version and publication order

`0.1.0` is the defensible release line while async UDFs and dispatcher remain
preview. A `1.0` commitment requires a distinct semver/support decision; RC4
makes no version change.

1. Publish `excel-api-sys` only after explicit approval.
2. Publish `excel-api-macros` only after explicit approval.
3. Wait for both crates to index, then rerun `excel-api` package and dry run.
4. Publish `excel-api` only after explicit approval.

## Draft release communications (not published)

**Changelog:** `0.1.0` introduces a native 64-bit Excel 12 XLL foundation:
audited ABI types, callback-borrowed and owned values, return ownership,
registration, macros, typed contexts/calls, lifecycle, diagnostics, and preview
async/dispatch support.

**GitHub Release / crates.io:** Native Rust building blocks for 64-bit Microsoft
Excel XLL add-ins. The release contains only `excel-api-sys`,
`excel-api-macros`, and `excel-api`. RTD, COM/Ribbon UI, task panes,
`xlcOnTime`, and autonomous wake are out of scope.

**Migration and known limitations:** There is no prior public version. 32-bit
Excel, Mac Excel, Excel Online, arbitrary worker-thread Excel calls, RTD,
COM/Ribbon, task panes, and autonomous notification are unsupported or deferred.
Async UDFs and cooperative dispatch remain preview.

## Next action

Resolve every blocker, rerun RC4 and all package checks, then make a new
publication decision. This branch must not publish crates, tag, or release.
