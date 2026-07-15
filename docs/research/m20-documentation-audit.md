# RC1 documentation audit

**Scope:** `excel-api-sys`, `excel-api`, and `excel-api-macros` only.

**Date:** 2026-07-15

This is a release-readiness audit, not an assertion that optional RTD, COM,
Ribbon, task-pane, `xlcOnTime`, or autonomous-notification work is supported.

## Status summary

| Crate/module | Crate overview | Public-item coverage | Safety coverage | Examples/guide | Feature docs | Finding |
| --- | --- | --- | --- | --- | --- | --- |
| `excel-api-sys` root, functions, errors, types | Complete | Complete except documented raw constants exception | Raw pointer/calling-convention obligations documented | Raw `no_run` example and guide boundary | No features | Ready for review |
| `excel-api-sys::constants` | Complete module rationale | Narrow `missing_docs` exception | ABI/ownership policy at crate/module level | Raw API only | No features | Important before stable: retain the checked-header audit for every constant change |
| `excel-api` root and prelude | Complete | Complete | Ownership/callback model documented | Complete landing example and guide links | `macros`; research feature excluded from docs.rs | Ready for review |
| `borrowed`, `value`, `convert`, `metadata` | Module/user-guide coverage complete | Complete | Callback-lifetime rules documented | Values, strings, arrays, references guide pages | N/A | Ready for review |
| `return_plan`, `excel_owned`, `return_alloc` | Architecture and guide coverage | Complete | Return/DLLFree/xlFree ownership documented | Ownership-safe return guidance | N/A | Ready for review |
| `registration`, `runtime`, `context`, `excel_call` | Guide and landing coverage | Complete | Context and lifecycle restrictions documented | Function/command/call guides | N/A | Ready for review |
| `async_udf`, `dispatcher` | Preview status documented | Complete | Generation/shutdown restrictions documented | Preview guide pages | N/A | Important before stable: retain preview labels and finish live validation |
| `excel-api-macros` | Complete | Public attributes complete | Generated ABI thunk boundary explained | Complete macro reference and trybuild links | No features | Ready for review |

## Evidence

The baseline inherited from PR #40 had **501** missing-documentation diagnostics
in `excel-api`; `excel-api-sys` and `excel-api-macros` were already clean except
for the deliberately narrow checked-header constants-module exception. The
final intended public API count is **zero**, confirmed with
`cargo rustdoc -p excel-api --lib --all-features -- -D missing-docs`.

Docs.rs now documents the normal `macros` authoring surface rather than enabling
all features: `xlcontime-research` remains repository-testable but is not part
of published documentation. Guide examples were added for strings, arrays,
optional values, controlled worksheet errors, preview async UDFs, and manual
cooperative dispatch.

`cargo package` successfully built and verified the `excel-api-sys` and
`excel-api-macros` archives, which were unpacked and inspected: both contain
their README, dual licences, source Rustdoc, and (for macros) compile-fail
fixtures, with no local machine paths or prototype files. `excel-api` package
creation and `cargo publish --dry-run` are blocked because `excel-api-macros`
version `0.1.0` is not yet present on crates.io; this is the existing publish
ordering blocker, not a packaging-content failure.

## Release checklist

### Must fix before stable 1.0

- Run `cargo doc --workspace --all-features --no-deps` with
  `RUSTDOCFLAGS="-D warnings"` after the public-item pass.
- Inspect generated docs for all three crates and re-run package-content review.

### Important before stable release

- Keep doctest/guide examples synchronized with macro `trybuild` fixtures.
- Add a local documentation check to CI if the release branch does not already
  run the warning-denied documentation command.
- Validate the user guide against a supported 64-bit Excel host once the
  existing live-validation gate is unblocked.

### Post-release improvement

- Add API-level examples to more value, registration, and call descriptor
  pages after the complete member-doc pass.
- Publish a versioned hosted guide once the crate versioning policy is final.

### Optional research

- RTD clean-host activation comparison, production RTD API, Ribbon/COM UI,
  custom task panes, and autonomous dispatcher notification remain outside the
  core release boundary.

## Classification

The user guide, crate landing pages, macro reference, package README content,
raw ABI safety documentation, and intended public API surface are ready for
review. No broad missing-documentation suppression was introduced.
