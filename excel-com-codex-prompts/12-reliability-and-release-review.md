# Codex Prompt 12 — Reliability, Retry Policy, Guards, and `0.1` Review

## Objective

Harden `excel-com` for practical interactive desktop automation and assess readiness for an initial `0.1.0` release.

Do not expand into events, charts, pivots, or generated interfaces unless required to fix an architectural defect.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
feature/excel-com-12-reliability-review
```

## Required implementation

### Busy-call handling

Implement the approved bounded retry mechanism for rejected COM calls, including relevant HRESULT cases such as call rejected and server retry later.

If the architecture selected `IMessageFilter`, implement registration, restoration of the previous filter, retry decisions, bounded timeout, cancellation, apartment confinement, and panic containment.

Do not retry permanent errors or use unbounded sleeps.

### Retry policy

Expose a clear policy such as:

```rust
RetryPolicy::none()
RetryPolicy::bounded(...)
```

Document defaults. The default must not hang indefinitely.

### Application-state guards

Implement RAII guards for selected settings where safe:

- `DisplayAlerts`;
- `ScreenUpdating`;
- `EnableEvents`;
- calculation mode only if fully supported.

Guards must capture the original value, set the temporary value, restore it, define restoration failure behavior, avoid panicking in `Drop`, and support explicit restoration where callers need the error.

### Cleanup workflows

Add helpers or documented patterns for closing workbooks after failure, explicitly quitting created instances, preserving attached instances, avoiding hidden child references, and predictable release order.

Do not add force-kill behavior to the normal API.

### Diagnostics

Add structured context for activation, member lookup, invocation, rejected calls, conversion, shape mismatch, save/close/quit, cleanup failures, and missing/unsupported members.

Never log raw interface pointer values.

### Stress harness

Create an opt-in real-Excel stress harness covering:

- repeated create/add/write/save/close/quit;
- repeated attach/read/release;
- rectangular values and formulas;
- mid-workflow failures;
- busy/retry behavior where reproducible;
- unsaved prompts;
- hidden and visible instances;
- process cleanup.

Record Excel version/build, Windows version, Rust target, test count, failures, remaining Excel processes, and manual cleanup required.

### Documentation

Complete crate README, crate-level docs, object-model guide, lifecycle guide, threading guide, error guide, range-value guide, and examples.

Clearly state Windows desktop Excel and interactive-session requirements; unsupported server/service automation; `!Send`/`!Sync`; created versus attached behavior; explicit quit policy; raw dispatch escape hatch; and incomplete object-model coverage.

### Public API review

Audit accidental public items, unsafe exposure, error consistency, naming, semver hazards, feature flags, documentation completeness, thread-safety markers, dependency weight, package contents, and MSRV compliance.

Add API snapshot tooling if practical and proportionate.

## Release boundary

Assess 0.1 support for application creation/attachment, workbook operations, worksheet operations, scalar and rectangular range values, formulas, macro execution, raw dispatch, structured errors, retry, and cleanup.

Explicitly defer events, Ribbon, RTD, task panes, full generated interfaces, broad chart/pivot wrappers, macOS, server automation, and automatic `excel-api` integration.

## Validation

Run:

```text
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo doc --workspace --all-features --no-deps
cargo package --manifest-path crates/excel-com/Cargo.toml
```

Run all applicable Windows and real-Excel validation scripts. Never claim tests that were not executed.

## Required release audit

Create:

```text
docs/release/excel-com-0.1-readiness-audit.md
docs/release/excel-com-0.1-release-checklist.md
```

Classify findings as blocker, high priority, follow-up, or explicitly deferred.

Recommend one of:

- ready for `0.1.0`;
- ready after listed blockers;
- not ready, with a specific next milestone.

## Completion

Commit, push, and open a draft PR. The PR must include implementation summary, exact validation evidence, real-Excel evidence, remaining risks, and release recommendation.

Do not publish, tag, merge, or create a release. Stop after reporting the draft PR.
