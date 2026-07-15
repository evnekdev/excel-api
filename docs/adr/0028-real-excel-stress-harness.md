# ADR-0028: Isolated real-Excel stress harness

## Decision

Use a PowerShell parent/worker harness. Every lifecycle cycle runs in a new
Excel COM process; the parent applies a hard worker timeout, records artifacts,
and cleans up only Excel processes started by that cycle.

## Consequences

This validates registration, unload/reload, MTR calculation, and observable
workbook outputs without requiring a desktop Excel installation on ordinary
GitHub-hosted CI. The harness records version/build, timing, process memory and
handle snapshots, Windows crash-event evidence where readable, worker logs, and
workbooks. A production XLL must not expose a test-only panic function, so real
panic injection is not claimed; Rust thunk tests continue to establish panic
containment and the harness tests a real controlled error fallback.
