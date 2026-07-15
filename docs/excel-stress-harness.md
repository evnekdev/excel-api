# M15 real-Excel stress harness

`scripts/excel-stress-harness.ps1` runs the minimal XLL in isolated Excel COM
processes. It is intentionally opt-in: GitHub-hosted Windows runners do not
have a supported desktop Excel installation, and no global Excel security
setting is changed.

Build the release XLL, then run the deterministic PR smoke command on a
64-bit Excel machine:

```powershell
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/excel-stress-harness.ps1 -Mode Smoke
```

The command creates a timestamped directory below `target/excel-stress/` with
one workbook, worker log, and JSON record per isolated lifecycle cycle, plus
`summary.json`. Each worker registers the XLL, evaluates every sample function
and the command, performs MTR recalculation, unregisters the XLL, closes the
workbook, and quits Excel. The parent process enforces a hard timeout and only
terminates Excel processes created after that cycle began.

For an extended manual/self-hosted run use:

```powershell
powershell -File scripts/excel-stress-harness.ps1 -Mode Soak
```

The default soak run performs 25 process/load/unload cycles and 2,000 full
recalculations in each cycle (50,000 rebuilds). Override counts only when
recording them with the artifacts. `-ValidateOnly` checks command construction
without starting Excel, which is suitable for non-Excel CI.

The probes cover scalar values, direct UTF-16 formula and cell strings,
value-only `Q` arrays, reference-preserving `U` input, missing and blank input,
Excel error values, a controlled conversion fallback, all registered sample
functions, and `RUST.PING.COMMAND`. The sample XLL deliberately has no public
function that panics; the production panic-to-`#VALUE!` invariant is exercised
by Rust thunk tests, while this harness records the real controlled error path.
It does not claim to inject a panic into a shipping add-in.

The minimal XLL currently has no COM-readable diagnostics snapshot. Each JSON
record therefore states that status explicitly and captures controlled errors,
worker exit codes, process/handle/memory snapshots, and Windows crash events as
the available live diagnostic evidence. A later public diagnostics export can
be added without changing the parent/worker containment model.

## Runner matrix

Run the smoke command at least once for every supported self-hosted Windows
runner image and record the generated `summary.json` in the PR:

| Runner | Required evidence |
|---|---|
| Windows x64, Excel Current Channel | version/build, MTR settings, artifacts |
| Windows x64, Excel Monthly Enterprise Channel | version/build, MTR settings, artifacts |
| Windows x64, Excel Semi-Annual Enterprise Channel (when supported) | version/build, MTR settings, artifacts |

Do not weaken macro, Protected View, or trusted-location policy for the test.
Use a dedicated test account and a locally built, unblocked XLL instead.
