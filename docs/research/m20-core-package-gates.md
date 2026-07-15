# M20 Core Package Gate Record

Date: 2026-07-15

The initial publish candidate set is `excel-api-sys`, `excel-api-macros`, and
`excel-api`. No prototype, example, RTD control server, or research helper is a
publish candidate.

## Package contents

`cargo package --list` passed for all three packages. The lists contain only
each package's own manifest/lock/source/tests and Cargo metadata; none contains
`examples/minimal-rtd-server`, RTD scripts, control-server source, or a COM
dependency.

## Dry-run results

- `cargo publish -p excel-api-sys --dry-run`: passed (9 files packaged and
  verified; upload aborted because this was a dry run).
- `cargo publish -p excel-api-macros --dry-run`: passed (13 files packaged and
  verified; upload aborted because this was a dry run).
- `cargo publish -p excel-api --dry-run`: blocked before verification because
  crates.io does not contain the unpublished optional dependency
  `excel-api-macros` version `0.1.0`.

This is an M20 release-process blocker, not an RTD/COM/Ribbon blocker. A later
release rehearsal must validate publish ordering and version availability:
publish/verify `excel-api-sys` and `excel-api-macros` first, then rerun the
`excel-api` dry run. Nothing was published during this scope milestone.
