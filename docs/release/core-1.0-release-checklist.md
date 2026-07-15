# Core 1.0 Release Checklist

## Must fix before 1.0

- [ ] Document every stable public item in `excel-api` and the raw ABI contract
  in `excel-api-sys`; enable a missing-docs CI gate after the backlog reaches
  zero.
- [ ] Resolve the `xlcontime-research` feature boundary so unsupported research
  is not accidentally presented as stable published-core API, or record an
  explicit maintainer exception.
- [ ] Complete supported-host live validation for async UDF lifecycle and the
  cooperative dispatcher pump.
- [ ] Complete the required stress/soak/channel evidence on a working Excel
  host, or reduce the stable support claim before release.
- [ ] Stage registry availability in dependency order (`excel-api-sys`,
  `excel-api-macros`, then `excel-api`) and obtain a passing `excel-api`
  `cargo publish --dry-run` without publishing the final release accidentally.
- [ ] Set the intended 1.0 versions and exact inter-crate requirements only
  after all prior gates pass.

## Recommended before 1.0

- [ ] Add an API-semver/public-surface snapshot and review it as the 1.0
  compatibility baseline.
- [ ] Add a changelog and release notes that distinguish stable, pending-live,
  experimental, deferred, and unsupported capabilities.
- [ ] Add non-Windows CI for the three portable core packages.
- [ ] Replace internal descriptor-result `expect` calls with controlled
  invariant errors where doing so does not distort the typed call design.
- [ ] Run additional concurrency modelling or sanitizers on a supported target
  where practical; state tool limitations precisely.

## Future improvements

- Coverage-guided malformed-value fuzzing.
- Legacy Excel/x86 support only through a separately designed compatibility
  effort.
- Additional verified Excel call-catalogue entries after authoritative
  contracts are available.

## Optional research

- RTD clean-host activation and production API design.
- Ribbon, custom task panes, and general COM integration.
- Autonomous dispatcher notification adapters and issue #30.
- `xlcOnTime` compatibility research.

None of the optional research items blocks core 1.0.
