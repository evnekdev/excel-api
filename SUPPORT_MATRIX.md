# Core 1.0 Support Matrix

This table defines the release boundary. “Stable target for 1.0” identifies the
review scope; it does not turn pending live validation into a completed claim.

| Category | Scope | Release status |
|---|---|---|
| Stable target for 1.0 | Excel 12 ABI; callback-borrowed and owned semantic values; return planning/allocation; DLLFree/`xlAutoFree12`; Excel-owned result RAII; manual/macro registration; procedural macros/diagnostics; typed contexts and call catalogue; synchronous worksheet functions/commands; lifecycle; packaging/build infrastructure | Core 1.0 review scope |
| Implemented but live validation pending | Async UDF lifecycle; cooperative dispatcher pump; stress-harness smoke/soak/channel matrix | Automated coverage exists; real Excel matrix remains pending |
| Experimental | `examples/minimal-rtd-server`; RTD control server; `xlcOnTime` compatibility probe | Windows-only/unpublished research; no production support |
| Deferred | Ribbon UI; general COM automation; custom task panes; autonomous dispatcher notification; production RTD API | Optional post-1.0 roadmap |
| Unsupported | 32-bit Excel; arbitrary background-thread Excel C API calls; unverified RTD-to-M17 dispatch | Not supported by 1.0 |

The published 1.0 crate set is `excel-api-sys`, `excel-api`, and
`excel-api-macros`. Experimental artifacts have no dependency edge into those
packages and are never included in ordinary XLL packaging.

Issue #30 tracks only optional autonomous notification research. Issue #37
tracks optional RTD validation and production-design prerequisites. Neither is
a core 1.0 blocker.
