# Architecture Index

For user-facing crate documentation and examples, start with the
[user guide](docs/guide/README.md). This index remains a maintainer-oriented
architecture map; RTD, COM/Ribbon UI, task panes, `xlcOnTime`, and autonomous
notification are optional or experimental, not core 1.0 prerequisites.

| Area | Document | Status | Milestones |
|---|---|---|---|
| Overall | `ARCHITECTURE.md` | Core architecture implemented; release audit active | M0-M20 |
| ABI | `ABI_ARCHITECTURE.md` | Implemented | M1 |
| Memory | `MEMORY_OWNERSHIP_ARCHITECTURE.md` | M2-M8 implemented (raw XLFree return transfer intentionally deferred) | M1-M8 |
| Strings | `STRING_ARCHITECTURE.md` | Partial (DLLFree returns implemented) | M1-M8 |
| Conversion | `TYPE_CONVERSION_ARCHITECTURE.md` | M3-M5 implemented | M2-M5 |
| Arrays/references | `ARRAY_REFERENCE_ARCHITECTURE.md` | Partial (DLLFree return multis implemented) | M2-M8 |
| Threading | `THREADING_ARCHITECTURE.md` | Core through M17 implemented; live async/dispatcher validation pending | M1-M17 |
| Contexts | `RUNTIME_CONTEXT_ARCHITECTURE.md` | M8 callback capabilities implemented | M4-M12 |
| C API calls | `EXCEL_CALL_ARCHITECTURE.md` | M8 narrow production catalogue implemented | M7-M12 |
| Lifecycle | `CALLBACK_LIFECYCLE_ARCHITECTURE.md` | M8 implemented; automated live Excel passed | M6-M12 |
| Registration | `REGISTRATION_ARCHITECTURE.md` | M8 manual oracle, M9 generation, and M10 diagnostic conformance implemented | M8-M10 |
| Errors | `ERROR_ARCHITECTURE.md` | Typed core errors implemented; M20 consistency audit active | M2-M20 |
| Testing | `TESTING_ARCHITECTURE.md` | Partial (automated real-Excel smoke/soak harness implemented) | All |
| Macros | `PROC_MACRO_ARCHITECTURE.md` | M9 generation and M10 compile-time conformance implemented | M9-M10 |
| Packaging | `PACKAGING_ARCHITECTURE.md` | Reproducible M14 XLL packaging implemented; publication rehearsal active | M14-M20 |
| Async | `ASYNC_ARCHITECTURE.md` | M16 implementation and race hardening complete; live validation pending | M16 |
| Dispatcher | `MAIN_THREAD_DISPATCH_ARCHITECTURE.md` | Cooperative M17 implemented; live validation and autonomous wake pending | M17 |
| RTD | `RTD_STREAMING_ARCHITECTURE.md` | Experimental Windows-only prototype; optional post-1.0 | Optional E1-E2 |
| COM/Ribbon | `COM_ARCHITECTURE.md`, `RIBBON_UI_ARCHITECTURE.md` | Deferred optional integrations | Optional E3-E4 |
| Support/release | `SUPPORT_MATRIX.md`, `OPTIONAL_INTEGRATIONS_ROADMAP.md`, `docs/release/core-1.0-release-checklist.md` | Core 1.0 boundary defined; release audit active | M20 / E1-E5 |

## Dependency order

```text
ABI
 -> memory/string
 -> conversion/arrays/references
 -> threading/contexts/calls
 -> lifecycle/registration
 -> macros
 -> async/cooperative dispatch
 -> M20 core stabilization

Optional after core 1.0: RTD/COM/Ribbon/notification adapters
```

## Release boundary

Core 1.0 excludes optional RTD, COM/Ribbon, custom task panes, autonomous
notification, and `xlcOnTime` research. See ADR-0033 and
`SUPPORT_MATRIX.md`; those integrations remain separate post-1.0 decisions.

## Freeze policy

A design becomes `Implemented` only after code and tests exist. It becomes
`Stable` only after real Excel integration and API review.
