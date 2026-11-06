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
| Excel C API-first inventory | `docs/architecture/c-api-inventory-layer.md`, `knowledge/excel-c-api/README.md`, `knowledge/excel-c-api/schema.md`, `knowledge/excel-c-api/symbols.jsonl`, `docs/evidence/excel-c-api-source-census.md` | **1,132 source records including 15 seed identities; independent agent technical review recorded; authenticated completeness and Rust/Excel acceptance pending** | #116 / EDNA-02 #90 |
| Excel-DNA subsystem-to-crate map | `docs/architecture/excel-dna-subsystems-and-crates.md` | Proposed ownership for 25 identified EDNA capability subsystems; 3 core crates retained, optional Ribbon/RTD and research-only host wake; no code split performed | #78 / #79 |
| Excel-DNA-equivalent Rust infrastructure | `docs/architecture/excel-dna-capability-roadmap.md` | Proposed long-term C-API-first capability roadmap; PyroApp full parity is first acceptance consumer; implementation/verification independently tracked | P0 / B1-B7 |
| Native XLL host integration | `docs/architecture/native-xll-hosting-roadmap.md` | Architecture baseline complete; optional host/Ribbon implementation not started | H0-H8 |
| RTD | `RTD_STREAMING_ARCHITECTURE.md` | Experimental Windows-only prototype; optional post-1.0 | Optional E1-E2 / H7 if required |
| COM/Ribbon | `COM_ARCHITECTURE.md`, `RIBBON_UI_ARCHITECTURE.md` | Deferred optional integrations; concrete host/Ribbon path is now defined by the H roadmap | Optional E3-E5 / H2-H5 |
| Excel COM extraction | `docs/architecture/com-automation-roadmap.md` | Architecture decision and staged generic COM extraction roadmap | C0-C9 |
| Excel COM knowledge base | `knowledge/excel-object-model/README.md`, `docs/research/excel-com/02-object-model-knowledge-base.md` | Documentation-derived, attributed research foundation; no runtime implementation | Optional E4 preparation |
| Excel COM object-model analysis | `docs/research/excel-com/03-excel-object-model-analysis.md`, `knowledge/excel-object-model/generated/analysis/` | Documentation-derived navigation, candidate, typelib, and runtime evidence backlog; no runtime implementation | Optional E4 preparation |
| Excel COM typelib audit | `docs/research/excel-com/04-core-excel-typelib-audit.md`, `knowledge/excel-object-model/typelib/`, `knowledge/excel-object-model/generated/typelib/` | Installed Excel type-library evidence; no activation, dispatch implementation, or runtime conclusions | Optional E4 preparation |
| Excel COM Range runtime probe | `docs/research/excel-com/05-range-variant-safearray-runtime.md`, `knowledge/excel-object-model/runtime/`, `knowledge/excel-object-model/generated/runtime/` | Isolated raw-COM probe infrastructure; current host blocked before Range observation | Optional E4 preparation |
| Support/release | `SUPPORT_MATRIX.md`, `OPTIONAL_INTEGRATIONS_ROADMAP.md`, `docs/release/core-1.0-release-checklist.md` | Core 1.0 boundary defined; release audit active | M20 / E1-E5 / H0-H8 |

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

Optional after core 1.0:
 -> native XLL host integration (H1-H4)
 -> Ribbon/COM integration (H5, coordinated with C milestones)
 -> application cutover and async parity research (H6-H8)
```

## Release boundary

Core 1.0 excludes optional RTD, COM/Ribbon, custom task panes, autonomous
notification, and native desktop host integration. See ADR-0033,
`SUPPORT_MATRIX.md`, and `docs/architecture/native-xll-hosting-roadmap.md`;
those integrations remain separate post-1.0 decisions.

## Freeze policy

A design becomes `Implemented` only after code and tests exist. It becomes
`Stable` only after real Excel integration and API review.

The [native XLL header provenance inventory](docs/evidence/native-xll-header-provenance.md) and [C API validator](tools/excel-c-api-inventory/README.md) support EDNA-02 evidence and catalogue integrity; exhaustive declaration reconciliation remains open.
