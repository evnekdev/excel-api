# Architecture Index

| Area | Document | Status | Milestones |
|---|---|---|---|
| Overall | `ARCHITECTURE.md` | Proposed | M0+ |
| ABI | `ABI_ARCHITECTURE.md` | Implemented | M1 |
| Memory | `MEMORY_OWNERSHIP_ARCHITECTURE.md` | M2-M8 implemented (raw XLFree return transfer intentionally deferred) | M1-M8 |
| Strings | `STRING_ARCHITECTURE.md` | Partial (DLLFree returns implemented) | M1-M8 |
| Conversion | `TYPE_CONVERSION_ARCHITECTURE.md` | M3-M5 implemented | M2-M5 |
| Arrays/references | `ARRAY_REFERENCE_ARCHITECTURE.md` | Partial (DLLFree return multis implemented) | M2-M8 |
| Threading | `THREADING_ARCHITECTURE.md` | Partial (thread-independent AutoFree implemented) | M1-M17 |
| Contexts | `RUNTIME_CONTEXT_ARCHITECTURE.md` | M8 callback capabilities implemented | M4-M12 |
| C API calls | `EXCEL_CALL_ARCHITECTURE.md` | M8 narrow production catalogue implemented | M7-M12 |
| Lifecycle | `CALLBACK_LIFECYCLE_ARCHITECTURE.md` | M8 implemented; automated live Excel passed | M6-M12 |
| Registration | `REGISTRATION_ARCHITECTURE.md` | M8 manual oracle, M9 generation, and M10 diagnostic conformance implemented | M8-M10 |
| Errors | `ERROR_ARCHITECTURE.md` | Partial (M3-M6 return policy implemented) | M2-M13 |
| Testing | `TESTING_ARCHITECTURE.md` | Partial (automated real-Excel smoke/soak harness implemented) | All |
| Macros | `PROC_MACRO_ARCHITECTURE.md` | M9 generation and M10 compile-time conformance implemented | M9-M10 |
| Packaging | `PACKAGING_ARCHITECTURE.md` | Minimal M8 XLL build implemented | M14 |
| Async | `ASYNC_ARCHITECTURE.md` | M16 implementation in progress; live validation pending | M16 |
| Dispatcher | `MAIN_THREAD_DISPATCH_ARCHITECTURE.md` | Planned | M17 |
| RTD | `RTD_STREAMING_ARCHITECTURE.md` | Planned | M18 |
| COM/Ribbon | `COM_ARCHITECTURE.md`, `RIBBON_UI_ARCHITECTURE.md` | Planned | M19+ |

## Dependency order

```text
ABI
 -> memory/string
 -> conversion/arrays/references
 -> threading/contexts/calls
 -> lifecycle/registration
 -> macros
 -> async/RTD/COM
```

## Freeze policy

A design becomes `Implemented` only after code and tests exist. It becomes
`Stable` only after real Excel integration and API review.
