# Architecture Index

| Area | Document | Status | Milestones |
|---|---|---|---|
| Overall | `ARCHITECTURE.md` | Proposed | M0+ |
| ABI | `ABI_ARCHITECTURE.md` | Implemented | M1 |
| Memory | `MEMORY_OWNERSHIP_ARCHITECTURE.md` | Partial (M2-M4 implemented) | M1-M7 |
| Strings | `STRING_ARCHITECTURE.md` | Partial (planning implemented; allocation pending) | M1-M8 |
| Conversion | `TYPE_CONVERSION_ARCHITECTURE.md` | M3-M4 implemented | M2-M5 |
| Arrays/references | `ARRAY_REFERENCE_ARCHITECTURE.md` | Partial (flat return planning implemented) | M2-M8 |
| Threading | `THREADING_ARCHITECTURE.md` | Proposed | M1-M17 |
| Contexts | `RUNTIME_CONTEXT_ARCHITECTURE.md` | Proposed | M4-M12 |
| C API calls | `EXCEL_CALL_ARCHITECTURE.md` | Proposed | M7-M12 |
| Lifecycle | `CALLBACK_LIFECYCLE_ARCHITECTURE.md` | Proposed | M6-M12 |
| Registration | `REGISTRATION_ARCHITECTURE.md` | Proposed | M8-M10 |
| Errors | `ERROR_ARCHITECTURE.md` | Partial (M3-M4 conversion/return errors implemented) | M2-M13 |
| Testing | `TESTING_ARCHITECTURE.md` | Proposed | All |
| Macros | `PROC_MACRO_ARCHITECTURE.md` | Planned | M9-M10 |
| Packaging | `PACKAGING_ARCHITECTURE.md` | Planned | M14 |
| Async | `ASYNC_ARCHITECTURE.md` | Planned | M16 |
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
