# Migration from Excel-DNA

Excel-DNA offers a managed integration model; `excel-api` is a native Rust XLL
framework. Map concepts, not implementation details:

| Excel-DNA-style concern | `excel-api` approach |
| --- | --- |
| Managed function export | `#[excel_function]` plus native thunk/registration metadata. |
| `ExcelReference` | Callback-borrowed `ExcelReferenceArg<'call>` / `ExcelReference<'call>`. |
| Managed return | Owned semantic value → `ReturnPlan` → DLLFree allocation. |
| Excel API call | Typed descriptor through a legal callback context. |
| Async task | Preview bounded async UDF generation; no worker Excel calls. |
| UI/RTD integration | Deferred or experimental; not a core replacement promise. |

Do not port managed object-lifetime assumptions directly. Native callback input
storage is borrowed, raw ABI unions require tag validation, and Excel/DLL
ownership protocols are explicit.
