# ADR-0011: Excel-owned API results

## Decision

Use callback-scoped `ExcelOwnedValue<'call>` RAII with a stable boxed root and
a borrowed, release-only backend capability. Explicit release, consuming copy,
and Drop each converge on one terminal release attempt. A failed attempt is not
retried because Excel may have released some or all auxiliary storage before
reporting failure.

The release policy is supplied by verified call/result metadata, not inferred
from pointer shape or an Excel-set ownership bit. Microsoft requires `xlFree`
for Excel12 results of type string, multi, and reference, says other C-API
result roots are safe to pass, and documents `xlbitXLFree` as a bit the XLL sets
only when returning an Excel-created result back to Excel.

## Transfer deferral

No raw `xlbitXLFree` transfer API is exposed in M7. A safe pre-commit token
consumes the owner, retains its root and fallback release, sets no bit, and
exposes no pointer. Microsoft documents that
Excel copies the return and frees auxiliary storage, and warns against a shared
static root in thread-safe UDFs. It also documents that `xlFree` does not
destroy the root. It does not provide a cleanup callback for a root carrying
only XLFree or document combining XLFree with DLLFree. Prompt 08 must prove a
per-call root lifetime and cleanup strategy before exposing a raw pointer.

Sources: [Memory Management in Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/memory-management-in-excel),
[xlFree](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfree),
[Multithreaded recalculation](https://learn.microsoft.com/en-us/office/client-developer/excel/multithreaded-recalculation-in-excel),
and [Excel4/Excel12](https://learn.microsoft.com/en-us/office/client-developer/excel/excel4-excel12).
