# ADR-0020: Validate logical returns before ABI allocation

## Status

Accepted and implemented in M4.

## Decision

Insert an owned, safe planning domain between `ExcelValue` and future
`ReturnAllocation`:

```text
ExcelValue -> ExcelReturnValue -> ReturnPlan -> future ReturnAllocation
```

`ReturnPlan` consumes and retains the original logical payload, validated shape
and text metadata, exact ABI payload totals, and the sole ordinary-return
strategy `DllOwnedXloper12`. UTF-8 text remains UTF-8 while its exact UTF-16
length is counted; UTF-16 text retains arbitrary code units without
transcoding.

The Prompt 05 layout is one root allocation, one contiguous `XLOPER12` element
allocation for a multi, and one counted UTF-16 allocation per text. Project
budget bytes exactly sum root, element, and prefix-plus-payload ABI storage.
Rust container headers and allocator bookkeeping are explicitly outside this
total. Allocation count includes the root and every independently allocated
backing object.

Zero-dimensional multis, nested arrays, references, ABI-incompatible
dimensions, and configured resource excesses are rejected before ABI storage
exists. A semantic 0x0 `ExcelArray` remains valid but must return as another
value such as `Empty`.

## Consequences

Planning is deterministic, fully owned, `Send + Sync`, and safe Rust. It
creates no raw pointers, ABI values, ownership bits, or FFI calls. Prompt 05
must materialize exactly the retained plan without changing policy or claiming
allocator bookkeeping as part of the planned ABI-byte total.
