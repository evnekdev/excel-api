# Memory and Ownership Architecture

## Status

- **Status:** Partially implemented through M4.
- **Implemented in:** `borrowed.rs` for callback views; `value.rs` and
  `convert.rs` for owned semantic values and bounded deep copies; and
  `return_plan.rs` for fully validated logical return plans.
- **Test coverage:** callback lifetime compile-fail tests, deep-copy
  independence, owned `Send + Sync + 'static`, arrays, strings, and conversion
  limits, plus deterministic return storage accounting and return limits.
- **Remaining limitations:** Excel-owned API results, DLL-owned return
  allocation, ownership-bit handoff, and `xlAutoFree12` remain future
  milestones.

## Ownership domains

| Domain | Wrapper | Release |
|---|---|---|
| Callback input | `ExcelValueRef<'call>` | Excel |
| Owned semantic data | `ExcelValue` | Rust `Drop` |
| Logical return plan | `ReturnPlan` | Rust `Drop` |
| Excel C API result | `ExcelOwnedValue` | `xlFree` or `xlbitXLFree` transfer |
| XLL return before handoff | `ExcelReturn` | Rust `Drop` |
| XLL return after handoff | raw `*mut XLOPER12` | `xlAutoFree12` |

## Implemented callback-borrowing boundary

`RawExcelValue<'call>` is the sole unsafe entry point from a callback-owned
`XLOPER12` into the safe value layer. Its safety contract requires the root and
all reachable SDK storage to remain readable, immutable, and valid for
`'call`. One decoder masks `xlbitXLFree` and `xlbitDLLFree`, validates the base
tag before every union read, and produces `ExcelValueRef<'call>`.

Every pointer-bearing callback view carries `'call` and an explicit marker that
makes it neither `Send` nor `Sync`. The views do not implement `Clone` or
`Copy`. Safe code can therefore observe callback memory but cannot extend its
lifetime, move the view to another thread, mutate it, or free it.

## Implemented owned semantic boundary

`ExcelString`, `ExcelArray`, and `ExcelValue` contain only Rust-owned semantic
data. They carry no callback lifetime and expose no Excel raw pointer. Their
fields naturally provide `Send` and `Sync`; no unsafe trait implementation is
used. Deep conversion preflights element, string, aggregate-byte, and depth
limits before allocating, then copies every pointer-bearing payload.

The owned value model preserves `xltypeInt` as `ExcelValue::Integer(i32)` and
keeps missing and empty values distinct. References are rejected during deep
conversion because an owned workbook/sheet identity contract is not yet
specified.

## Implemented logical return boundary

`ExcelReturnValue` distinguishes return intent from general owned semantics.
Planning consumes it into a `ReturnPlan` containing only owned planned values,
exact ABI payload totals, and the selected `DllOwnedXloper12` strategy. The
planner uses safe Rust and creates no `XLOPER12`, pointer, prefixed string
buffer, ownership bit, or FFI call.

The accounting model counts exactly one future root allocation, one contiguous
element allocation for a multi, and one counted-string allocation per text
value. `total_bytes` is the exact sum of the root `XLOPER12`, array-element
`XLOPER12` storage, and UTF-16 prefix-plus-payload storage for the Prompt 05
layout. Rust container headers and allocator bookkeeping are deliberately
excluded and are not claimed as heap cost.

## Initial return-root policy

Use one fresh heap-owned root per call.

This deliberately follows the simplest thread-safe model:

```text
Box<ReturnAllocation>
  -> root XLOPER12 at offset zero
  -> set xlbitDLLFree
  -> Excel
  -> xlAutoFree12
```

The book describes both per-call heap allocation and thread-local return slots.
The project chooses per-call allocation first because it is easier to audit and
does not require TLS lifetime management.

## Excel-owned results

Only values returned by Excel API calls may be released with `xlFree`.

`ExcelOwnedValue` tracks one of these states:

```text
Owned -> copied -> xlFree
Owned -> consumed -> xlbitXLFree transfer
Owned -> no-release-required
```

`xlbitXLFree` must be applied:

- after the C API call creates the value;
- after the value is no longer passed to other C API calls;
- immediately before return to Excel.

## XLL-owned returns

`xlbitDLLFree` is applied only at final handoff.

`xlAutoFree12` frees:

- top-level root allocation;
- string backing storage;
- multi element storage;
- string elements in multis;
- external reference storage when supported.

## Arrays

The book allows mixed ownership inside DLL-created multis, but warns that
consistency is essential.

The Rust design deliberately chooses the simpler invariant:

> Every pointer-bearing element inside a DLL-owned return tree is DLL-owned by
> the same top-level `ReturnAllocation`.

Therefore:

- Excel-owned strings are deep-copied;
- static dynamic mixtures are not used;
- one destructor frees the whole tree;
- arrays-of-arrays and arrays containing references are initially rejected.

## Failure safety

All validation and allocation happen before handoff. Before handoff, normal RAII
cleans up partial state. After handoff, no fallible work is permitted.

## `xlFree`

- Safe only for C API results.
- Never used for callback arguments.
- Never used for DLL-created/static values.
- For an Excel-owned `xltypeMulti`, call `xlFree` on the top-level result only,
  never on individual elements.

## Debug instrumentation

Optional feature:

- live allocation count;
- handoff/free counters;
- allocation ID;
- magic/layout version;
- state poisoning.
