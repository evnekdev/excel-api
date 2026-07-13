# Memory and Ownership Architecture

## Ownership domains

| Domain | Wrapper | Release |
|---|---|---|
| Callback input | `ExcelValueRef<'call>` | Excel |
| Owned semantic data | `ExcelValue` | Rust `Drop` |
| Excel C API result | `ExcelOwnedValue` | `xlFree` or `xlbitXLFree` transfer |
| XLL return before handoff | `ExcelReturn` | Rust `Drop` |
| XLL return after handoff | raw `*mut XLOPER12` | `xlAutoFree12` |

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
