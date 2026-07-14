# Array and Reference Architecture

## Status

- **Status:** Borrowed and owned flat arrays, logical planning, and DLLFree
  return multis implemented through M6; references remain borrowed only.
- **Implemented in:** `borrowed.rs` (`ExcelArrayView` and reference views),
  `value.rs` (`ExcelArray`), `convert.rs` (bounded deep copy), and
  `return_plan.rs` (`ExcelReturnArray` and `PlannedArray`), and
  `return_alloc.rs` (stable boxed `XLOPER12` element storage, consuming
  handoff, and callback reclamation).
- **Test coverage:** shape overflow/mismatch, row-major indexing, rows,
  columns, mixed values, deep-copy independence, element/byte/depth limits,
  nested-array rejection, reference rejection, ABI dimension checks, exact
  element/string accounting, zero-dimensional return rejection, pointer
  stability, deep string buffers, partial-failure cleanup, root-only DLLFree,
  nested pointer survival through handoff, and exactly-once callback cleanup.
- **Remaining limitations:** owned references, reference coercion, FP12 safe
  wrappers and Excel-owned API result arrays.

## Arrays

### `FP12`

Specialized floating-point array ABI.

Use when:

- only doubles are required;
- performance and compactness matter;
- mixed values and strings are unnecessary.

### `xltypeMulti`

Mixed rectangular array of `XLOPER12`.

Properties:

- row-major storage;
- may contain numbers, Booleans, errors, strings, empty values;
- Excel-created multis are released as one top-level object;
- never call `xlFree` on individual elements of an Excel-created multi.

Initial return restrictions:

- no arrays of arrays;
- no arrays containing references;
- flat rectangular arrays only.

The callback-borrowing implementation applies the same flatness restriction to
borrowed multis. `ExcelArrayView<'call>` validates positive SDK-bounded
dimensions, checked row-major element count, a non-null element pointer, and
every element tag before it is exposed. Nested multis and reference elements
are rejected. Indexing and row/column iterators borrow elements through the one
audited `XLOPER12` decoder and allocate nothing.

`ExcelArray` owns a `Box<[ExcelValue]>` and keeps rows and columns private. Its
constructor checks multiplication and exact element count and rejects nested
arrays. It provides immutable checked indexing, row slices, column iterators,
and row-major iteration.

`ExcelReturnArray` is the separate fully owned logical return form. Planning
checks shape, Excel 12 row/column bounds, a conservative project element limit,
depth, total bytes, and allocations before any ABI storage exists. A successful
`PlannedArray` preserves row-major order and contains only supported scalar or
text elements. Prompt 05 will allocate one contiguous `XLOPER12` element block
and one counted UTF-16 buffer per text element.

M5 implements that layout exactly: one `Box<[XLOPER12]>` with `rows * columns`
elements in row-major order, plus one final counted UTF-16 box per text element.
The element box reaches its final address before the root `lparray` pointer is
constructed. Every element has base type bits only. Rust owner fields, not raw
element tags, drive local cleanup.

M6 consumes the top-level owner, marks only the multi root with
`xlbitDLLFree`, and leaves every element at base type bits. Element and nested
string addresses remain stable while Excel owns the raw root. `xlAutoFree12`
casts that offset-zero root back to the exact `ReturnAllocation`; dropping its
fields releases the element block and every nested string without walking the
raw multi.

### Empty return-array policy

A zero row or zero column `xltypeMulti` is rejected as
`ReturnError::EmptyArrayUnsupported`. Callback multis already require positive
dimensions, and a zero-element pointer contract would complicate safe
materialization without representing anything that `ExcelValue::Empty` cannot
express more clearly. `ExcelArray` remains able to represent 0x0 semantic data;
the restriction applies at the Excel return boundary.

The borrowed view intentionally validates every element eagerly and decodes an
element again when accessed. Prompt 03 preserves that correctness-first trade-
off. Bounded deep copy adds a non-allocating preflight traversal before the
materialization traversal, so conversion performs additional decoding rather
than allocating before all limits and unsupported values are known.

## Registration effect

Reference-preserving general arguments may yield:

- `xltypeRef`;
- `xltypeSRef`;
- scalar values;
- multis;
- missing/nil.

Value-only general arguments cause references to be converted to:

- scalar value;
- `xltypeMulti`;
- `xltypeNil`.

This distinction must be represented in registration descriptors.

## References

### `xltypeSRef`

- one rectangular area;
- current sheet context;
- structure stored inline.

### `xltypeRef`

- one or more rectangular areas;
- explicit sheet ID;
- separately allocated `XLMREF12`.

Safe public types:

```rust
ExcelReference<'call>
OwnedExcelReference
ReferenceArea
```

Reference values remain distinct from arrays.

`ExcelReference<'call>` preserves the ABI distinction with separate
`ExcelSingleReference<'call>` (`xltypeSRef`) and
`ExcelMultiReference<'call>` (`xltypeRef`) variants. The single-reference count,
multi-reference pointer/count, sheet ID, and each rectangular area are
validated and observed directly; the borrowing layer performs no coercion or
worksheet lookup.

Owned semantic conversion returns `ConversionError::UnsupportedReference` for
both reference forms. It does not copy a context-dependent sheet identifier,
call `xlCoerce`, or perform a worksheet lookup.

## Coercion

When a range is used as data, coercing to `xltypeMulti` is often simpler.

Coercion may fail for:

- invalid/deleted references;
- uncalculated cells in thread-safe contexts;
- extreme ranges;
- context restrictions.

## Current versus active

Reference APIs must distinguish:

- active workbook/sheet/cell;
- current workbook/sheet/cell being calculated.

The library must avoid APIs whose names blur this distinction.
