# Array and Reference Architecture

## Status

- **Status:** Borrowed and owned flat arrays implemented through M3;
  references remain borrowed only.
- **Implemented in:** `borrowed.rs` (`ExcelArrayView` and reference views),
  `value.rs` (`ExcelArray`), and `convert.rs` (bounded deep copy).
- **Test coverage:** shape overflow/mismatch, row-major indexing, rows,
  columns, mixed values, deep-copy independence, element/byte/depth limits,
  nested-array rejection, and reference rejection.
- **Remaining limitations:** owned references, reference coercion, FP12 safe
  wrappers, return multis, and Excel-owned API result arrays.

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
