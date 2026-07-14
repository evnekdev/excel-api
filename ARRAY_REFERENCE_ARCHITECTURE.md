# Array and Reference Architecture

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
