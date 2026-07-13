# Type Conversion Architecture

## Two independent conversion layers

### Excel-selected conversion

Registration type text controls what Excel converts before calling the XLL.

Examples:

- value-only general arguments coerce ranges to values/multis;
- reference-preserving general arguments may pass `xltypeRef`/`xltypeSRef`;
- scalar registration forms may cause Excel to convert before the thunk runs.

If Excel cannot perform required conversion, the function may not be called.

### Rust-selected conversion

Once the thunk receives a valid ABI value:

```rust
FromExcel<'call>
IntoExcel
```

perform explicit Rust conversions.

## Important semantic differences

- Missing argument and empty referenced cell are distinct.
- Excel may coerce empty to zero or empty string; Rust conversion must not do so
  unless requested.
- Worksheet numbers are doubles; Rust integer conversion policy is explicit and
  checked rather than blindly copying Excel's truncation behavior.
- Arrays retain element types; Excel does not necessarily convert each element
  of a mixed array to the scalar type expected by the user.

## Traits

```rust
pub trait FromExcel<'call>: Sized {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError>;
}

pub trait IntoExcel {
    fn into_excel_value(self) -> Result<ExcelReturnValue, ConversionError>;
}
```

## Coercion API

Expose explicit context-aware coercion separately:

```rust
WorksheetContext::coerce(...)
```

This may call `xlCoerce` and returns an `ExcelOwnedValue`.

Do not hide C API coercion inside every normal Rust conversion.
