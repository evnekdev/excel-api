# Type Conversion Architecture

## Status

- **Status:** M3 semantic conversion and M4 return planning implemented.
- **Implemented in:** `convert.rs`, with owned storage in `value.rs`, logical
  return planning in `return_plan.rs`, and error types in `error.rs`.
- **Test coverage:** every supported scalar target, strict strings, arrays,
  missing/empty policy, reference rejection, resource limits, and numeric edge
  cases.
- **Remaining limitations:** worksheet coercion, Excel-owned API results,
  return materialization, and owned references.

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
    fn into_excel(self) -> Result<ExcelValue, ConversionError>;
}
```

`FromExcel<ExcelValue>` performs a deep copy with conservative default
`ConversionLimits`. Limit-aware entry points are also available on
`ExcelValue`, `ExcelString`, and `ExcelArray`.

Default limits are:

- 32,767 UTF-16 code units per string;
- 65,536 array elements;
- 16 MiB of conservatively counted destination storage;
- depth 8 (root is depth zero).

Array accounting includes `size_of::<ExcelValue>()` for every element plus all
owned UTF-16 payload bytes. Strict UTF-8 conversion budgets three bytes per
UTF-16 code unit. Checked arithmetic is used throughout, and the full array is
preflighted before destination allocation.

`xltypeInt` remains `ExcelValue::Integer(i32)`. `f64` accepts number and raw
integer inputs. Integer targets accept raw integers or finite integral numbers
within the exact target range; fractional, non-finite, and out-of-range inputs
have distinct errors. No worksheet-style coercion is implicit.

The existing `IntoExcel` scaffold emits `Integer` for `i16`, `i32`, and `u16`.
`u32` values within `i32` use `Integer`; larger `u32` values use an exactly
representable `Number`. This remains semantic planning, not ABI return storage.

`Option<T>` maps missing and empty to `None`. `OptionalValue<T>` and
`ExcelValue` preserve the distinction. Strict `String` conversion rejects
unpaired surrogates; `ExcelString` copies code units without Unicode loss.

## Return planning conversion

The implemented return path is:

```text
Rust T -> IntoExcel -> ExcelValue -> ExcelReturnValue -> ReturnPlan
```

`From<ExcelValue>` preserves integer, missing/empty, array order, and arbitrary
UTF-16 identity. Direct `String`/`&str` conversions to `ExcelReturnValue` retain
UTF-8 intent. Planning is deterministic and fallible through `ReturnError`;
ABI allocation is not part of this conversion layer.

## Coercion API

Expose explicit context-aware coercion separately:

```rust
WorksheetContext::coerce(...)
```

This may call `xlCoerce` and returns an `ExcelOwnedValue`.

Do not hide C API coercion inside every normal Rust conversion.
