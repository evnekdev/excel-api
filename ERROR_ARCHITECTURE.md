# Error Architecture

## Status

- **Status:** M3 owned-value/conversion and M4 return-planning errors
  implemented.
- **Implemented in:** `error.rs`, with return error production in
  `return_plan.rs`.
- **Test coverage:** precise numeric, UTF-16, shape, unsupported-reference,
  element, aggregate-byte, string, and depth failures.
- **Remaining limitations:** return materialization/handoff, Excel-call,
  registration, lifecycle, and panic-boundary error layers remain future
  milestones.

## Layers

```rust
AbiError
ConversionError
OwnedValueError
Utf16ConversionError
ReturnError
ExcelCallError
RegistrationError
LifecycleError
PanicError
```

## Worksheet errors

```rust
ExcelError
```

represents Excel-visible error values.

## Mapping policy

- invalid argument type -> usually `#VALUE!`;
- invalid reference -> `#REF!`;
- numeric domain/range failure -> `#NUM!`;
- unavailable result -> `#N/A`;
- panic -> controlled default error plus diagnostics.

Mappings must be explicit and overridable where appropriate.

## Implemented conversion distinctions

`ConversionError` distinguishes unexpected types, unsupported references,
invalid UTF-16, non-finite/fractional/out-of-range numbers, invalid array
shape, string/element/aggregate limits, nested arrays, depth limits, and a
borrowed-value decode failure. Errors contain only owned metadata and never a
callback pointer.

`OwnedValueError` reports direct array shape overflow/mismatch and nested-array
construction. `Utf16ConversionError` is the strict owned UTF-16 decoding
failure. All implement `Debug`, `Display`, and `std::error::Error`.

`ReturnError` independently distinguishes the Excel string hard limit, project
string/array/byte/allocation/depth limits, invalid or zero-dimensional array
shapes, ABI dimension overflow, nested arrays, references, unsupported semantic
variants, and checked byte/allocation overflows. It contains no raw pointers or
callback lifetimes. Arbitrary UTF-16 is not an error for return planning.

## C API return codes

Preserve return-code distinctions such as:

- invalid function;
- abort;
- uncalculated;
- invalid context;
- not thread safe.

Do not collapse all C API failures into `#VALUE!`.

## Result<T, E>

`IntoExcel` support for `Result<T, E>` requires an explicit trait that maps the
error into an Excel-visible value and diagnostic record.

## Destructors

Release failures are diagnostic only; destructors never panic.
