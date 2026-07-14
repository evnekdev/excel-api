# Error Architecture

## Status

- **Status:** M3 owned-value and conversion errors implemented.
- **Implemented in:** `error.rs`.
- **Test coverage:** precise numeric, UTF-16, shape, unsupported-reference,
  element, aggregate-byte, string, and depth failures.
- **Remaining limitations:** return, Excel-call, registration, lifecycle, and
  panic-boundary error layers remain future milestones.

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
