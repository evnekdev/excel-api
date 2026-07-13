# Error Architecture

## Layers

```rust
AbiError
ConversionError
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
