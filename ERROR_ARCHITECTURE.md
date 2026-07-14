# Error Architecture

## Status

- **Status:** M3 owned-value/conversion, M4 planning, M5 materialization, and
  M6 callback panic policy implemented.
- **Implemented in:** `error.rs`, with return error production in
  `return_plan.rs` and materialization error production in `return_alloc.rs`.
- **Test coverage:** precise numeric, UTF-16, shape, unsupported-reference,
  element, aggregate-byte, string, and depth failures.
- **Remaining limitations:** Excel-call, registration, and broader lifecycle
  error layers remain future milestones.

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

`ReturnMaterializationError` distinguishes plan/storage disagreement, UTF-8
encoded-length disagreement, counted-buffer disagreement, array-shape
disagreement, unsupported planned values, fallible backing allocation failure,
and test-only injected failures. It owns only scalar diagnostic metadata.

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

M6's safe consuming handoff is infallible. Every debug invariant check occurs
while the local `Box<ReturnAllocation>` is still owned, so it needs no new
post-materialization error type and cannot leak on an error path. After
`Box::into_raw` there is no fallible work.

`xlAutoFree12` has no error return channel. It ignores null, reclaims a valid
unique handoff, and contains any unwinding panic without formatting the panic
payload, logging, calling Excel, or rethrowing. Production return destructors
have no panic points. An aborting panic remains uncatchable by definition.
