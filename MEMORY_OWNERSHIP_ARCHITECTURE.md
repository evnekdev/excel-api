# Memory and Ownership Architecture

## Status

Proposed design for the production memory layer of `excel-api`. It must be verified against the official Excel SDK headers and runtime behavior before being treated as complete.

## Safety goals

The implementation must prevent callback-lifetime escape, use-after-free, double-free, incorrect allocator use, invalid tagged-union reads, pointer invalidation during return construction, and leaks during errors, panics, recalculation, workbook close, or add-in unload.

## Ownership domains

Every raw Excel value belongs to exactly one domain:

| Domain | Origin | Lifetime | Release | Safe representation |
|---|---|---|---|---|
| Borrowed callback input | Excel calls a thunk | Current callback | Excel | `ExcelValueRef<'call>` |
| Owned Rust value | Deep copy or user result | Normal Rust scope | Rust `Drop` | `ExcelValue` |
| Excel-owned API result | `Excel12v` | Until verified Excel release | Excel-call layer | `ExcelOwnedValue` |
| XLL return under construction | Return materializer | Until handoff/error | Rust `Drop` | `ExcelReturn` |
| XLL return after handoff | Returned with DLL-free ownership | Until `xlAutoFree12` | `xlAutoFree12` | raw pointer only at FFI boundary |

Ownership is established by constructors and state transitions, never guessed in general-purpose code from a raw pointer.

## Public and internal value types

### `ExcelValueRef<'call>`

A validated borrowed view over callback arguments:

```rust
pub enum ExcelValueRef<'call> {
    Number(f64),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(ExcelStr<'call>),
    Array(ExcelArrayView<'call>),
    Reference(ExcelReference<'call>),
}
```

It cannot outlive the callback, owns nothing, and is not `Send` or `Sync` by default. Retaining data requires a deep copy.

### `ExcelValue`

An ordinary Rust-owned value independent of Excel memory:

```rust
pub enum ExcelValue {
    Number(f64),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(String),
    Array(ExcelArray),
}
```

It contains no Excel pointers and may be cached or moved to workers when its contents permit.

### `ExcelReturnValue`

A logical return tree before ABI allocation. It should initially support only numbers, Booleans, errors, empty values, strings, and flat rectangular arrays. Unsupported variants are rejected before any raw pointer is published.

### `ExcelReturn`

An opaque RAII owner of a materialized XLL return. Before handoff it behaves like normal Rust memory. Handoff consumes it and produces the raw root pointer returned to Excel.

### `ExcelOwnedValue`

An RAII wrapper for results allocated or managed by Excel APIs. It carries an explicit, verified release policy and must never be freed through `xlAutoFree12`.

## Borrowed input parsing

Only generated/manual thunks may construct the internal raw wrapper:

```rust
pub(crate) struct RawExcelValue<'call> {
    ptr: NonNull<XLOPER12>,
    _lifetime: PhantomData<&'call XLOPER12>,
}
```

The parser must:

1. mask ownership bits before matching the base type;
2. access a union field only after validating the corresponding tag;
3. validate nullness, alignment, signed lengths, and dimensions;
4. use checked arithmetic for `rows * columns` and aggregate sizes;
5. create slices only after pointer and length validation;
6. enforce string, array, nesting, and aggregate-byte limits;
7. never free callback inputs;
8. convert malformed values into controlled errors rather than speculative dereferences.

Unsafe functions must document pointer validity, lifetime, tag/union correspondence, ownership, threading assumptions, and allowed callers.

## Strings

Excel 12 strings are length-prefixed UTF-16. `ExcelStr<'call>` exposes the payload units without exposing ownership or the prefix.

Rules:

- verify the exact prefix representation and maximum length against the SDK;
- allow interior NUL because length, not termination, defines the string;
- provide strict and explicitly lossy UTF-16 conversions;
- validate length before constructing a slice;
- returned strings store prefix and payload in stable heap memory.

## Arrays

Borrowed arrays expose dimensions and a borrowed element span. All dimension conversions and element counts are checked. Owned arrays enforce `values.len() == rows * columns`.

The first return implementation should support flat rectangular arrays of numbers, Booleans, errors, empty values, and strings. Nested `xltypeMulti` values and returned references are deferred.

## Two-phase return construction

Return conversion is split into:

```text
Rust result -> ExcelReturnValue -> ReturnPlan -> ExcelReturn -> raw pointer
```

`ReturnPlan` validates all types, dimensions, lengths, limits, and allocation sizes without publishing pointers. Materialization allocates final stable buffers and only then patches raw pointers.

This separation ensures semantic/conversion failures happen while all data is still ordinary RAII-owned Rust memory.

## Self-contained return allocation

Complex returns use one top-level allocation:

```rust
#[repr(C)]
struct ReturnAllocation {
    root: XLOPER12,              // field zero, offset zero
    header: ReturnHeader,
    strings: Vec<Box<[u16]>>,
    arrays: Vec<Box<[XLOPER12]>>,
}
```

The concrete storage may later be optimized, but these invariants are fixed:

- `root` is at offset zero;
- all published pointers target stable allocations;
- no pointer targets stack memory or a `Vec` that may reallocate;
- all backing buffers are allocated before pointer publication;
- construction is fully RAII-owned before handoff;
- production correctness does not depend on a global pointer registry.

A diagnostic header may contain a magic number, layout version, sequence number, and state marker. It is defense-in-depth, not permission to accept arbitrary pointers.

## Stable-pointer construction sequence

1. Validate the complete logical tree.
2. Compute all counts and byte sizes with checked arithmetic.
3. Allocate final string and array buffers.
4. Convert growable buffers to stable boxed slices.
5. Populate raw elements only after all backing addresses are stable.
6. Construct the root `XLOPER12`.
7. Apply DLL-free ownership only during final handoff.
8. Consume the allocation with `Box::into_raw`.

Do not build self-references and then move the parent object. Separate stable boxed buffers are preferred over a complicated pinned self-referential structure.

## Handoff state machine

```text
logical -> planned -> materialized (RAII) -> handed off -> freed by xlAutoFree12
                          |
                          +-> local Drop on failure
```

Handoff consumes `ExcelReturn`. After `Box::into_raw`, Rust must not drop the allocation locally. Handoff is the final non-fallible operation in the thunk: no logging, formatting, metadata lookup, validation, or allocation occurs afterward.

## `xlAutoFree12`

The callback should do only the following:

1. validate the pointer according to the verified ABI contract;
2. cast the offset-zero root pointer back to `ReturnAllocation`;
3. perform non-allocating debug-header checks;
4. reconstruct `Box<ReturnAllocation>`;
5. drop it exactly once;
6. catch any unexpected panic and never unwind across FFI.

Cleanup must not rely on the originating thread, call back into Excel, acquire fragile locks, or perform heap-allocating logging.

## Ownership flags

Ownership-bit handling is centralized in one internal module.

- callback arguments are borrowed and are never freed by this library;
- Excel API results use `ExcelOwnedValue` and a verified release policy;
- XLL returns receive DLL-free ownership only at handoff;
- copies deliberately clear or replace ownership bits;
- invalid ownership combinations are unrepresentable in safe code.

## Excel-owned API results

`Excel12v` results follow a separate path:

```text
Excel12v -> raw result slot -> ExcelOwnedValue -> borrow/copy -> verified Excel release
```

Release behavior is documented per supported call category. Such values are never adopted as XLL return allocations and never passed to `xlAutoFree12`.

## Panic and failure safety

Every exported callback uses `catch_unwind`.

Recommended order:

1. parse borrowed arguments;
2. invoke the Rust function;
3. convert to a logical return;
4. plan and materialize the return;
5. hand off as the final non-fallible step;
6. return the raw pointer.

Panics or errors before handoff trigger ordinary RAII cleanup. Destructors in the memory layer must not panic. The library should not claim recoverable out-of-memory behavior unless fallible allocation is used throughout; it should nevertheless reject absurd sizes before allocation.

## Threading

Each return allocation is independent and contains no thread-affine state. Callback views remain on the invoking thread. Async/worker code deep-copies required inputs first. No COM pointer, UI handle, mutex guard, task handle, or thread-affine resource is stored in `ReturnAllocation`.

The design must not assume Excel invokes `xlAutoFree12` on the originating calculation thread unless this is verified and intentionally documented.

## References

References are context-bearing Excel objects, not plain process-owned memory. Initially they are borrowed only, may be coerced/copied while a legal Excel context exists, and are not accepted as return values. Owned references are deferred until workbook identity, release rules, and threading restrictions are understood.

## Unsafe module boundaries

Unsafe behavior is concentrated in:

```text
excel-api-sys              raw ABI declarations only
excel-api/src/raw          tag decoding, union reads, raw validation
excel-api/src/memory       planning, stable allocation, handoff, cleanup
excel-api/src/excel_owned  Excel API result ownership and release
```

Safe conversion and registration code must not dereference raw pointers.

## Debug instrumentation

An optional `memory-debug` feature may provide atomic live-allocation counts, allocation IDs, header validation, state poisoning, byte counters, and a bounded diagnostic registry. The registry is observational only; release correctness must not depend on it.

## Rejected alternatives

- **Global registry as primary ownership:** adds synchronization, global mutable state, and new leak modes.
- **Returning pointers into `String` or growable `Vec`:** addresses may move or be freed too early.
- **Leaking returns:** recalculation causes unbounded growth.
- **One type for all ownership states:** lifetimes and destructors become ambiguous.
- **Safe raw `XLOPER12` construction by users:** pointer stability and allocator matching cannot be guaranteed.
- **Zero-copy user buffers initially:** ownership transfer and Excel retention are too easy to misuse.

## Verification blockers

Before claiming production safety, verify:

1. exact `XLOPER12` layout, alignment, tags, and ownership bits;
2. exact UTF-16 prefix and limits;
3. array order and zero-size behavior;
4. `xlAutoFree12` signature, invocation conditions, and thread behavior;
5. release rules for supported `Excel12v` results;
6. behavior during cancellation, workbook close, and add-in unload;
7. how long Excel may retain returned pointers;
8. legality of nested arrays and references as returns;
9. practical aggregate result-size limits.
