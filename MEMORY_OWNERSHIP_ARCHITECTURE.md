# Memory and Ownership Architecture

## Status

- **Status:** M8 production Excel-owned RAII integration implemented; raw XLFree return integration remains deferred.
- **Implemented in:** `borrowed.rs` for callback views; `value.rs` and
  `convert.rs` for owned semantic values and bounded deep copies; and
  `return_plan.rs` for fully validated logical return plans; and
  `return_alloc.rs` for stable ABI return trees, consuming DLLFree handoff,
  and exact callback reclamation; and `excel_owned.rs` for callback-scoped
  Excel result ownership and exactly-once release.
- **Test coverage:** callback lifetime compile-fail tests, deep-copy
  independence, owned `Send + Sync + 'static`, arrays, strings, and conversion
  limits, deterministic return storage accounting, pointer stability,
  injected partial failures, local cleanup, every supported handed-off root,
  nested pointer stability, cross-thread cleanup, panic containment, and 1,000
  repeated handoff/callback cycles.
- **Remaining limitations:** Automated live Excel validation passed; interactive
  UI cases remain. Raw
  `xlbitXLFree` return integration remains deferred because per-call root
  reclamation is not proved and Microsoft does not document combining it with
  DLLFree.

M8's `CallCapability` implements the release backend with exactly
`Excel12v(xlFree, null, 1, [&mut root])`. Lifecycle call owners are dropped or
consumed before the runtime can unlink the callback entry point.

## Ownership domains

| Domain | Wrapper | Release |
|---|---|---|
| Callback input | `ExcelValueRef<'call>` | Excel |
| Owned semantic data | `ExcelValue` | Rust `Drop` |
| Logical return plan | `ReturnPlan` | Rust `Drop` |
| Excel C API result | `ExcelOwnedValue` | `xlFree` or `xlbitXLFree` transfer |
| XLL return before handoff | `ExcelReturn` | Rust `Drop` |
| XLL return after handoff | raw `*mut XLOPER12` | `xlAutoFree12` |

## Implemented callback-borrowing boundary

`RawExcelValue<'call>` is the sole unsafe entry point from a callback-owned
`XLOPER12` into the safe value layer. Its safety contract requires the root and
all reachable SDK storage to remain readable, immutable, and valid for
`'call`. One decoder masks `xlbitXLFree` and `xlbitDLLFree`, validates the base
tag before every union read, and produces `ExcelValueRef<'call>`.

Every pointer-bearing callback view carries `'call` and an explicit marker that
makes it neither `Send` nor `Sync`. The views do not implement `Clone` or
`Copy`. Safe code can therefore observe callback memory but cannot extend its
lifetime, move the view to another thread, mutate it, or free it.

## Implemented owned semantic boundary

`ExcelString`, `ExcelArray`, and `ExcelValue` contain only Rust-owned semantic
data. They carry no callback lifetime and expose no Excel raw pointer. Their
fields naturally provide `Send` and `Sync`; no unsafe trait implementation is
used. Deep conversion preflights element, string, aggregate-byte, and depth
limits before allocating, then copies every pointer-bearing payload.

The owned value model preserves `xltypeInt` as `ExcelValue::Integer(i32)` and
keeps missing and empty values distinct. References are rejected during deep
conversion because an owned workbook/sheet identity contract is not yet
specified.

## Implemented logical return boundary

`ExcelReturnValue` distinguishes return intent from general owned semantics.
Planning consumes it into a `ReturnPlan` containing only owned planned values,
exact ABI payload totals, and the selected `DllOwnedXloper12` strategy. The
planner uses safe Rust and creates no `XLOPER12`, pointer, prefixed string
buffer, ownership bit, or FFI call.

The accounting model counts exactly one future root allocation, one contiguous
element allocation for a multi, and one counted-string allocation per text
value. `total_bytes` is the exact sum of the root `XLOPER12`, array-element
`XLOPER12` storage, and UTF-16 prefix-plus-payload storage for the Prompt 05
layout. Rust container headers and allocator bookkeeping are deliberately
excluded and are not claimed as heap cost.

## Implemented stable local return allocation

`ReturnPlan::materialize` consumes a validated plan into opaque `ExcelReturn`:

```text
ExcelReturn
  -> Box<ReturnAllocation>
       root: XLOPER12                         // offset zero
       array_elements: Option<ReturnArrayBuffer>
         storage: Box<[XLOPER12]>
       string_buffers: Box<[ReturnUtf16Buffer]>
         each buffer.storage: Box<[XCHAR]>
       test-only root tracker
```

The concrete `ReturnAllocation` is `repr(C)` and tests prove `root` is at byte
offset zero. The string-buffer owner table is Rust container bookkeeping, not
ABI backing storage, so it remains outside Prompt 04 byte and allocation-count
totals. Debug counters track the ABI backing objects defined by that accounting
model: one root, an optional element block, and each independent string.

Construction consumes text into final boxed counted buffers first, freezes the
owner table, allocates one final boxed element block, initializes elements with
stable string pointers, verifies every planned total, constructs the root, and
finally boxes the root-first owner. No pointer targets the owner object itself;
moving `ExcelReturn` moves only its `Box` handle. Root, element, and string
addresses therefore remain unchanged.

Local cleanup follows Rust fields only. It never traverses raw tags or unions.
Normal error unwinding drops partial string and array storage, and test-only
atomic counters prove zero live backing objects after every injected failure.
Backing `Vec` reservations use fallible allocation APIs before conversion to
boxes. Stable Rust's final small `Box<ReturnAllocation>` allocation retains the
standard process-OOM behavior.

The materialized tree contains base type bits only.

## Implemented ownership handoff and callback reclamation

An `ExcelReturn` has three conceptual states:

```text
Local -- ordinary Drop --> Freed
Local -- consuming handoff --> HandedOff -- matching xlAutoFree12 --> Freed
```

`HandedOff -> local Drop`, a second handoff, access after `Freed`, and a second
callback are forbidden. The safe API makes a second handoff impossible by
consuming `ExcelReturn`. Once memory has been freed, an arbitrary duplicate
callback cannot be detected without itself reading freed memory; it is an
Excel/XLL ownership-contract violation outside safe recovery.

`ExcelReturn::into_raw_for_excel` moves the `Box<ReturnAllocation>` into a
local, checks debug invariants while that local owner still exists, sets
`xlbitDLLFree` on `root.xltype`, records test-only handoff tracking, and calls
`Box::into_raw`. The returned root is the allocation pointer cast to
`XLOPER12`, not a copied root. There is no allocation, formatting, logging,
Excel call, or other fallible work after `Box::into_raw`.

Only the root receives `xlbitDLLFree`; `xlbitXLFree` remains clear and all
nested elements retain base type bits only. Scalar roots are reclaimed just
like pointer-bearing roots because the top-level `ReturnAllocation` is always
heap allocated.

The callback's one unsafe reclamation primitive casts the root pointer back to
`ReturnAllocation` and drops `Box::from_raw` of that exact type. This depends on
`repr(C)` and `offset_of!(ReturnAllocation, root) == 0`; tests prove both pointer
identity at handoff and offset zero. It never constructs `Box<XLOPER12>` and
never traverses raw tags to release nested memory. The allocation must be freed
by the matching callback in the same loaded XLL binary; the internal layout is
not a serialization or cross-version ABI.

The reusable `unsafe extern "system"` callback body defensively ignores null
and contains unwinding panics with `catch_unwind` without logging, calling
Excel, or rethrowing. Production destructors are infallible. `panic = "abort"`
cannot be contained and retains its normal abort behavior.

## Initial return-root policy

Use one fresh heap-owned root per call.

This deliberately follows the simplest thread-safe model:

```text
Box<ReturnAllocation>
  -> root XLOPER12 at offset zero
  -> set xlbitDLLFree
  -> Excel
  -> xlAutoFree12
```

The book describes both per-call heap allocation and thread-local return slots.
The project chooses per-call allocation first because it is easier to audit and
does not require TLS lifetime management.

## Excel-owned results

Only values returned by Excel API calls may be released with `xlFree`.

M7 implements `ExcelOwnedValue<'call>` with a boxed root and a borrowed release
backend capability. Rust owns and eventually drops the root box; Excel owns
only auxiliary storage reachable from that root. The callback lifetime and a
raw-pointer marker keep the owner in the legal callback and make it neither
`Send` nor `Sync`.

The call-result constructor is crate-private and unsafe. The future call layer
must supply `XlFreeRequired` for results whose call/type metadata says Excel
allocated auxiliary storage and `NoReleaseRequired` otherwise. Microsoft
requires release for `xltypeStr`, `xltypeMulti`, and `xltypeRef`, and says it is
safe to pass any C-API result root to `xlFree`; Excel does not set
`xlbitXLFree` on ordinary call results. Big-data call metadata remains a Prompt
08 catalogue responsibility because the general `xlFree` reference does not
list `xltypeBigData` among its required types.

The implemented state machine is:

```text
Active -- explicit release / consuming copy / Drop --> Attempted
```

`Attempted` is terminal even when Excel reports failure. Retrying would be
unsafe because a failure code does not prove that no payload was released.
Explicit release returns the exact backend error. Drop contains backend panics,
ignores the diagnostic result, and never retries or panics. A consuming deep
copy composes conversion and release failures without losing either.

For an Excel-created multi the central decoder borrows every element, the
conversion layer deep-copies every supported element and UTF-16 payload, and
the backend receives only the top-level root.

`ExcelOwnedValue` tracks one of these states:

```text
Owned -> copied -> xlFree
Owned -> consumed -> xlbitXLFree transfer
Owned -> no-release-required
```

`xlbitXLFree` must be applied:

- after the C API call creates the value;
- after the value is no longer passed to other C API calls;
- immediately before return to Excel.

Microsoft says Excel frees auxiliary memory after copying the returned value,
but its example uses a static root and warns that the root must not be
overwritten while another thread is still using it. `xlFree` explicitly does
not destroy the root, and there is no callback for reclaiming an XLL-allocated
root carrying only `xlbitXLFree`. Therefore M7 exposes only a pre-commit owned
transfer token: it consumes the owner without releasing, retains the boxed root
and fallback release capability, sets no ownership bit, and exposes no pointer.
Both the required root lifetime and a leak-free committed-root cleanup strategy
must be integrated and proved in Prompt 08. Combining `xlbitXLFree` and
`xlbitDLLFree` is also deferred because Microsoft does not document that pair.

Official sources: [Memory Management in Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/memory-management-in-excel),
[xlFree](https://learn.microsoft.com/en-us/office/client-developer/excel/xlfree),
and [Excel4/Excel12](https://learn.microsoft.com/en-us/office/client-developer/excel/excel4-excel12).

## XLL-owned returns

`xlbitDLLFree` is applied only at final handoff.

`xlAutoFree12` now frees through the reconstructed Rust owner:

- top-level root allocation;
- string backing storage;
- multi element storage;
- string elements in multis;
- external reference storage when supported.

## Arrays

The book allows mixed ownership inside DLL-created multis, but warns that
consistency is essential.

The Rust design deliberately chooses the simpler invariant:

> Every pointer-bearing element inside a DLL-owned return tree is DLL-owned by
> the same top-level `ReturnAllocation`.

Therefore:

- Excel-owned strings are deep-copied;
- static dynamic mixtures are not used;
- one destructor frees the whole tree;
- arrays-of-arrays and arrays containing references are initially rejected.

## Failure safety

All validation and allocation happen before handoff. Before handoff, normal RAII
cleans up partial state. After handoff, no fallible work is permitted.

## `xlFree`

- Safe only for C API results.
- Never used for callback arguments.
- Never used for DLL-created/static values.
- For an Excel-owned `xltypeMulti`, call `xlFree` on the top-level result only,
  never on individual elements.

## Debug instrumentation

Optional feature:

- live allocation count;
- handoff/free counters;
- allocation ID;
- magic/layout version;
- state poisoning.

M6 unit tests use atomic, test-only live root/string/array counts, outstanding
handed-off-root count, and cumulative callback-free count. Production cleanup
does not depend on tracking or on a pointer registry.
