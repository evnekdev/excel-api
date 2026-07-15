# Runtime Context Architecture

## Context tokens

```rust
WorksheetContext<'call>
ThreadSafeContext<'call>
MacroSheetContext<'call>
CommandContext<'call>
LifecycleContext<'call>
```

## Purpose

The context type is a capability token. It controls which C API operations are
available.

## Example capability split

### `ThreadSafeContext`

May expose verified thread-safe C API-only functions such as:

- `xlFree`;
- selected `xlCoerce`;
- stack/query helpers;
- abort polling;
- sheet ID/name helpers where verified.

### `WorksheetContext`

May expose worksheet-safe calls but not commands.

### `MacroSheetContext`

May expose additional macro-sheet functionality but is never thread-safe.

### `CommandContext`

May mutate Excel state and call command-equivalent operations.

### `LifecycleContext`

May register/unregister and initialize runtime state.

## Construction

Only runtime/exported thunks may construct contexts.

Context values:

- are not forgeable publicly;
- are not retained beyond the callback;
- carry runtime state needed for call validation.

## M7 release capability

`ExcelOwnedValue<'call>` borrows a release-only backend for `'call`. This is
the selected capability model: a context-free `'static` destructor is not
allowed. Microsoft permits C API calls only while Excel has passed control to
the XLL, forbids them from XLL-created threads and `DllMain`, and makes
`xlFree` the sole permitted cleanup callback after `xlretAbort` or
`xlretUncalced`. Prompt 08 will construct the backend from the runtime context
and keep call pointers linked through every owner drop.
## M8 implementation

`WorksheetContext`, `ThreadSafeContext`, and `LifecycleContext` contain an
unforgeable callback-lifetime borrow of `CallCapability`. Only the lifecycle
context currently exposes calls (`xlGetName`, register, unregister); worksheet
contexts remain intentionally narrow, and thread-safe contexts do not expose
command or macro-sheet operations. `xlFree` is an internal release capability.

M11 adds verified `xlCoerce`, sheet helpers, and `xlAbort` to the appropriate
worksheet, thread-safe, and macro callback capabilities. `xlfCaller` is kept
out of `ThreadSafeContext`; it is available only to worksheet and macro
callbacks. The cancellation clear/preserve choice is explicit and must not be
described as a calculation-state API. Lifecycle capabilities do not expose it.

M16 does not make a background Excel context. `AsyncCancellationToken` is a
cooperative request-state view with no callback backend and no C API methods.
The private async controller owns the only `xlAsyncReturn` path, while
`LifecycleContext` privately registers the two calculation event procedures.
No controller capability crosses generations: scheduling clones one stable
`Arc`, shutdown removes that generation before joining, and reopen consumes a
fresh executor/controller. Neither the generation holder nor the controller
lifecycle mutex is held while calling Excel, running user code, or joining.

Only the non-default `xlcontime-research` feature exposes doc-hidden,
experimental `xlfNow` and `xlcOnTime` methods on `MacroContext`. The default
crate has no such methods or descriptors. Its bootstrap and close probes use an
`unsafe`, feature-gated lifecycle bridge whose caller must already be inside a
genuine synchronous Excel lifecycle callback on that callback thread; a linked
backend is insufficient. This does not promote arbitrary XLM commands or
`xlcOnTime` into the stable lifecycle capability. The registered callback must
demonstrate a real `MacroContext` operation and generation match during live
validation before any production context decision is made.
