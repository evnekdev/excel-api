# Callback and Lifecycle Architecture

## Exported callbacks

- `xlAutoOpen`;
- `xlAutoClose`;
- `xlAutoAdd`;
- `xlAutoRemove`;
- `xlAddInManagerInfo12`;
- optional `xlAutoRegister12`;
- `xlAutoFree12`.

## Initialization state machine

```text
Uninitialized
 -> Initializing
 -> Initialized
 -> Closing
 -> Closed/Uninitialized
```

Initialization must be idempotent.

The book documents callback sequences where:

- Add-in Manager info may be requested before normal open initialization;
- add/open can occur without a preceding remove/close;
- close can be called during a shutdown attempt that is later cancelled.

Therefore:

- `xlAutoOpen` checks whether already initialized;
- `xlAddInManagerInfo12` and `xlAutoAdd` may ensure initialization;
- duplicate menus/registrations must be avoided;
- destructive shutdown must be conservative and version-aware.

## `xlAutoClose`

Responsibilities:

- unregister functions where reliable;
- stop/coordinate workers;
- release runtime-owned resources;
- unlink Excel call interface last.

Do not unload call pointers before destructors/releases that need them.

## `xlAutoFree12`

- allocation-free;
- panic-contained;
- no Excel calls in the initial ownership model;
- reconstruct and drop one DLL-owned return root.

## `DllMain`

No Excel C API calls. Avoid nontrivial initialization under loader lock.
