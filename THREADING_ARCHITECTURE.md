# Threading Architecture

## Excel threading domains

- Excel UI/main thread;
- Excel multi-threaded recalculation workers;
- XLL-created worker threads;
- future async completion/dispatcher thread interactions.

## Thread-safe UDF rules

A function marked thread-safe:

- cannot also be macro-sheet equivalent;
- must not use thread-unsafe C API operations;
- must not use shared mutable static return storage;
- must protect shared resources;
- must not assume callback order.

## Return memory

Initial strategy:

- allocate one return root per call;
- free through `xlAutoFree12`;
- avoid TLS return slots.

This is easier to reason about than persistent thread-local return objects.

M6 implements this strategy without thread-local ownership. The handed-off
allocation contains `Box<[XLOPER12]>`, `Box<[XCHAR]>`, ordinary scalar
metadata, and test-only atomic tracking. These fields use the process allocator
and have no thread affinity, locks, COM state, workbook state, or creator-thread
dependency, so the matching callback may drop them on a thread different from
construction. A cross-thread reclamation test exercises that contract.

`ExcelReturn` is not given manual `Send` or `Sync` implementations. The
required boundary is the documented ownership transfer of its unique raw
pointer to Excel, not a blanket promise that the local wrapper is shareable.
`xlAutoFree12` makes no Excel calls and acquires no project locks.

## C API calls

Thread-safe contexts expose only verified thread-safe operations.

The book identifies C API-only operations that are generally thread-safe, while
many worksheet/XLM calls are not. The safe wrapper must encode a whitelist,
not an optimistic default.

## XLL-created threads

A background thread must not call the Excel C API or COM directly unless a
specific documented mechanism permits it.

Worker threads receive owned Rust data only.

M7's `ExcelOwnedValue<'call>` is deliberately neither `Send` nor `Sync`. Its
boxed root is stable, but the ability to call back into Excel is scoped to the
originating callback. Microsoft documents `xlFree` as thread-safe during MTR;
that does not make calls from arbitrary XLL-created threads legal. Convert to
ordinary `ExcelValue` before worker-thread use.

## Shutdown

Lifecycle shutdown must account for:

- outstanding workers;
- pending callbacks;
- Excel closing/cancelling;
- late cleanup calls;
- cancelled close sequences.

## Shared state

Prefer:

- immutable descriptors;
- atomics for counters/state;
- scoped locks;
- no lock in `xlAutoFree12` if avoidable.
## M8 classifications

`RUST.ADD`, `RUST.ECHO`, `RUST.ARRAY.ECHO`, and `RUST.OPTION.KIND` are pure and
registered `$`; each produces fresh per-call DLLFree storage. The
reference-preserving probe is deliberately not `$`. Registration runs only in
lifecycle context. Panic and catastrophic static scalar fallbacks are
immutable, pointer-free, and contain no ownership bit.
