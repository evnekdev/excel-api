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

## C API calls

Thread-safe contexts expose only verified thread-safe operations.

The book identifies C API-only operations that are generally thread-safe, while
many worksheet/XLM calls are not. The safe wrapper must encode a whitelist,
not an optimistic default.

## XLL-created threads

A background thread must not call the Excel C API or COM directly unless a
specific documented mechanism permits it.

Worker threads receive owned Rust data only.

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
