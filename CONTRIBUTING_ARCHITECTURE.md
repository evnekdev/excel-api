# Contributing Architecture Rules

## Unsafe code

Unsafe code is allowed only in clearly named internal modules.

Every unsafe function must document:

- pointer validity;
- alignment;
- readable/writable extent;
- union-tag correspondence;
- lifetime;
- ownership domain;
- thread/callback context;
- permitted caller.

## FFI

- All exports use the correct Windows calling convention.
- All exports catch panics.
- No destructor used at the FFI boundary may panic.
- No fallible work occurs after ownership handoff.

## Memory

- Callback arguments are never freed or modified.
- Excel API results use `ExcelOwnedValue`.
- XLL return allocations use `ExcelReturn`.
- No Excel-owned pointer may be embedded in a DLL-owned return tree unless the
  design explicitly supports and tracks mixed ownership. The initial design
  always deep-copies.
- No static mutable return root in a thread-safe UDF.

## Registration

- Type text must come from verified type mappings.
- Thread-safe and macro-sheet flags may not be combined.
- Reference-preserving and value-only registration forms must remain distinct.
- Unsupported signatures fail at compile time once macros exist.

## Testing

Changes touching ABI or memory require:

- unit tests;
- failure-path tests;
- integration plan;
- updated architecture/ADR if invariants change.

## Documentation drift

Any PR that changes a documented invariant must update:

1. the architecture document;
2. the affected ADR;
3. the architecture index;
4. unexecuted Codex prompts that depend on it.
