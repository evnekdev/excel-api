# ADR-0003: Memory and ownership domains

- Status: Proposed replacement
- Supersedes: the initial one-line ADR-0003

## Context

Excel values cross FFI in incompatible ownership states: borrowed callback arguments, normal Rust-owned values, Excel-owned API results, and XLL-owned returns later released through `xlAutoFree12`.

## Decision

Use separate types and explicit transitions:

- `ExcelValueRef<'call>` for callback-borrowed inputs;
- `ExcelValue` for ordinary owned Rust data;
- `ExcelOwnedValue` for RAII-managed Excel API results;
- `ExcelReturnValue` for logical returns;
- `ExcelReturn` for materialized XLL-owned returns before handoff;
- raw pointers only at final FFI boundaries.

Complex returns use a self-contained top-level allocation with the root `XLOPER12` at offset zero and stable backing buffers. Handoff consumes ownership; `xlAutoFree12` reconstructs and drops the exact allocation. Production correctness does not depend on a global pointer registry.

## Consequences

This adds explicit conversion stages and more types, but aligns lifetimes and destructors with the actual ABI, makes partial failure RAII-safe, and isolates unsafe code.

## Constraints

- no panic crosses FFI;
- no raw pointer is published before backing storage is stable;
- callback inputs are never freed;
- Excel API results never use `xlAutoFree12`;
- DLL-free ownership is applied only at final handoff;
- cleanup only accepts allocations created by this library.
