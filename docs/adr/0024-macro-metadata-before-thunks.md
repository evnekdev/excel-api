# ADR-0024: Generate typed macro metadata before ABI thunks

## Status

Accepted and implemented in M9A.

## Decision

`#[excel_function]` preserves the ordinary Rust function and generates only a
deterministically named `FunctionRegistration`. A closed syntactic type map
selects existing typed registration families, with explicit wrappers for Q,
U, counted UTF-16, and NUL-terminated UTF-16. Context references are omitted
from Excel-visible arguments.

The future thunk symbol is explicit metadata. No FFI export, unsafe callback
code, argument conversion, ownership handoff, or panic boundary is generated
until M9B.

## Consequences

M8 fixture comparisons can review metadata independently from ABI glue.
Unsupported or ambiguous signatures fail during macro expansion. Generated
metadata names are deterministic within their module, and Rust reports a
normal duplicate-item error if a user creates a colliding item.
