# ADR-0010: XLL return allocation

Use a top-level per-call `ReturnAllocation` with offset-zero root, consuming
handoff, and exactly-once `xlAutoFree12`.

## Status

Implemented through M6. `ReturnPlan` selects the DLL-owned strategy, M5
materializes an offset-zero root plus stable locally owned string and array
backing storage, and M6 consumes that owner into a root pointer carrying only
`xlbitDLLFree`.

The ownership states are `Local`, `HandedOff`, and `Freed`. Allowed transitions
are local Drop or `Local -> HandedOff -> Freed`; handoff consumes the safe owner
and the second transition occurs only through the matching callback. Duplicate
callbacks after free are contract violations and cannot be made safe by
examining the already-freed pointer.

The callback casts the offset-zero root pointer to the exact
`ReturnAllocation`, reconstructs `Box<ReturnAllocation>`, and relies on normal
Rust field drop for all nested storage. It never reconstructs `Box<XLOPER12>`,
walks raw tags, calls `xlFree`, or calls Excel. The allocation layout belongs
to the producing binary and must be reclaimed by the callback in that same
loaded XLL.

The reusable panic-contained `unsafe extern "system"` body lives in
`excel-api`; `minimal-xll` owns the exact `xlAutoFree12` export and delegates to
it. This avoids an unconditional exported symbol in the reusable `rlib` while
preserving the SDK's `void WINAPI (LPXLOPER12)` ABI.
