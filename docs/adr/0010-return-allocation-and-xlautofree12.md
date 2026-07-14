# ADR-0010: XLL return allocation

Use a top-level per-call `ReturnAllocation` with offset-zero root, consuming
handoff, and exactly-once `xlAutoFree12`.

## Status

Partially implemented in M4 as planning only. `ReturnPlan` selects the future
DLL-owned strategy and counts the root and all backing objects, but creates no
`ReturnAllocation`, raw pointer, ownership bit, handoff, or free callback.
