# ADR-0010: XLL return allocation

Use a top-level per-call `ReturnAllocation` with offset-zero root, consuming
handoff, and exactly-once `xlAutoFree12`.

## Status

Partially implemented through M5. `ReturnPlan` selects the future DLL-owned
strategy and M5 materializes an offset-zero root plus stable locally owned
string and array backing storage. No ownership bit, consuming handoff, raw
reconstruction, or free callback is implemented yet.
