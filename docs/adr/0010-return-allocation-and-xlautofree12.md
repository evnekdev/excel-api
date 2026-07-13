# ADR-0010: XLL return allocation

Use a top-level per-call `ReturnAllocation` with offset-zero root, consuming
handoff, and exactly-once `xlAutoFree12`.
