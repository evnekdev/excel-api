# ADR-0003: Ownership domains

Use distinct types for borrowed inputs, Rust-owned values, Excel-owned C API
results, and XLL-owned returns.

## Status

Accepted. Callback-borrowed, Rust-owned semantic, logical return-plan, and
pre-handoff XLL-owned return-allocation domains are implemented through M5.
Excel-owned API results and post-handoff XLL return ownership remain pending.
