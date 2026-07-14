# ADR-0003: Ownership domains

Use distinct types for borrowed inputs, Rust-owned values, Excel-owned C API
results, and XLL-owned returns.

## Status

Accepted. Callback-borrowed and Rust-owned semantic domains are implemented in
M2-M3. Excel-owned API results and XLL-owned return allocations remain pending.
