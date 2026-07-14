# ADR-0019: Owned semantic values preserve Excel distinctions

## Status

Accepted and implemented in M3.

## Decision

Owned semantic values contain Rust-owned data only. `ExcelString` stores
payload-only UTF-16, `ExcelArray` stores a flat immutable boxed row-major value
sequence, and `ExcelValue` preserves `xltypeInt` as `Integer(i32)` rather than
erasing it into `Number(f64)`. Missing and empty remain distinct.

Callback deep copies are preflighted with configurable string, element,
aggregate-byte, and depth limits before allocation. References are rejected
without coercion because workbook/sheet identity is context-dependent.

## Consequences

Owned values are naturally `Send + Sync + 'static`, can outlive callbacks, and
can be cached or moved to worker threads. Return-buffer layout and ownership
bits remain separate later milestones.
