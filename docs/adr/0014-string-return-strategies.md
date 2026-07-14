# ADR-0014: String return strategies

Use `XLOPER12 | xlbitDLLFree` for general dynamic strings. Defer direct dynamic
simple-string returns.

## Status

Logical UTF-8/UTF-16 return planning is implemented in M4. Counted UTF-16
backing buffers and the `xlbitDLLFree` handoff remain M5-M6.
