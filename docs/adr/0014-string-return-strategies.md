# ADR-0014: String return strategies

Use `XLOPER12 | xlbitDLLFree` for general dynamic strings. Defer direct dynamic
simple-string returns.

## Status

Logical UTF-8/UTF-16 planning is implemented in M4 and stable counted UTF-16
backing buffers in M5. The `xlbitDLLFree` handoff remains M6.
