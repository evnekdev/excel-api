# ADR-0014: String return strategies

Use `XLOPER12 | xlbitDLLFree` for general dynamic strings. Defer direct dynamic
simple-string returns.

## Status

Logical UTF-8/UTF-16 planning is implemented in M4, stable counted UTF-16
backing buffers in M5, and root-only `xlbitDLLFree` handoff with exact
owner-driven callback reclamation in M6. Nested strings carry base type bits
only and are never freed by traversing their raw pointers.
