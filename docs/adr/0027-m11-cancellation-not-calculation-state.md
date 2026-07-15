# ADR-0027: Model `xlAbort` as cancellation polling only

## Status

Accepted for M11.

## Decision

The typed Excel C API catalogue exposes `xlAbort` only as cancellation polling.
An omitted argument or TRUE preserves a pending break; FALSE clears one. Its
immediate `xltypeBool` reports a user Esc/CANCEL request and has no Excel-owned
auxiliary allocation or `xlFree` obligation.

The catalogue intentionally does not expose Done/Calculating/Pending. Microsoft
documents `Application.CalculationState` through VBA/COM, not through a
verified `Excel12`/`Excel12v` callback. `xlretUncalced` remains a call return
code, not a calculation-state result. The project will revisit this only when
an authoritative C API function ID, selector, argument contract, result
ownership, and legal callback contexts are available.

## Consequences

No guessed XLM selector, placeholder API, or COM integration is added in M11.
The precise cancellation API is available only through callback capabilities
whose documented MTR legality is encoded by the catalogue.
