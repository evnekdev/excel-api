# ADR-0023: Freeze the M8 handwritten implementation as the M9 oracle

## Status

Accepted.

## Decision

The manual M8 XLL is the conformance oracle for M9 macro-generated ABI glue.
Fixture tests cover each registration descriptor's symbol, Excel name, type
text, flags, argument metadata, return strategy, and error policy. Windows CI
also checks the compiled PE export table.

## Consequences

Macro work must match this observable contract before it can replace manual
thunks. The freeze does not expand the M8 call catalogue or change ownership:
dynamic returns remain DLLFree-owned, while raw XLFree return transfer remains
explicitly deferred.
