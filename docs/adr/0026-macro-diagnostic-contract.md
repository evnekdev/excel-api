# ADR-0026: Macro diagnostic and generated-symbol contract

## Status

Accepted and implemented in M10.

## Decision

`#[excel_function]` rejects every unsupported signature or inconsistent
registration form during macro expansion. Checked `trybuild` snapshots make
the errors stable enough to remain actionable: they mark the invalid type,
attribute, context, flag, or signature and state the supported replacement
where one exists. Microsoft permits a cluster-safe function without the
thread-safe marker, so the macro does not invent that restriction. It rejects
the documented cluster-incompatible `U` reference family; a macro-sheet
function cannot be thread- or cluster-safe. Behavioral cluster safety cannot
be proved from a signature.

The macro's doc-hidden generated Rust names are deterministic implementation
details, not public semver API. The user-supplied `thunk` export remains the
public XLL ABI identity and must be changed only as a deliberate breaking
change. Duplicate generated Rust names fail in Rust compilation; duplicate
exports fail at link time.

## Consequences

Invalid workbook entry points fail before Excel executes. The shared closed
argument/result kind model remains the source of both registration metadata
and thunk ABI, so a diagnostic-only milestone does not add any runtime FFI,
ownership, callback, or threading behavior.
