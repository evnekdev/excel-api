# ADR-0025: Generate worksheet thunks from one typed ABI model

## Status

Accepted and implemented in M9B. Automated parity, release build, PE export
inspection, and the two-process real Excel rerun pass.

## Decision

Extend the M9A closed syntactic type model so each supported Rust argument and
result selects both its `FunctionSignature` family and its exact Windows x64
`extern "system"` ABI type. Microsoft's xlfRegister table is authoritative:
`A` is a 16-bit signed Boolean, `B` is `double`, `J` is a 32-bit signed integer,
`C%`/`D%` are UTF-16 pointers, and `Q`/`U` are `XLOPER12*`.

Generated code contains only the exact exported function and calls into
`excel-api::thunk`. A callback-scope token centralizes unsafe Q/U and direct-
string borrowing, ties every view and injected context to one invocation, and
uses the same shared production callback backend as `Runtime::production`.
The helper layer owns conversion, panic containment, error policy, return
planning, materialization, and consuming DLLFree handoff.

Q-returning successes allocate one fresh `ReturnAllocation`; only its root
receives `xlbitDLLFree`, and the existing matching `xlAutoFree12` reclaims it.
All supported Excel-visible errors have immutable pointer-free fallback roots.
Direct scalar returns cannot encode an Excel error and therefore return the
documented zero/false fallback after conversion failure, ordinary
`Result::Err`, or panic.

The user-provided thunk name is validated as an ASCII x64 export identifier
and emitted exactly. Generated Rust item names are deterministic; duplicate
exact exports fail during linking.

## Consequences

Registration text and thunk signatures share one source of truth. Macro
expansion cannot retain callback memory, perform return allocation policy, or
let an unwinding panic cross the FFI boundary. The minimal XLL can replace its
five handwritten production thunks while keeping their descriptors and
behavioral fixtures as the M9 oracle.

Production runtimes in one loaded binary share the SDK callback backend so
generated contexts observe the entry installed by lifecycle initialization.
The intended architecture remains one lifecycle runtime per XLL; broader
multi-runtime coordination is deferred.

Raw XLFree return transfer, direct dynamic simple-string returns, commands,
async functions, dispatcher work, and expanded Excel call capabilities remain
out of scope.

Official reference: [xlfRegister (Form 1)](https://learn.microsoft.com/en-au/office/client-developer/excel/xlfregister-form-1).
