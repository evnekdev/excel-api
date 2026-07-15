# M9 manual implementation oracle

The M8 `examples/minimal-xll` implementation is the normative compatibility
oracle for M9 generated metadata and thunks. It is frozen by the fixture tests
in that crate and by Windows PE-export inspection in CI.

M9 output must preserve these properties unless a reviewed ADR deliberately
supersedes this policy:

- every exported thunk uses the verified `extern "system"` Excel12 ABI and
  catches unwinding panics before they cross the FFI boundary;
- descriptor type text is derived from the declared signature: `QBB$`, `QQ$`,
  `QQ$`, `QU`, and `QQ$` for the current functions, with `Q` value-only and
  `U` reference-preserving;
- `RUST.ADD`, `RUST.ECHO`, `RUST.ARRAY.ECHO`, and `RUST.OPTION.KIND` are
  thread-safe; `RUST.REFERENCE.KIND` is not;
- each successful dynamic result is fresh XLL-owned `DllOwnedXloper12` storage,
  marked only with `xlbitDLLFree` at handoff and reclaimed only by the matching
  `xlAutoFree12` callback;
- errors map to immutable scalar roots: invalid input and panic to `#VALUE!`,
  unsupported references to `#REF!`, numeric domain/range to `#NUM!`, and
  allocation/unavailability to `#N/A` where applicable;
- lifecycle exports remain idempotent, no runtime lock is held while calling
  Excel, and callback-scoped Excel-owned results are released before unlinking.

The oracle does not bless raw `xlbitXLFree` return transfer. That behavior is
intentionally deferred because the unique root lifetime has no documented
per-call reclaim callback.

## M9B parity result

The minimal XLL now uses generated worksheet thunks in production. Test-only
handwritten descriptors and behavioral fixtures remain as the oracle. Tests
cover all five symbols, exact type text, scalar/text/array/reference/optional
values, error and panic fallbacks, root and nested tags, Q/U distinction,
fresh DLLFree-only roots, and exact AutoFree reclamation. Windows release PE
inspection reports exactly the same 12 exports as the frozen M8 artifact.
