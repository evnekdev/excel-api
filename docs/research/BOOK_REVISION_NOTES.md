# Book-driven Revision Notes

The uploaded book materially strengthened these areas:

- exact lifecycle callback ordering and the need for idempotent initialization;
- distinction between reference-preserving and value-only general arguments;
- separate treatment of `FP12`, mixed multis, references, missing, nil, and
  binary names;
- `xlFree` rules for Excel-created multis;
- timing of `xlbitXLFree`;
- per-call heap return roots versus TLS return slots;
- thread-safety constraints on static return objects;
- arrays-of-arrays and arrays containing references should not be returned;
- current versus active workbook/sheet/cell terminology;
- C API legality depends on function/command/macro context;
- interface code should remain separate from core calculation code.

The architecture deliberately adopts stricter ownership invariants than the
book's flexible mixed-ownership array example because Rust can make a
single-owner return tree safer and easier to audit.
