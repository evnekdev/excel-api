# Book-driven Revision Notes

The uploaded second edition of Steve Dalton's book strengthened the design for:

- idempotent lifecycle initialization;
- reference-preserving versus value-only arguments;
- `FP12`, multis, references, missing, and nil;
- `xlFree` and `xlbitXLFree` timing;
- thread-safe return storage;
- current versus active Excel context;
- C API legality by callback context;
- separation of interface glue from calculation logic.

The Rust architecture intentionally uses stricter single-owner return trees than
flexible mixed-ownership C examples.
