# String Safety Review

- [x] Correct parser for counted versus NUL-terminated ABI.
- [x] Prefix/terminator excluded from payload.
- [x] Length/scan bounded.
- [x] Embedded NUL policy is explicit.
- [x] Unpaired surrogate policy is explicit.
- [x] No implicit lossy conversion.
- [x] Dynamic return uses root XLOPER12 DLLFree.
- [x] Direct simple-string return is not assumed AutoFree-compatible.
- [x] String array elements are deeply owned.
- [x] Nested strings retain base-only tags and are dropped through the top-level owner.
