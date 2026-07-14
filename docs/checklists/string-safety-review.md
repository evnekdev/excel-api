# String Safety Review

- [x] Correct parser for counted versus NUL-terminated ABI.
- [x] Prefix/terminator excluded from payload.
- [x] Length/scan bounded.
- [x] Embedded NUL policy is explicit.
- [x] Unpaired surrogate policy is explicit.
- [x] No implicit lossy conversion.
- [ ] Dynamic return uses XLOPER12 DLLFree.
- [x] Direct simple-string return is not assumed AutoFree-compatible.
- [x] String array elements are deeply owned.
