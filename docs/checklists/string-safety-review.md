# String Safety Review

- [ ] Correct parser for counted versus NUL-terminated ABI.
- [ ] Prefix/terminator excluded from payload.
- [ ] Length/scan bounded.
- [ ] Embedded NUL policy is explicit.
- [ ] Unpaired surrogate policy is explicit.
- [ ] No implicit lossy conversion.
- [ ] Dynamic return uses XLOPER12 DLLFree.
- [ ] Direct simple-string return is not assumed AutoFree-compatible.
- [ ] String array elements are deeply owned.
