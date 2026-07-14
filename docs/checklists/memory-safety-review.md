# Memory Safety Review

- [x] Ownership domain is explicit.
- [x] Callback inputs are never freed or modified.
- [ ] Excel-owned result uses `xlFree` or consuming XLFree transfer.
- [ ] DLL-owned return uses DLLFree and AutoFree.
- [ ] No fallible work after handoff.
- [x] No pointer targets movable storage.
- [x] DLL-owned multis deeply own all nested strings.
- [ ] Excel-created multis are freed only at the top level.
- [x] Arrays-of-arrays/references are rejected.
- [x] Thread-safe function uses no static mutable return root.
- [ ] Panic paths leak nothing.
- [x] Injected partial-failure paths leak nothing before handoff.
- [x] Debug live count returns to zero before handoff.
