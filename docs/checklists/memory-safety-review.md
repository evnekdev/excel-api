# Memory Safety Review

- [ ] Ownership domain is explicit.
- [ ] Callback inputs are never freed or modified.
- [ ] Excel-owned result uses `xlFree` or consuming XLFree transfer.
- [ ] DLL-owned return uses DLLFree and AutoFree.
- [ ] No fallible work after handoff.
- [ ] No pointer targets movable storage.
- [ ] DLL-owned multis deeply own all nested strings.
- [ ] Excel-created multis are freed only at the top level.
- [ ] Arrays-of-arrays/references are rejected.
- [ ] Thread-safe function uses no static mutable return root.
- [ ] Panic paths leak nothing.
- [ ] Debug live count returns to zero.
