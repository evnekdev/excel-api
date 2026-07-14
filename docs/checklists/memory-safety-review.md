# Memory Safety Review

- [x] Ownership domain is explicit.
- [x] Callback inputs are never freed or modified.
- [ ] Excel-owned result uses `xlFree` or consuming XLFree transfer.
- [x] DLL-owned return uses DLLFree and AutoFree.
- [x] No fallible work after handoff.
- [x] No pointer targets movable storage.
- [x] DLL-owned multis deeply own all nested strings.
- [ ] Excel-created multis are freed only at the top level.
- [x] Arrays-of-arrays/references are rejected.
- [x] Thread-safe function uses no static mutable return root.
- [x] AutoFree panic boundary contains unwinding; the test-only pre-reclaim panic path leaks nothing.
- [x] Injected partial-failure paths leak nothing before handoff.
- [x] Debug live and outstanding-handoff counts return to zero after callback.
- [x] Callback reconstructs `Box<ReturnAllocation>`, never `Box<XLOPER12>`.
- [x] Root is offset zero and only the root carries DLLFree.
- [x] Duplicate callback is documented as an unrecoverable ownership-contract violation.
