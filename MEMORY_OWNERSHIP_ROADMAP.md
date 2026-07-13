# Memory and Ownership Roadmap

1. Verify ABI and ownership bits.
2. Borrow callback inputs.
3. Build owned semantic values.
4. Plan return trees without pointers.
5. Materialize stable DLL-owned storage.
6. Implement `xlbitDLLFree` handoff and `xlAutoFree12`.
7. Implement `ExcelOwnedValue` with `xlFree`.
8. Add consuming `xlbitXLFree` transfer.
9. Stress-test MTR, workbook close, and repeated recalculation.
10. Revisit TLS only if profiling shows per-call root allocation is a bottleneck.
