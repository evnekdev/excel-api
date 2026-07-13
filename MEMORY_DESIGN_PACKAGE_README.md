# Memory design package

Copy the files into the repository root while preserving paths.

- Add `MEMORY_OWNERSHIP_ARCHITECTURE.md` and `MEMORY_OWNERSHIP_ROADMAP.md`.
- Replace `docs/adr/0003-memory-model.md`.
- Add ADRs 0009-0011, three Mermaid diagrams, and the memory-safety checklist.

The key decisions are separate ownership-domain types, callback-lifetime borrowing, two-phase return planning/materialization, stable self-contained XLL returns, consuming handoff with `xlAutoFree12`, and a separate RAII path for Excel-owned API results.
