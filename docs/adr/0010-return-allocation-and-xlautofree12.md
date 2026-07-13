# ADR-0010: Return allocation and `xlAutoFree12`

- Status: Proposed

Each complex return is materialized into a self-describing top-level heap allocation with the root `XLOPER12` at offset zero and stable boxed backing buffers. Construction remains RAII-owned until a consuming handoff uses `Box::into_raw`. Excel later calls `xlAutoFree12`, which reconstructs the exact allocation with `Box::from_raw` and drops it once. DLL-free ownership is set only during handoff. No global registry is required for correctness.
