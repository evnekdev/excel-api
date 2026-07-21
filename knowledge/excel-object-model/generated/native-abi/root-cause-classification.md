# Root-cause classification

**Inconclusive, with a narrowed negative finding:** the native C ABI shim and lower-level Rust generic `IDispatch` path succeed, while the full high-level Rust local/0x0400 harness fails. The minimal high-level reproduction succeeds on both current and preceding released windows-rs versions. This does not confirm a windows-rs regression. The final standalone C++ runner's conflicting local/0x0400 result prevents a clean Case D classification. Prompt 05 remains blocked pending a targeted production-harness repair and revalidation.
