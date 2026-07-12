# Contributing (Architecture)

Rules:

- No Rust panic may cross FFI.
- Unsafe only in designated modules.
- ABI changes require ABI tests.
- Preserve crate layering.
- Document every public API.
- Prefer explicit ownership over hidden allocations.
