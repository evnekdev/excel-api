# Initialization comparison

The faithful C++ and Rust controls use `CoInitialize(NULL)`. A five-run `CoInitializeEx(COINIT_APARTMENTTHREADED)` differential also completed and naturally exited. It is recorded only as a post-baseline comparison, not as a source-derived replacement.
