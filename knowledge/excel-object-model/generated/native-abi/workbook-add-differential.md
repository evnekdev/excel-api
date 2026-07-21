# Workbooks.Add differential

The native C ABI shim and windows-sys generic-IDispatch paths succeed in every required activation mode. The high-level windows harness fails only for `CoCreateInstance(CLSCTX_LOCAL_SERVER)` with LCID 0x0400; its isolated minimal high-level reproduction succeeds. This is not a crate-version regression. The final standalone C++ runner's conflicting local/0x0400 failure remains unresolved.
