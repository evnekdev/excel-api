# Excel ABI checker

Compares selected `excel-api-sys` definitions with C values compiled from
Microsoft's `XLCALL.H`.

```powershell
cargo run --manifest-path tools/abi-check/Cargo.toml
```

Requires Windows, MSVC build tools, and the SDK. Override the SDK root with
`EXCEL_XLL_SDK_DIR`.

This is a scaffold. Prompt 01 must expand coverage to every supported layout,
constant, and function signature.
