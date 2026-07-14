# Excel ABI checker

Compares `excel-api-sys` definitions with C values compiled from Microsoft's
`XLCALL.H` on Windows x86_64 MSVC.

```powershell
cargo run --manifest-path tools/abi-check/Cargo.toml
```

The checker covers primitive widths, complete structure/union sizes and
alignments, every nested field offset, supported constants and function IDs,
and compile-time compatibility of the `Excel12` and `Excel12v` prototypes. It
also verifies Microsoft-documented Excel 12 limits and registration codes.

Requires Windows x86_64, MSVC build tools, and the SDK. Override the SDK root
with `EXCEL_XLL_SDK_DIR`.
