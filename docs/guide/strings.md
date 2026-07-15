# Strings

`ExcelString` owns UTF-16 code units and preserves data that may not be valid
Unicode scalar text. Use `as_utf16` for exact units or the documented strict
and lossy conversion paths when converting to Rust text.

Excel's callback string representation can be counted or NUL-terminated
depending on the API contract. `ExcelStr<'call>`, `CountedUtf16Arg<'call>`, and
`NullTerminatedUtf16Arg<'call>` retain that distinction; none may outlive the
callback. Return strings through owned return planning so `xlAutoFree12`, not
Rust or `xlFree`, performs the correct DLL cleanup.
