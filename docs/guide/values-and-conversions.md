# Values and conversions

```rust
use excel_api::{ExcelValue, IntoExcel};

let value: ExcelValue = 7_i32.into_excel();
assert!(matches!(value, ExcelValue::Integer(7)));
```

`ExcelValueRef<'call>` is a callback-borrowed view of one `XLOPER12`. It cannot
escape the callback lifetime. `ExcelValue` is an owned semantic value with no
pointers into Excel-owned storage and may be moved to ordinary Rust code.

Use `FromExcel` to copy from a borrowed view and `IntoExcel` to produce an
owned value. Conversion limits bound UTF-16 length, array elements, aggregate
bytes, and nesting. Conversion errors are explicit; references are not silently
coerced to values.

`ExcelValueArg<'call>` maps a general `Q` argument. `ExcelReferenceArg<'call>`
maps `U` and preserves reference semantics. These names make the registration
contract visible in a Rust signature.
