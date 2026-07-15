# Optional, missing, empty, and nil

These are distinct Excel concepts:

| Form | Meaning |
| --- | --- |
| `Missing` | An omitted argument. |
| `Empty` | An empty Excel value. |
| `Nil` | Excel's distinct nil value in a borrowed callback representation. |
| `OptionalValue<T>` | Owned policy-preserving representation of missing, empty, or a value. |

Rust `Option<T>` follows the documented conversion policy; it is not a lossless
substitute for all three Excel states. Return the specific `ExcelValue` or
`ExcelReturnValue` variant when a worksheet contract must preserve the exact
Excel distinction.
