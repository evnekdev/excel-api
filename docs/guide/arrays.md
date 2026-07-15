# Arrays and dynamic arrays

```rust,no_run
use excel_api::{ExcelArray, ExcelValue};

let values = vec![ExcelValue::Number(1.0), ExcelValue::Number(2.0)].into_boxed_slice();
let array = ExcelArray::new(1, 2, values)?;
assert_eq!(array.dimensions(), (1, 2));
# Ok::<(), excel_api::OwnedValueError>(())
```

`ExcelArrayView<'call>` borrows a rectangular Excel array during a callback.
`ExcelArray` owns a row-major rectangular collection of `ExcelValue`s. Construct
an owned array only with dimensions matching the supplied element count; errors
reject malformed shapes and nested arrays that Excel's supported model cannot
represent.

Returning `ExcelArray` or `ExcelReturnArray` plans stable DLL storage for a
dynamic-array result. The plan is pointer-free until the thunk materializes it;
the materialized allocation belongs to the XLL and is reclaimed by
`xlAutoFree12`.
