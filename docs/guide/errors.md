# Errors

```rust
use excel_api::ExcelError;

let result: Result<f64, ExcelError> = Err(ExcelError::Div0);
assert!(matches!(result, Err(ExcelError::Div0)));
```

Use `ExcelError` for controlled worksheet errors and return it (or an accepted
`Result<_, E>`) from a generated function. The thunk maps documented framework
errors to Excel error values and contains panics at the ABI boundary.

Do not return borrowed data in an error path. Conversion, return planning,
registration, call, lifecycle, async, and dispatcher errors are structured Rust
errors with `Display`, `Debug`, and `Error` implementations where applicable.
They distinguish Excel return codes, user-visible Excel errors, cancellation,
shutdown, and internal invariant failures.
