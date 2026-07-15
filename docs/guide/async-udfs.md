# Async UDFs

**Preview — live lifecycle validation pending.** Async UDFs accept only owned,
`Send + 'static` inputs, run user work outside Excel, and publish at most once
through the documented async return path. They use bounded capacity, permanent
per-open executor generations, cancellation tokens, and event handlers
registered at most once per loaded binary.

An executor's shutdown is irreversible for its generation: after shutdown
returns, no accepted job may execute XLL code. Install a fresh executor before
reopen. Real-Excel validation is still required for cancellation,
recalculation-replacement, unload, reopen, MTR initiation, and capacity
rejection.

```rust,no_run
use excel_api::prelude::*;

#[excel_function(name = "RUST.ASYNC.DOUBLE", thunk = "rust_async_double")]
fn async_double(value: f64, cancellation: AsyncCancellationToken) -> Result<f64, ExcelError> {
    if cancellation.is_cancellation_requested() { return Err(ExcelError::Na); }
    Ok(value * 2.0)
}
```

This preview example permits owned background work, but not arbitrary Excel C
API calls. The runtime alone uses the narrow `xlAsyncReturn` completion exception.
