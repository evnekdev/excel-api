# Cooperative dispatcher

**Preview — live pump validation pending.** The dispatcher transports owned,
closed-catalogue operations to a later legal Excel callback. It has bounded
pending capacity and batch size, tickets, generation isolation, capability
matching, cancellation, cooperative expiry, and panic-safe retirement.

**Enqueueing a request does not wake Excel.** A request remains queued until a
user invokes `RUST.DISPATCH.PUMP` or another explicitly approved Excel-issued
callback drains compatible work. There are no timers, hidden windows,
`PostMessage`, COM, RTD, or `xlcOnTime` wake mechanisms in the core dispatcher.
Do not wait on a ticket from an Excel callback thread.

```rust,no_run
use excel_api::{dispatcher, DispatchOperation, ExcelValue};

let ticket = dispatcher::enqueue(DispatchOperation::EchoOwned(ExcelValue::Number(42.0)))?;
assert!(ticket.try_result().is_none()); // enqueue does not wake Excel
// A user later invokes the registered RUST.DISPATCH.PUMP command in Excel.
# Ok::<(), excel_api::DispatchEnqueueError>(())
```

Only a genuine compatible Excel callback may drain it; a background thread may
enqueue and poll its owned ticket but cannot manufacture a context.
