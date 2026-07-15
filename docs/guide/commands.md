# Commands

```rust,no_run
use excel_api::prelude::*;

#[excel_command(name = "RUST.HELLO", thunk = "rust_hello")]
fn hello(_context: MacroContext<'_>) {}
```

Commands are macro-sheet callbacks, not worksheet functions. Define a command
with one `&MacroContext` parameter and return `()` or a supported
`Result<(), E>`.

```rust,no_run
use excel_api::{excel_command, MacroContext};

#[excel_command(name = "RUST.PING", thunk = "rust_ping", description = "Records a command.")]
fn ping(_context: &MacroContext<'_>) {}
```

The context exists only for the Excel-issued command callback. Commands cannot
be thread-safe, cannot use arbitrary worker-thread calls, and should perform
only calls whose catalogue legality is documented for `MacroContext`.
