# Lifecycle

```rust,no_run
use excel_api::{AddInDescriptor, Runtime};

fn open(runtime: &Runtime, add_in: &AddInDescriptor) { let _ = runtime.initialize(add_in); }
fn close(runtime: &Runtime) { let _ = runtime.close(); }
```

`Runtime` coordinates registration and teardown. Initialize from a genuine XLL
open callback, provide the exact Excel callback entry point, and close before
unload. Registration is rollback-aware; close failures are represented rather
than being silently treated as a healthy initialized runtime.

Closing removes async and dispatcher generations before unregistering and
unlinking the backend. Do not retain contexts, raw Excel result pointers, or
background work across unload. A fresh open installs a new generation; stale
async or dispatcher work cannot publish into it.
