# ADR-0009: Borrowed input lifetimes

- Status: Proposed

Excel callback pointers are represented by `ExcelValueRef<'call>` and related borrowed views. Only thunk/runtime code may construct them. Safe code cannot retain them beyond the callback; async and worker-thread use requires a deep copy to an owned Rust type. Borrowed views are not `Send` or `Sync` by default.
