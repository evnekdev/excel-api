# Excel COM native ABI differential

This is an isolated research tool, not an Excel client library.  It records a
small, repeatable `IDispatch` differential without changing the workspace's
production dependencies.

`generate --root <knowledge-root>` creates deterministic source and operation
records.  `run-native`, `run-shim`, `run-high`, and `run-lower` append only
copied diagnostics to that root.  The fixture path is deliberately never
written to evidence; it is only passed to Excel for `Workbooks.Open`.

The lower Rust path uses `windows-sys` for COM activation and the SDK's generic
`IDispatch` ABI layout.  It does not contain an Excel interface vtable.
