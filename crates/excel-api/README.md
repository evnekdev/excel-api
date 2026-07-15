# excel-api

Safe Rust building blocks for native 64-bit Microsoft Excel XLL add-ins.

The crate provides callback-borrowed and owned Excel values, return planning
and allocation, registration and procedural-macro integration, typed callback
contexts and Excel calls, lifecycle management, asynchronous one-shot UDFs,
and a bounded cooperative dispatcher.

The cooperative dispatcher does not wake Excel: queued work runs only when a
legal Excel-issued callback, such as the explicit pump command, drains it.
Async UDF and dispatcher lifecycle validation in real Excel remains a release
gate even though deterministic automated coverage is present.

RTD, Ribbon, general COM UI, custom task panes, autonomous notification, and
`xlcOnTime` research are not part of the stable core support promise.

Licensed under either Apache-2.0 or MIT.
