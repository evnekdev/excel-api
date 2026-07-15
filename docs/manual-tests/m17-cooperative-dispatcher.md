# M17 cooperative dispatcher manual validation

## Status

Pending. The available host cannot create a plain Excel workbook, so the pump
has not been exercised in real Excel.

## Procedure on a supported 64-bit Excel host

1. Build the default release XLL and confirm no `RUST.ONTIME.*` registration or
   export exists.
2. Load the XLL and invoke `RUST.DISPATCH.ENQUEUE`.
3. Read `RUST.DISPATCH.STATUS`; confirm the owned echo remains queued and no
   Excel call occurred on the background thread.
4. Invoke `RUST.DISPATCH.PUMP`; confirm the status becomes completed.
5. Enqueue more than one batch and confirm each pump processes at most the
   configured batch.
6. Exercise cancellation and close with pending work; confirm tickets retire
   and no operation occurs after unlink.
7. Reload, and confirm stale tickets from the old generation neither execute
   nor receive new-generation results.
8. Exercise a nested drain test path and confirm it is suppressed.

Record Excel architecture/version/build, exact commands, queue/batch settings,
registration and unlink ordering, cancellation and reload results, and any
diagnostics. Do not claim autonomous progress: enqueueing does not wake Excel.
