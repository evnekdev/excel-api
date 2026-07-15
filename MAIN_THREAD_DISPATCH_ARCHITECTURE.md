# Main-Thread Dispatch Architecture

## Current decision

M17 has not selected an autonomous wake mechanism. A research-only
`xlcOnTime` compatibility probe exists, but the decision is **inconclusive**
until its complete contract, security behavior, cancellation, and unload
behavior are reproduced on a working current Excel host.

The production fallback remains a manually pumped cooperative dispatcher. The
research code does not create a dispatcher queue and is not a public arbitrary
XLM-command API.

## Historical evidence and verified boundary

Steve Dalton, sections 9.10.1 and 9.11.9, describes `xlcOnTime` as a polling
bridge that asks Excel to invoke a registered XLL command. Checked-in
`XLCALL.H` confirms only the IDs:

- `xlfNow = 74`;
- `xlcOnTime = 148 | xlCommand = 32916 = 0x8094`.

Microsoft's current C API documentation confirms that a registered XLL command
called by Excel is a class-3 context and can make command-equivalent calls. It
also states that Excel12v cannot be called from a background thread or an
operating-system timer callback. Current Microsoft documentation does not
publish the modern `xlcOnTime` argument/result/cancellation contract. The
similar VBA `Application.OnTime` contract is corroborating evidence, not proof
of the C API command.

## Experimental surface

The typed spike accepts only:

- two-argument schedule: Excel serial time and exact registered command name;
- schedule with a latest Excel serial time;
- four-argument cancellation: time, command name, missing latest time, FALSE;
- zero-argument `xlfNow` for Excel's own serial clock.

Every call keeps its counted command string and all XLOPER12 roots live through
Excel12v. It records both the raw C API return code and the immediate Boolean,
Excel-error, or unexpected result tag. Immediate scalar/error results create no
`ExcelOwnedValue` and no `xlFree` obligation.

The minimal XLL registers `RUST.ONTIME.SCHEDULE`,
`RUST.ONTIME.CALLBACK`, `RUST.ONTIME.CANCEL`, and bounded diagnostic helpers.
The callback receives `MacroContext`, checks the active runtime generation,
records process/thread/order/time information, and performs a harmless
preserving `xlAbort` poll. It is not connected to an M17 queue.

## Experimental lifecycle rule

Pending entries retain the exact scheduled serial, command, form, and runtime
generation. Test-mode bootstrap is enabled only by a coordination marker whose
PID matches the current Excel process. Close marks the generation inactive,
attempts cancellation while the backend remains linked, and only then permits
command unregistration and unlink. A failed cancellation makes the sample
close callback fail rather than intentionally proceeding with an unproved
pending callback.

This lifecycle behavior is an implementation to validate, not evidence that
Excel honors cancellation. Production approval requires a live proof that no
pending callback can enter unloaded code.

## Production acceptance gate

`xlcOnTime` may become the default only after the issue #30 matrix establishes:

- normal current security settings allow the registered-command callback;
- callback context and main-thread legality through more than thread identity;
- two/latest/cancel result contracts are reproducible;
- repeated scheduling remains bounded and coalescible;
- close/unload cancellation is reliable across reload and Excel shutdown;
- XLM macro policy is not a required weakening;
- editing, modal UI, calculation, copy/paste, undo, latency, and idle CPU are
  acceptable.

Until then, M17 must not use this mechanism autonomously.
