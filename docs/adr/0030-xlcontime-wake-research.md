# ADR-0030: xlcOnTime wake research is inconclusive

## Status

Inconclusive.

## Context

Issue #30 originally found no verified modern main-thread wake mechanism for
M17. Steve Dalton's historical examples identify `xlcOnTime` as a possible
bridge from background state to an Excel-invoked registered XLL command.

Checked-in `XLCALL.H` proves `xlcOnTime = 148 | xlCommand` and `xlfNow = 74`.
Microsoft currently documents the general Excel12v callback legality classes,
registered XLL commands, the serial result of NOW, the analogous VBA
`Application.OnTime` arguments, and the separate Excel 4.0 macro security
setting. It does not provide a current C API reference for the complete
`xlcOnTime` contract.

## Decision

Do not approve `xlcOnTime` as M17's default or opt-in wake mechanism yet. Keep
the typed wrapper and registered commands internal, experimental, and behind
the non-default `xlcontime-research` feature. The default minimal XLL and normal
context API retain their pre-spike surfaces. Do not connect the experiment to a
production dispatcher queue.

The spike preserves the historical 2/3/4-argument forms, exact raw return code,
and immediate XLOPER12 result. It uses `xlfNow`, not Unix conversion, and keeps
the command string and argument roots alive for the synchronous call. Pending
callbacks are generation-tagged and close attempts cancellation before
unregistration and backend unlink.

The research lifecycle bridge is unsafe. Its caller must be synchronously
inside a genuine Excel-issued lifecycle callback on the callback thread; linked
backend state does not establish that fact. When the PID coordination marker is
absent, ordinary XLL open performs no experiment. When present, bootstrap writes
an explicit attempted/succeeded/failure record. `xlAutoOpen` still reports the
already-successful ordinary runtime initialization, so the harness must inspect
the artifact and cannot treat `RegisterXLL` returning TRUE as research success.

If pending cancellation fails, the experiment records the failure and declines
ordinary runtime cleanup. No authoritative contract establishes that returning
0 from `xlAutoClose` prevents Excel from unloading the DLL. A pending callback
therefore remains a potential unload hazard. The harness must cancel before
unload and, if it cannot prove cancellation, classify the run unsafe and
terminate only its isolated PID/start-time-matched Excel process.

The available Microsoft 365 64-bit host reported Excel 16.0 build 20131, but
plain `Workbooks.Add` failed before XLL loading with the known host document
creation error. `RegisterXLL` returned TRUE, yet without a workbook Excel did
not enter the isolated test bootstrap, and COM `Application.Run` rejected the
registered diagnostic command. `ExecuteExcel4Macro` also produced no command
invocation. Therefore no timed callback, command context, cancellation, or
unload claim can be made from this host.

## Consequences

- M17 remains unimplemented and the manually pumped cooperative design remains
  the safe fallback.
- The automated Rust/ABI/export checks establish wrapper shape, not Excel
  behavior.
- Ordinary builds contain neither the experimental scheduling API nor the
  `RUST.ONTIME.*` registrations/exports.
- The experiment cannot use an `xlAutoClose` return value as an unload safety
  mechanism.
- The live matrix in `MAIN_THREAD_DISPATCH_ARCHITECTURE.md` remains mandatory.
- If a working host proves the acceptance gate, a later reviewed ADR may change
  the outcome to accepted-default or accepted-experimental.
- If current security requires enabling legacy XLM macros, cancellation cannot
  be made reliable, or user interaction is materially damaged, reject the wake
  mechanism.

## Sources

- Steve Dalton, *Financial Applications Using Excel Add-in Development in
  C/C++*, sections 9.10.1 and 9.11.9 (historical secondary evidence).
- Microsoft Learn: Calling into Excel from the DLL or XLL.
- Microsoft Learn: Excel4/Excel12 and Excel4v/Excel12v.
- Microsoft Learn: Accessing XLL code in Excel.
- Microsoft Learn: Excel `Application.OnTime` (analogous VBA/COM contract).
- Microsoft Support: NOW function and Excel 4.0 macro security settings.
