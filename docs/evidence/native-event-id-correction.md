# #136: signed native event registration identifiers

The earlier diagnosis that `xlEventRegister` returned zero was incorrect. Its
error reported **non-positive**, which also included negative values. A raw x64
probe observed tag 2048 (`xltypeInt`), signed value `-853147637`, unsigned bits
`3441819659`. Registration order, short names, explicit module qualification,
trailing NUL storage, and an unoptimized build did not fix the rejection.
`xlGetName` identified the correct XLL. Both handler exports were present.

The correction preserves every nonzero i32 ID and rejects zero. Microsoft
[documents](https://learn.microsoft.com/en-us/office/client-developer/excel/xleventregister)
positive success IDs and zero failure; the sign-bit case is observed behavior on
Excel x64 16.0.20527.20150, not a claim that all hosts follow an undocumented
contract. A Microsoft SDK C++ control also returned a nontrivial integer ID.
No ABI, callback origin, registration ordering, or handler names were changed.
Regression coverage accepts 1, i32::MAX, i32::MIN and the observed negative ID,
and rejects zero.

With the correction alone, RUST.ADD returned 5 and RUST.PING.COMMAND returned 1
in isolated native sessions (`20261107T060415Z-85411420`). The acceptance feature
adds scratch-only functions/commands for reference reads, writes/clearing,
document information, recalculation, HelpTopic, and event counters. The default
registration surface remains unchanged.

Exploratory sweep `20261107T060842Z-df18c2ec` observed 12 PASS, one dispatcher
status assertion FAIL, two BLOCKED. The queue failure was cached worksheet
diagnostics: direct fresh invocation observed completion and processed=1 in
`20261107T061050Z-fb49eee6`. The runner now uses that fresh observation. The
exploratory runner changed during development; these runs are not a frozen
final acceptance baseline. Posted F9 did not deliver an observed event; that
probe remains BLOCKED, with no claim of failed production event delivery.

These are basic native probes, not acceptance of every outstanding edge case.
Cancellation/races, protected/stale/array-formula targets, full document selector
coverage, Function Wizard rendering, and same-process unload/retry still require
additional acceptance work under #126. Autonomous wake remains unimplemented.
x86 remains separate under #85. No CI or screenshot tools were used.
