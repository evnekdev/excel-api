# x64 cancellation and same-process lifecycle acceptance

Owner #126; related #134, #124. Acceptance-only fixture source is the review
working tree based on fabda4d, adding RUST.TEST.DELAY and RUST.TEST.ASYNC.STATUS.
The actual tested XLL is acceptance-cancellation.xll, SHA-256
f0bac026cf8621df647565f9a4eb9dbeb98fe143d1a49a16b9eaf2a7362f4aa7.
These results predate the subsequent #138 Wizard correction.

## Cooperative cancellation

Run 20261107T070047Z-8a8b0ade: PASS. Operator input was F9, then Escape during
bounded asynchronous work (maximum ten seconds). Initial diagnostic tuple
(pending,cancel_events,end_events,active_workers,canceled_workers,completed)
was (0,0,1,0,0,0); final tuple was (0,2,3,0,2,0). Both host cancellation and
cooperative worker cancellation were observed; no request or worker remained.
The owned PID 33580 exited after normal cleanup. The maintainer observed #CALC;
that visual error alone was not used as the acceptance assertion.

The earlier run 20261107T065914Z-a635cb4a was BLOCKED by transient COM
RPC_E_CALL_REJECTED. No explicit Quit was requested while work was uncertain.
A subsequent process inventory showed that owned PID had exited. The poll was
corrected to retry only Excel busy and RPC call-rejected errors within its
existing deadline. This native pass does not establish every completion race,
materialization interleaving or pending-work unload; deterministic Rust race
tests remain separate evidence.

## Idle same-process deactivation and registration retry

Run 20261107T070447Z-f0a7fed5: PASS, owned PID 22068. Pending/active counts were
zero before UNREGISTER(module-path). Excel returned True; RUST.ADD then returned
Excel error -2146826273 instead of 5. Two subsequent RegisterXLL calls restored
RUST.ADD(2,3)=5 and RUST.PING.COMMAND. The HWND PID remained unchanged.
The scratch instance exited after cleanup. This validates host deactivation
and reinitialization, not a physical OS module-unmapping audit or unsafe unload
while work is outstanding.

Earlier run 20261107T070356Z-f3c1592d reported FAIL because the fixture treated
any non-exception Application.Run result as callable. Recording and checking
the returned Excel error corrected that assertion; the historical FAIL remains.

Authoritative host operation:
https://learn.microsoft.com/en-us/office/client-developer/excel/xlfunregister-form-2

All runs preserved pre-existing PID 19684. No Office settings, CI, x86 execution,
forced process termination or screenshot automation was used. #137 remains
unresolved and its stale crash probe was not repeated.
