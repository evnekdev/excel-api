# Native calculation-ended delivery: manual F9

Owner #126; follows #136 registration correction. Run
`20261107T065550Z-0306963c` passed EVENTS on the existing x64 host and
acceptance-extended.xll (SHA-256
`f02edb333a1cd35901e502b0211fbfbd91a089b857be40741c9ece49b1cb53f0`).
The runner used manual calculation, dirtied a RUST.ADD dependency, and waited
for the maintainer to focus the isolated scratch instance and press F9.
The maintainer explicitly confirmed F9 worked; Fn+F9 did not. The output
folder name mentioning fn-f9 describes the attempted instruction, not the
actual key used for acceptance.

Observed: C1 changed from 21 to 22 and the calculation-ended diagnostic
counter increased from 2 to 3. This independently verifies callback delivery
for a user-triggered calculation. PID 4044 was absent from the initial PID
inventory, owned through its HWND, and exited after its scratch workbook was
closed and Quit requested. Pre-existing PID 19684 was untouched.

Earlier injected-key and manual attempts remained BLOCKED; they did not
recalculate the dirty formula. One immediate automation query returned
Excel busy 0x800AC472; the revised poll tolerates only that specific transient
error within a deadline. Desktop visibility was uncertain in earlier attempts.
No causal claim that Fn-lock caused those failures is supported.

Microsoft documents that these events are not raised during programmatic
recalculation: https://learn.microsoft.com/en-us/office/client-developer/excel/handling-events

This is not cancellation-race, same-process unload/retry, or Wizard metadata
acceptance. Those remain outstanding. #137 is also unresolved; the stale
reference crash probe was not repeated. No Office modifications, CI, screen
capture, x86 execution, or unrelated Excel cleanup occurred.
