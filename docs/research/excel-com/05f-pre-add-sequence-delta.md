# Prompt 05F: pre-`Add` delta and session-state investigation

## Outcome

No individual pre-`Workbooks.Add` operation was established as causal. The
result is **high-level runtime/path sensitivity, inconclusive**. The committed
fresh-process prefix matrix passed, while the separate recovery path and later
independent high-level controls failed before workbook creation with
`DISP_E_EXCEPTION` (`0x80020009`) and Excel `scode` `0x800A03EC`.

No production client or Range code was changed. Prompt 05 Range work remains
blocked because the required Rust-owned workbook and `A1.Value2` smoke test did
not complete in the failing path.

## Protocol and cold baseline

The work began from `origin/master` commit
`bbf767d522e968969cd5afd143dfcc80718c57cf` on branch
`research/excel-com-05f-pre-add-delta`. The cold-session controls were run
only after reboot, with no `EXCEL.EXE` process, no Excel opened since reboot,
and a separately recorded system boot and test-start time. Raw process IDs,
HWNDs, pointers, local paths, and account identifiers are deliberately not
committed.

The cold control order was minimal high-level, full high-level local/`0x0400`,
lower-level `windows-sys`, Rust-to-native C ABI shim, then native direct. All
created a workbook with `S_OK`; every owned process exited without forced
termination. This is a **cold-session success**. It makes an external/session
condition credible, but it is not a repair or a causal claim.

The user then opened and normally closed the same macro-enabled workbook used
in their normal workflow. The initial clean warm-session controls still passed.
The session-state record deliberately names only that normalized context, not
the workbook path or its contents.

## Exact delta sequence

The primary prefixes use the production-harness order and add exactly one
operation per row:

| Prefix | Added operation |
| --- | --- |
| A0 | baseline: activate, get `Workbooks`, `Add`, close, `Quit` |
| A1 | `Application.Version` |
| A2 | `Workbooks.GetTypeInfoCount` |
| A3 | `Workbooks.QueryInterface(IUnknown)` |
| A4 | `Workbooks.QueryInterface(IDispatch)` |
| A5 | `Workbooks.Count` |
| A6 | clone then clear lifetime transition |
| A7 | retain then clear lifetime transition |
| A8 | query-interface then clear lifetime transition |

Every live row used an STA child process, one owned Excel Automation instance,
and one of three activation modes: local server with `0x0400`, server with
`0x0000`, or `CoCreateInstanceEx` server with `0x0000`. The parent required
zero pre-existing Excel processes before every child, a confirmed owned-process
exit after it, and zero remaining Excel processes before proceeding.

The primary matrix contained 375 passing `Add` rows: 135 prefixes, 90 ownership
rows, 30 storage rows, 96 property-read rows, 12 type-information rows, and 12
process-instrumentation rows. A separately seeded 27-row A0-A8 pass across all
three modes also passed. Thus the stored prefix evidence contains 402 passing
rows and no first failing prefix. In particular, the 27-row pass includes A8 in
all modes after earlier recovery-path failures.

## State transitions and independent controls

The clean warm controls after the manual workbook close, a controlled
Automation warm-up, and a deliberately retained-then-released `Workbooks`
reference all passed. Later, two recovery attempts failed before `Open` in L,
S, and X. Immediate isolated A8 controls also failed in all three modes; a
subsequent reduction showed A0 failing in local/`0x0400`. This means the later
failure cannot be attributed to an optional pre-`Add` operation.

| Current local/`0x0400` control | `Add` | Cleanup |
| --- | --- | --- |
| Minimal high-level windows repro | `0x80020009` | not separately attributed |
| 05F A0 baseline | `0x80020009` | owned process exited |
| `windows-sys` generic `IDispatch` | `0x00000000` | exited after bounded wait |
| Rust to native C ABI shim | `0x00000000` | owned process exited |
| Native direct executable | `0x00000000` | owned process exited |

The native direct and shim controls also completed ten successful cold-context
local/`0x0400` repetitions each. Their current success does not resolve the
separate Prompt 05E host-context discrepancy; it only shows that the current
high-level failure is not a general Excel server or generic-dispatch failure.

## Recovery and Range status

The recovery command runs A8, then only if `Add` succeeds opens a known-good
temporary `.xlsx`, writes numeric `42` to `A1.Value2`, reads back the exact
`VT_I4` value `42`, clears the cell, closes the workbook, and quits Excel. The
read-back was strengthened to verify the value as well as its VARTYPE.

Both recorded recovery attempts failed at `Add` in every activation mode, so
`Open` and the Range smoke test were correctly not run. No repair was applied;
there is no successful Range claim and no resumption of Prompt 05 runtime
semantics work.

## Type-library validation

The ordinary `excel-com-typelib-audit check` is intentionally stale for the
historical manifest because it substitutes `not-recorded` for host labels. The
new read-only `check-historical` mode reads the committed historical labels,
checks their manifest structure, and compares every type-library-derived
artifact against two fresh inspections of the current registered Excel type
library. It does not overwrite the 05B-05E artifacts. It passed.

## Evidence and status

- [delta evidence](../../../knowledge/excel-object-model/pre-add-delta/)
- [prefix matrix](../../../knowledge/excel-object-model/generated/pre-add-delta/prefix-results.md)
- [session transitions](../../../knowledge/excel-object-model/generated/pre-add-delta/excel-session-state-transition.md)
- [current independent controls](../../../knowledge/excel-object-model/generated/pre-add-delta/current-state-controls.md)
- [recovery validation](../../../knowledge/excel-object-model/generated/pre-add-delta/repair-validation.md)
- [root-cause classification](../../../knowledge/excel-object-model/generated/pre-add-delta/root-cause.md)

The next safe step is not a production repair. It is a smaller high-level
binary/control-flow reproduction that preserves the current A0 failure while
changing one implementation factor at a time. Any future repair must first
produce successful `Add`, `Open`, and exact `A1.Value2 = 42` smoke rows in all
required modes.
