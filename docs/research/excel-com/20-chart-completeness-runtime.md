# Prompt 20: chart completeness and Excel runtime sessions

## Baseline and runtime blocker

The branch starts at `1e471b85f117ea589c52a14f7c702787fca25392`. The single
controlled `Workbooks.Add` baseline activated an owned invisible Excel server,
then failed with outer HRESULT `0x80020009` and Excel SCODE `0x800A03EC`.
No Excel process remained afterwards. This was captured before Prompt 20's
EXCEPINFO-string retention and diagnostics wrapper existed, so source,
description, version, Office bitness, workbook count, and `Ready` are honestly
recorded as unavailable rather than reconstructed. No second Add probe was
performed: Add is a non-idempotent operation even though the typelib may
describe it as a property get.

The normalized blocker and all future observations belong in
`knowledge/excel-object-model/chart-completeness-runtime/runtime-blockers.jsonl`.
The fixture system now provides controlled `Workbooks.Open` fallbacks without
touching a user workbook or changing Office configuration.

One controlled blank-fixture open ran after the ownership and fixture layers
were added. It passed integrity validation, opened and closed without calling
`Workbooks.Add`, then exceeded the original 15-second owned-process wait. The
exact process was absent immediately after the test harness returned, so this
is recorded as a partial pass with a deliberately increased 30-second future
wait bound, not as a hidden success or a forced cleanup.

## Runtime architecture

`OwnedApplication<'apartment>` is created only by the crate and is the sole
type with `quit`, `quit_and_wait`, and exact-process exit waiting. It observes
an owned process through that Application's own `Hwnd`, never process-name
enumeration, and it asks only for query/synchronization access. It never force
terminates a process. `AttachedApplication<'apartment>` obtains a shared
`IDispatch` through `GetActiveObject` and has no shutdown method; dropping it
only releases the crate's COM reference.

Excel registers a single active `Excel.Application` object through this
mechanism. Process-id selection and multiple-instance enumeration are omitted
because they were not proven reliable for classic Excel Automation. The API
returns a clear no-running-instance or ROT-access failure instead of guessing.

`ComMessageFilterGuard` is STA-thread-local, preserves the prior filter, and
restores it explicitly or best-effort on Drop. The filter itself never asks COM
to replay a call because it cannot know member idempotence. The dispatch layer
retains EXCEPINFO source/description and retries only safe reads and explicit
property puts under an opt-in bounded `ComRetryPolicy`. It classifies
`RPC_E_CALL_REJECTED`, `RPC_E_SERVERCALL_RETRYLATER`,
`RPC_E_SERVERCALL_REJECTED`, `CO_E_SERVER_EXEC_FAILURE`, and leaves
`0x800A03EC` permanent. Methods, object creation, deletion, and `Workbooks.Add`
are never automatically retried.

## Chart model

`ChartType` carries all 81 selectable values in the installed `XlChartType`
inventory and preserves unknown raw values. Selectability is not a claim that
the family has fully validated
creation, persistence, or export behavior. The support matrix marks those
unrun while the baseline remains blocked.

The wrapper now has typed `ChartGroups`, `Points`, point-level marker and data
label access, series marker/line and bubble controls, primary/secondary axis
minor and crossing controls, TickLabels, Gridlines, DataLabels, Trendlines,
ErrorBars, Office fill/line/color surfaces, chart/plot area geometry, walls,
floor, common 3-D controls, styles, layouts, and templates. Excel still
returns its own errors for chart-family-inapplicable members; the crate does
not emulate unsupported chart behavior.

Modern histogram, Pareto, box-and-whisker, waterfall, funnel, treemap,
sunburst, and map constants are selectable. The installed classic inventory
does not provide a stable, verified type-specific property set for these
families, so Prompt 20 intentionally does not claim their per-family controls.

## Remaining live campaign

The required ten-create/ten-Add/ten-fixture-open campaign and suite reruns are
not yet run because the first pre-layer baseline blocked and no test should
silently repeat Add. The next controlled run should first execute fixture-open
tests, then use the explicit retry/diagnostics layer to run the documented
matrix serially. It must record successes, failures, workbook-appearance after
error, and natural exit statistics without attaching to an unrelated session.
