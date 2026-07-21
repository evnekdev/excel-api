# Prompt 05G — `windows-sys` kernel and Range recovery

## Decision

The recovery backend for the research Range probe is now
`raw-windows-sys`. The older high-level `windows` implementation remains an
opt-in diagnostic control; it is not the default recovery path and its result
cannot replace raw Automation evidence.

The kernel is derived from the successful Prompt 05E lower-level control. It
declares only the generic SDK `IDispatch` vtable, uses no Excel dual-interface
layout, and activates the registered `Excel.Application` local server through
three explicit modes:

| Mode | Activation | LCID |
| --- | --- | --- |
| L | `CoCreateInstance(CLSCTX_LOCAL_SERVER)` | `0x0400` |
| S | `CoCreateInstance(CLSCTX_SERVER)` | `0` |
| X | `CoCreateInstanceEx(CLSCTX_SERVER)` | `0` |

## Kernel contract

`tools/excel-com-range-probe/src/raw/mod.rs` owns the apartment, generic COM
pointer, BSTR, VARIANT, and EXCEPINFO lifetimes. Positional arguments are
reversed into `rgvarg`; zero-argument calls use null argument pointers; and
property puts pair their single argument with `DISPID_PROPERTYPUT`. Every
result is initialized with `VariantInit` and released once with
`VariantClear`. EXCEPINFO deferred fill-in and returned BSTR cleanup happen
before diagnostics are copied. `puArgErr` is initialized to `UINT_MAX` and is
only reported for parameter-index HRESULTs. No evidence records pointer,
HWND, PID, or local fixture path values.

Each child begins only when the parent observes zero `EXCEL.EXE` processes. It
derives the owned process from `Application.Hwnd`, calls `Quit`, releases all
owned references, verifies that exact owned process exits, and never forcibly
terminates any process.

## Runtime result

The controlled fixture is the checked-in CSV
`tools/excel-com-range-probe/testdata/controlled-fixture.csv`. For every
successful child, the flow was `Workbooks.Add`, close the returned workbook,
`Workbooks.Open` the fixture, set `A1.Value2 = 42`, read the value exactly as
numeric 42, call `ClearContents`, close, and quit.

Excel normalized the `Value2` read-back to `VT_R8(42)`, so the smoke assertion
checks the value exactly rather than incorrectly requiring `VT_I4`. The VARTYPE
is retained in the observation.

The raw kernel completed **30/30** fresh-process sequences: ten each in L, S,
and X. All recorded `Add`, `Open`, write, read, clear, cleanup, and owned-exit
results succeeded. The narrow retry policy was evaluated once after stability:
the first L attempt succeeded, therefore it was not eligible and no retry was
made. A retry is permitted only after an Add/Open failure before a workbook is
returned with inner EXCEPINFO `0x800A03EC`, and is exactly one new-instance
attempt.

The high-level diagnostic was also run once in L for comparison; it succeeded
in that current session. This is a control observation, not a conclusion that
the earlier instability is resolved or that all session contexts are equal.

## Evidence and remaining work

- [kernel evidence](../../../knowledge/excel-object-model/windows-sys-kernel/)
- [kernel design](../../../knowledge/excel-object-model/generated/windows-sys-kernel/kernel-design.md)
- [fresh-process matrix](../../../knowledge/excel-object-model/generated/windows-sys-kernel/repeatability.md)
- [backend comparison](../../../knowledge/excel-object-model/generated/windows-sys-kernel/backend-comparison.md)

The Prompt 05 scalar and rectangular `Value`/`Value2` matrix has not yet been
rerun through the recovered backend. It remains explicitly not tested in this
evidence set; no public API or conversion semantics are claimed from the smoke
test alone. Cold/clean/warm high-level comparison contexts are likewise
unresolved rather than inferred from the single current-session control.

## Prompt 05H follow-up

The recovered raw backend subsequently ran the dedicated Prompt 05H scalar and
rectangular matrix. Its evidence and conclusions are intentionally separate
from this recovery record: [Prompt 05H runtime matrix](05h-value-safearray-runtime-matrix.md).
