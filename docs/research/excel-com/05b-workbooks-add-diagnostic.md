# Workbooks.Add raw-dispatch diagnostic

**Status:** narrowed, not resolved. `Workbooks.Add` did not create a workbook,
so the Range smoke test and full Prompt 05 matrix did not run.
**Date:** 2026-07-21

**Prompt 05D follow-up:** source-matched activation/LCID configurations did not
produce a Rust-owned workbook through either `Workbooks.Add` or `Workbooks.Open`.
The raw-diagnostic evidence remains preserved; see [client parity and workbook
recovery](05d-client-parity-and-workbook-recovery.md).

## Scope and baseline

This research-only repair is limited to `tools/excel-com-range-probe` and its
runtime evidence. It starts from `origin/master` commit
`dbbc9600af2628b45e3f05431ce168102ad9e6ae`. It does not create `excel-com`,
change an XLL, alter COM registration, modify canonical documentation/typelib
evidence, or freeze a public API.

The environment is Windows 10 Enterprise 25H2 build 26200.8875, Excel
16.0.20131.20154, Office 64-bit, and Excel Automation typelib 1.9
(`{00020813-0000-0000-C000-000000000046}`). Raw calls consistently use
`LOCALE_USER_DEFAULT` (`0x0400`).

## Independent control

**Control-confirmed:** the supplied Python pywin32 control used
`win32com.client.DispatchEx("Excel.Application")`, reported Excel version
16.0, and successfully ran `Workbooks.Add`, creating `Book1`. Its transient
HWND is intentionally not committed. The control proves neither raw
`VARIANT`/`SAFEARRAY` behavior nor a Rust invocation frame.

See the committed [control record](../../../knowledge/excel-object-model/runtime/observations.jsonl)
and [control comparison](../../../knowledge/excel-object-model/generated/runtime/control-comparison.md).

## Confirmed Rust defect and repair

The former helper reversed an empty `Vec<VARIANT>` and passed its non-null
dangling `as_mut_ptr()` values as empty `DISPPARAMS` pointers. The replacement
owned `InvocationFrame` builder now emits `cArgs=0`, `cNamedArgs=0`, and null
`rgvarg`/`rgdispidNamedArgs` for a zero-argument call. It also initializes the
result `VARIANT`, zero-initializes `EXCEPINFO`, supplies valid `puArgErr`
storage, retains every argument through `Invoke`, honors deferred exception
fill-in, and records no raw pointers.

## Invocation comparison

`Application.Workbooks` used audited/runtime DISPID 572/572, property-get
flag `0x0002`, LCID `0x0400`, and a zero-argument frame with both pointers
null. It returned `0x00000000`, `VT_DISPATCH`, and reflected type `Workbooks`.

`Workbooks.Add` targeted that returned dispatch object and used audited/runtime
DISPID 181/181, method flag `0x0001`, LCID `0x0400`, and the same null-pointer
zero-argument frame. It returned `DISP_E_EXCEPTION` (`0x80020009`); its
initialized result stayed `VT_EMPTY`. EXCEPINFO reported `Microsoft Excel`,
`xlmain11.chm`, and inner `scode=0x800A03EC`; deferred fill-in was absent.

The generated [Workbooks.Add diagnostic](../../../knowledge/excel-object-model/generated/runtime/workbooks-add-diagnostic.md)
and [invocation-frame report](../../../knowledge/excel-object-model/generated/runtime/invocation-frames.md)
retain the full copied diagnostics, including `puArgErr` as its physical
`rgvarg` index without guessing a source parameter.

## Retry and optional Template matrix

After the primary method call failed, the probe separately tried a 500 ms
delay, a bounded 500 ms client message-pump interval, a 2,000 ms delay, and
the combined method/property-get flag. Every call retained the same matching
DISPID, null zero-argument frame, and outer `0x80020009` / inner `0x800A03EC`
exception. The combined flag remains diagnostic only.

After the omitted form was recorded, separate one-argument calls used
`VT_ERROR/DISP_E_PARAMNOTFOUND`, `VT_EMPTY`, `VT_NULL`, and documented
`xlWBATWorksheet` as `VT_I4(-4167)`. None created a workbook; every one
returned the same outer/inner exception. Omission and the missing marker remain
separate observations, not equivalent representations.

## Narrowed classification and recovery

The most narrowly supported classification is **Excel returned an
application-level error despite a valid frame**. The empty-pointer defect was
real and is repaired, but it does not explain the remaining failure. The
evidence rules out the tested DISPID, target interface, zero-argument frame,
flags, LCID/name resolution, result initialization, optional inputs, and
bounded readiness variants. It does not contradict the pywin32 control or
establish a general Office, licensing, disk, memory, or registration problem.

No Add representation returned an owned workbook dispatch object. Therefore no
worksheet, `A1.Value2 = 42` smoke test, Range read/clear, C# projection
control, or full Prompt 05 scalar/array/formula/write/stress matrix ran.
Prompt 05 remains incomplete.

## Ownership, cleanup, and next blocker

The probe used `CoCreateInstance(CLSCTX_LOCAL_SERVER)`, obtained the created
Application HWND, verified its PID and process start time, and did not attach
through the ROT or select a process by enumeration. It requested
`Application.Quit`, released its owned Application and Workbooks references
before the bounded wait, and recorded `process_exited: true`. No process was
terminated by name or force, and no pointer or stable HWND was committed.

The standalone tool has pure tests for frame layout, flags, HRESULTs,
EXCEPINFO normalization, `puArgErr`, DISPID comparison, result initialization,
diagnostic serialization, report stability, and cleanup records. Final
repository and evidence validation outcomes are recorded in the draft PR.

Do not begin Prompt 06. The remaining gate is an isolated comparison of the
Rust raw activation/session path with the known-good pywin32 `DispatchEx`
control, followed by successful raw workbook creation and Range smoke testing.
