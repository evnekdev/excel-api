# Excel Range `VARIANT` and `SAFEARRAY` runtime research

**Status:** blocked before any Range operation; this Prompt 05 attempt is not
complete.
**Date:** 2026-07-21

**Prompt 05B update:** the original raw `DISPPARAMS` empty-pointer defect was
repaired and the call was fully diagnosed, but `Workbooks.Add` still returned
Excel's application-level error before a workbook existed. The blocker is not
resolved; see the [Prompt 05B diagnostic](05b-workbooks-add-diagnostic.md).

**Prompt 05D update:** source-derived pywin32/comtypes parity modes reached
fresh owned Excel sessions but all five Rust modes failed both `Workbooks.Add`
and a known-good `Workbooks.Open` fixture before a workbook was returned. The
Range smoke test and this matrix remain blocked; see [client parity and workbook
recovery](05d-client-parity-and-workbook-recovery.md).

## 1. Scope and evidence boundary

This change adds only an isolated, Windows-only raw-COM research probe and its
separate runtime evidence layer. It does not change
`knowledge/excel-object-model/data/` or `knowledge/excel-object-model/typelib/`,
does not create `excel-com`, and does not expose a production dispatch or public
Rust API. Documentation and typelib evidence were treated as planning inputs,
not as runtime proof.

## 2. Environment and isolation

The attempted host was Windows 10 Enterprise 25H2 build 26200.8875, 64-bit
Office, Excel file version 16.0.20131.20154, and Excel Automation typelib 1.9
(`{00020813-0000-0000-C000-000000000046}`). The starting `origin/master`
commit was `2ac52effadafe6cd5b95b448f356a62389fa54f2`.

The raw probe uses `CoCreateInstance(CLSCTX_LOCAL_SERVER)`, obtains the created
application's `Hwnd`, derives the candidate PID from that window, opens only
that PID for query/synchronization, and records its creation-time identity
before creating a workbook. It never attaches through the ROT, enumerates for
selection, or kills by process name. On this run, setup failed at
`Workbooks.Add` before a workbook or Range existed. The guard requested
`Application.Quit`; it did not force-terminate any process. The probe could not
commit a bounded process-exit result after this setup failure, so cleanup is
labelled **Inconclusive** rather than assumed successful.

## 3. Probe architecture

[`tools/excel-com-range-probe`](../../../tools/excel-com-range-probe/) is a
standalone unpublished Rust program using the official `windows` 0.62.2
projection. It contains narrow research-only support for an STA, reflected
Application CLSID loading, audited-DISPID invocation, `VARIANT` clearing,
`SAFEARRAY` inspection/construction, workbook-specific cleanup, evidence
serialization, and deterministic report generation.

The live command is explicit. It creates no permanent workbook, uses no
macros, changes no registration or trust setting, and runs a separate C#
dynamic-Automation control only after raw capture succeeds. The control is a
projection cross-check, never an ABI authority.

## 4. Typelib declarations used

**Typelib-declared:** the probe consumes Prompt 04's Application coclass CLSID
from `typelib/coclasses.jsonl`, verifies name lookups against the audited
DISPID for `Application.Workbooks` (572), `Workbooks.Add` (181),
`Application.ActiveSheet` (307), `Worksheet.Range` (197), `Range.Value` (6),
`Value2` (1388), `Formula` (261), and `Formula2` (1580). It is prepared to
record the exact invocation kind and raw VARTYPE for each call.

## 5. Scalar `Value` and `Value2`

**Not tested:** no temporary workbook could be created, so no scalar cell,
empty cell, empty-string formula, text, Unicode, Boolean, numeric, date,
currency, or error result was read. There is no runtime VARTYPE conclusion.

## 6. Rectangular Range results

**Not tested:** no 1×1, row, column, 2×2, 2×3, 3×2, or mixed rectangular Range
was available. The generated report intentionally has no fabricated rows.

## 7. SAFEARRAY dimensions and bounds

**Not tested:** the probe implements raw `SafeArrayGetDim`, lower/upper-bound,
element-VARTYPE, and copied-element recording, but no Excel-returned array was
available. No rank, bound, element-type, or ownership conclusion follows.

## 8. Excel row/column mapping

**Not tested:** the fixture would write coordinate-distinct mixed and numeric
matrices, then record `SafeArrayGetElement` coordinate previews. Since the
workbook was blocked, no descriptor-order or row/column mapping claim is made.

## 9. Dates and Currency

**Not tested:** no `Value`/`Value2` date or currency read occurred. The probe
is prepared to distinguish `VT_DATE`, `VT_CY`, and `VT_R8`; formatting is not
used as a substitute for a raw runtime VARTYPE.

## 10. Empty, null, missing, and errors

**Not tested:** no worksheet `VT_EMPTY`, formula `""`, `VT_ERROR`, or Excel
error code was returned. The tool preserves raw signed and unsigned 32-bit
error codes and treats `VT_ERROR/DISP_E_PARAMNOTFOUND` as distinct from a cell
error, but this run supplied no observation.

## 11. Formula and Formula2

**Not tested:** no Formula or Formula2 get/set occurred, including the planned
ordinary scalar, text, Boolean, error, multi-cell, and attempted `SEQUENCE`
dynamic-array cases.

## 12. Scalar writes

**Not tested:** the tool has isolated cases for BSTR, Unicode text, Boolean,
`VT_I4`, `VT_R8`, `VT_DATE`, `VT_CY`, `VT_EMPTY`, `VT_NULL`, and an accepted
Excel error value, each with `Value2` read-back. None reached Excel.

## 13. Rectangular writes

**Not tested:** the tool can construct and clear matching 2-D
`SAFEARRAY(VARIANT)` inputs with 0- and 1-based lower bounds, including mixed,
numeric, and text matrices. No acceptance, mapping, or coercion result was
captured.

## 14. Shape mismatch and rejected forms

**Not tested:** jagged, rank-one, empty, zero-length, and shape-mismatched
inputs are explicit remaining targets. They were not inferred from the
`SAFEARRAY` constructor or a projection client.

## 15. Optional Range arguments

**Not tested:** the intended narrow comparison of `Range.Value` argument
omission, `VT_ERROR/DISP_E_PARAMNOTFOUND`, `VT_EMPTY`, and `VT_NULL` did not
reach a Range. `Address`, `Find`, and `Sort` were deliberately left for a
separate restoration-focused experiment.

## 16. Control-client comparison

**Not tested:** the C# dynamic-Automation control did not run because the raw
probe stopped before workbook creation. The control is designed only to compare
managed scalar and array projection shape on its own owned workbook; it cannot
override raw COM VARTYPE evidence.

## 17. Ownership and cleanup

**Control-confirmed:** none.

**Runtime-observed:** `Workbooks.Add` returned HRESULT `0x800A03EC` from the
created local server. The session guard requested `Application.Quit` and did
not force any process termination. A later process check found no running
`EXCEL` process, but because the raw probe's bounded identity record was not
committed after the setup failure, it is supporting diagnostic information—not
an exact ownership proof. The planned 1,000 scalar/matrix clear stress loops
were not run.

Prompt 05B subsequently confirmed the `Application.Workbooks` target,
audited/runtime DISPIDs, zero-argument `DISPPARAMS`, flags, LCID, result
initialization, and exception details. It repaired the empty-pointer frame and
released owned dispatch references before the exit wait. The call nevertheless
failed with `DISP_E_EXCEPTION` and inner `EXCEPINFO.scode` `0x800A03EC`; no
workbook, Range smoke test, or full matrix result was captured.

## 18. Version-specific limitations

All future observations must remain scoped to this exact Excel 16.0.20131.20154
Win64 / Windows 25H2 environment. This blocked attempt establishes only the
host limitation, not Excel's Range transport semantics.

## 19. Architecture candidates

**Architecture candidate, not public API:** defer a concrete `AutomationValue`,
`ExcelValue`, `RangeValues`, `RangeShape`, or `ExcelErrorValue` design until
the raw scalar and array observations run successfully. The probe design keeps
these questions open: separate Automation versus Excel values, scalar versus
two-dimensional row-major results, original bounds preservation, date/currency
representation, unknown error fallback, shape validation, and a raw escape
hatch.

## 20. Explicit non-decisions

No final type names, module layout, traits, apartment/activation API,
created-versus-attached policy, retry/message-filter strategy, `Send`/`Sync`
claim, event architecture, optional-argument builder, compatibility range, or
wrapper-generation strategy was selected.

## 21. Unresolved questions

The principal blocker is recorded in
[runtime unresolved evidence](../../../knowledge/excel-object-model/runtime/unresolved.jsonl):
the owned local server rejected `Workbooks.Add` with `0x800A03EC`. The remaining
unresolved entries deliberately include `Find`/`Sort`, `Address`, locale
formula syntax, embedded NULs, and rejected/jagged/empty array shapes.

The distinct [Prompt 05B record](05b-workbooks-add-diagnostic.md) preserves the
independent pywin32 success and the valid Rust invocation frames; it does not
reclassify the original blocked observation as resolved.

## 22. Validation evidence

The Prompt 02 knowledge-base check and Prompt 04 typelib audit check both
passed before editing. The new standalone tool's formatting, clippy, and pure
tests pass. The explicit live command completed with a **host-blocked** capture
(zero Range observations, one inconclusive case), producing deterministic
runtime manifests and reports without invoking a Range member.

## 23. Recommended next prompt

Do not advance into lifecycle, retry, events, or a public API. First rerun this
same isolated probe on a host where an owned `Workbooks.Add` succeeds and exact
post-`Quit` identity verification can be recorded. Only then classify scalar,
array, formula, error, write, optional-argument, and control results as
runtime evidence.
