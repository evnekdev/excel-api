# Client-implementation knowledge base

**Status:** source-derived evidence complete; client controls are recorded separately from Excel runtime evidence. Prompt 05 remains incomplete because the Rust probe still did not receive a workbook. Prompt 05D reconciles pywin32 311 with the inspected source and records bounded Rust parity experiments in [client parity and workbook recovery](05d-client-parity-and-workbook-recovery.md).
**Date:** 2026-07-21
**Baseline:** `origin/master` `7be1a691926862a5a82fb1b3937f6f3f7fb4a60d`

## 1. Scope and evidence boundary

This research covers successful open-source client implementation mechanics across the narrow spine `Application → Workbooks → Workbook → Worksheets → Worksheet → Range`. It records four evidence categories separately: documentation, installed Excel typelib, prior runtime observations, and client implementation source.

Source behaviour is not Excel runtime evidence. The original Prompt 05B pywin32 `DispatchEx` control remains **Control-confirmed**: Excel 16.0 created `Book1`. Current controls are separately **Inconclusive** and do not replace it.

## 2. Existing runtime problem

Prompt 05B established that the Rust empty `Vec::as_mut_ptr()` issue was real and repaired. `Application.Workbooks` succeeds with matching DISPID 572; `Workbooks.Add` matches DISPID 181 but returns `DISP_E_EXCEPTION` with `EXCEPINFO.scode = 0x800A03EC` despite a valid zero-argument frame. No workbook was returned, so no Range smoke test or Prompt 05 matrix ran.

## 3. Upstream source revisions and licences

| Client | Pinned revision | Source version | Licence | Installed match |
| --- | --- | --- | --- |
| pywin32 | [`a992023`](https://github.com/mhammond/pywin32/tree/a992023bd2d2ef57f8b605b43c1bcc29cdc619e9) | 312.1 | BSD-3-Clause | No: local pywin32 is 311 |
| comtypes | [`339ea278`](https://github.com/enthought/comtypes/tree/339ea278d85defda3d3c0dba118969021018e5fb) | 1.4.16 | MIT | Version match in isolated control environment; exact wheel-to-commit match unverified |

Both revisions were retrieved on 2026-07-21 from immutable commit URLs. Only selected files were inspected; neither upstream repository is vendored. Exact file lists are in the two `SOURCE_MANIFEST.toml` records.

## 4. Installed client environments

The regular control interpreter is Python 3.11.7 with pywin32 311. `pythoncom` initializes the main thread automatically, using `sys.coinit_flags` if defined or apartment-threaded initialization otherwise. For the requested opt-in comtypes control, version 1.4.16 was installed in an isolated Python 3.11.7 environment, not added to the repository or its production dependencies. Its package metadata has no source commit, so the installed version matches the inspected source version but not a provable exact commit.

The pywin32 dynamic control identified `win32com.client.dynamic.CDispatch` for Application and Workbooks. The makepy control identified generated `_Application` and `Workbooks` classes. Both currently failed at `Workbooks.Add` with outer `0x80020009` and inner `0x800A03EC`; this is a new, separate **Inconclusive** result.

The initial isolated comtypes dynamic control selected `comtypes.client.lazybind.Dispatch` for Application, Workbooks, and Workbook, and created `Book1` in Excel 16.0. Its initial generated control loaded the installed Excel typelib and selected generated `_Application`, `Workbooks`, and `_Workbook` interface pointers; that chain also created `Book1` in Excel 16.0. A later recheck in both modes reached the same wrappers but returned `0x80020009` / `0x800A03EC` at `Workbooks.Add`. The success and later **Inconclusive** rechecks are both retained as separate observations. Each control closes any created workbook and quits Excel in `finally` blocks. They are bounded controls for the observed chain, not a claim that all corpus members were runtime exercised.

## 5. Representative corpus

The corpus contains 42 invocation forms: the mandatory members plus distinct property-get/property-put forms and `_NewEnum` for Workbooks and Worksheets. It covers activation, scalar/object properties, property puts, zero/required/optional/large-signature methods, returned objects, default members, enumeration, value/array/formula transport, errors, and named-argument surface forms.

It deliberately excludes charts, pivots, shapes, events, connections, and the rest of the Excel object model. The deterministic [member matrix](../../../knowledge/excel-object-model/generated/client-implementations/representative-member-matrix.md) gives every selected member’s DISPID, descriptors, and client path.

## 6. pywin32 activation path

[`DispatchEx`](https://github.com/mhammond/pywin32/blob/a992023bd2d2ef57f8b605b43c1bcc29cdc619e9/com/win32com/client/__init__.py) calls `pythoncom.CoCreateInstanceEx` with `CLSCTX_SERVER`, requests `IID_IDispatch`, and passes the result to `Dispatch`. `Dispatch` chooses a generated wrapper where cache/type information permits; otherwise it constructs `dynamic.CDispatch`. Remote server information is supplied only when requested by the caller.

## 7. pywin32 dynamic dispatch

`dynamic.CDispatch` uses type-information maps when available and caches `GetIDsOfNames(0, name)` as a fallback. Its ordinary property path calls `PyIDispatch.Invoke`; its in-memory method generator can construct a typed `InvokeTypes` call when function descriptors are present. Therefore “dynamic” describes wrapper selection, not a guarantee that every member uses untyped `Invoke`.

Positional Automation arguments are reversed into `rgvarg`. `PyCom_MakeUntypedDISPPARAMS` initializes both pointer fields to null for zero arguments and adds `DISPID_PROPERTYPUT` for puts.

## 8. pywin32 generated dispatch

[`gencache.EnsureDispatch`](https://github.com/mhammond/pywin32/blob/a992023bd2d2ef57f8b605b43c1bcc29cdc619e9/com/win32com/client/gencache.py) upgrades a wrapper using type-library-generated `win32com.gen_py` classes. Fixed generated methods normally call `InvokeTypes` with return and argument descriptors. Generated property gets use the typed base path; generated `Item` projection and property setters use `Invoke`. Generated optional `pythoncom.Missing` arguments stop trailing non-byref `InvokeTypes` argument emission.

The generated local wrapper supplied the otherwise absent `Application.Hwnd` descriptor: DISPID 1950 and `VT_I4`. It is marked generated-wrapper-derived, not retroactively inserted into canonical typelib evidence.

## 9. comtypes activation and dispatch

[`CreateObject`](https://github.com/enthought/comtypes/blob/339ea278d85defda3d3c0dba118969021018e5fb/comtypes/client/_create.py) uses `CoCreateInstance` locally and `CoCreateInstanceEx` with server information. `dynamic=True` requests `IDispatch`; with type information it returns [`lazybind.Dispatch`](https://github.com/enthought/comtypes/blob/339ea278d85defda3d3c0dba118969021018e5fb/comtypes/client/lazybind.py), otherwise it falls back to `client.dynamic._Dispatch`. Generated mode loads type information through `GetBestInterface` and may query the selected generated interface.

The generated strategy is mixed by design. The parser follows `GetRefTypeOfImplType(-1)` for a dual dispinterface and emits a typed vtable interface; otherwise `DispMemberGenerator` emits `IDispatch.Invoke` methods. The audited Excel metadata correlates Application, Workbooks, Workbook, and Worksheet as dual candidates, while Worksheets and Range are dispatch-only. The initial generated control confirms the vtable path for the observed Application → Workbooks → Workbook chain; the later recheck is Inconclusive at `Workbooks.Add`, and the remaining selected interfaces were not control-run.

## 10. COM initialization comparison

pywin32 initializes the importing main thread from `sys.coinit_flags` or `COINIT_APARTMENTTHREADED`; later COM threads require explicit initialization. It does not automatically call `CoUninitialize`. comtypes also initializes at import using the same flag/default policy and registers shutdown uninitialization; its worker-thread callers explicitly pair initialization and uninitialization.

The Rust probe explicitly selects STA. This evidence does not choose a final Rust apartment model or retry policy.

## 11. Member resolution comparison

pywin32 dynamic caches `GetIDsOfNames` fallback lookups with LCID 0. Makepy embeds DISPIDs in generated maps/methods. comtypes dynamic caches `GetIDsOfNames(name)` and tests a property get before producing a callable; generated code embeds memids from the typelib.

All 42 records join current typelib DISPIDs except `Application.Hwnd`, whose generated-wrapper-only correlation is explicit. The audited/runtime 572 (Workbooks) and 181 (Add) facts remain unchanged.

## 12. Invocation frame comparison

pywin32 untyped and typed paths reverse logical positional arguments into `rgvarg`; generated type descriptors drive `InvokeTypes`. The observed comtypes dynamic `lazybind` path resolves through `ITypeComp.Bind`, then calls private `IDispatch._invoke`; its zero-argument `DISPPARAMS` leaves both counts at zero and both pointers null. The no-type-information `_Dispatch` fallback uses public `IDispatch.Invoke`, which similarly packages reverse-order `VARIANT` arguments. A generated comtypes dual-interface vtable call does not use a `DISPPARAMS` frame at the caller.

Both libraries’ dispatch helper default LCID is 0. The Rust probe’s `0x0400` is therefore a confirmed comparison point, not an instruction to change it.

## 13. Property put comparison

For Visible, Worksheet.Name, Value2, and Formula, both pywin32 and comtypes attach `DISPID_PROPERTYPUT` to the `IDispatch` frame and reverse the logical value argument. comtypes selects `PROPERTYPUTREF` when assigning an object. Generated dual vtable setters use declared parameters instead of an `IDispatch` named-DISPID frame.

## 14. Optional argument comparison

Omission, explicit missing, empty, and null are distinct. pywin32 uses `pythoncom.Missing` for `VT_ERROR/DISP_E_PARAMNOTFOUND`, and `None` becomes `VT_NULL`. comtypes exposes `VARIANT.missing`, `VARIANT.empty`, and `VARIANT.null`; its dynamic `Invoke` does not accept arbitrary named keywords. Generated comtypes `VARIANT` optionals without a typelib default use `VARIANT.missing`.

This covers Add, Open, Close, SaveAs, Worksheets.Add, Find, Sort, and Run at source level. It does not claim Excel accepts every form.

## 15. Collection and enumeration comparison

Both pywin32 modes emit an iterator using `DISPID_NEWENUM`; comtypes dynamic invokes -4 and consumes `IEnumVARIANT.Next`. Item and enumeration are separate mechanisms. In particular, client source alone does not establish Excel’s collection indexing base.

## 16. Scalar VARIANT conversion

Both clients map empty/null to `None`, booleans and numeric primitive values to Python scalars, BSTR to text, date/currency to client-specific date/decimal values, and dispatch/unknown values to wrappers. pywin32 returns an integer for `VT_ERROR`; comtypes documents unsupported getter combinations and raises `COMError` on failures. These are client conversions, not a proposed Rust value API.

## 17. SAFEARRAY conversion

pywin32 converts sequences to `SAFEARRAY(VARIANT)` and buffers to UI1 arrays, then walks result dimensions into nested Python sequences. comtypes creates variant or typed primitive SAFEARRAYs from list/tuple, `array.array`, or compatible ndarray inputs; its result path calls `_midlSAFEARRAY(...).unpack()` and tracks destruction. Neither client source establishes the rank, bounds, or shape Excel will return for Range.Value2.

## 18. Result wrapping and ownership

pywin32 converts a result then calls `VariantClear`, with object results routed through Dispatch selection. comtypes AddRefs dispatch/unknown pointers before wrapping and uses ctypes out-parameter/VARIANT cleanup. Both support the Rust requirement to make result ownership explicit, but neither specifies its public API.

## 19. Error and EXCEPINFO handling

pywin32 initializes `puArgErr` to `UINT_MAX`, invokes deferred EXCEPINFO fill-in before exception construction, translates from reverse `rgvarg` order only for applicable parameter errors, and frees EXCEPINFO BSTRs. comtypes exposes EXCEPINFO details for `DISP_E_EXCEPTION`, reports `puArgErr` for parameter-not-found/type-mismatch cases, and relies on ctypes structures for cleanup.

This corrects the **interpretation**, not the bytes, of Prompt 05B’s physical `rgvarg_index: 0`: a zero-argument `DISP_E_EXCEPTION` with inner `0x800A03EC` does not identify a bad source parameter. Historical observations are preserved.

## 20. Cross-member patterns

The generated [pattern report](../../../knowledge/excel-object-model/generated/client-implementations/invocation-patterns.md) and companion reports derive reusable rules for activation, gets/puts, methods, defaults, enumeration, optional/missing arguments, object wrapping, conversions, errors, DISPIDs, and LCIDs.

## 21. Typelib correlation

The client records preserve their own source claims, then join to the existing Excel typelib’s DISPID, INVOKEKIND, parameter count/types, optional/default flags, return type, default-member, and enumeration entries. The reports distinguish matching client normalizations from the one generated-wrapper-only `Hwnd` descriptor.

## 22. Rust parity backlog

Prompt 05D’s ordered backlog is generated in [rust-parity-backlog.md](../../../knowledge/excel-object-model/generated/client-implementations/rust-parity-backlog.md). Highest priorities are `puArgErr`, activation/COM initialization comparison, LCID experimentation, zero-argument frames, puts, optionals, returned dispatch ownership, SAFEARRAYs, and EXCEPINFO.

## 23. Explicit non-decisions

This prompt does not choose a Rust public API, crate layout, `AutomationValue` or `ExcelValue` design, final apartment model, retry policy, event architecture, object coverage, supported Excel versions, or public generated/dynamic wrapper strategy.

## 24. Unresolved questions

The historical pywin32 success versus current pywin32 311 controls is unresolved. Initial comtypes dynamic and generated controls are confirmed for the Application → Workbooks → Workbook chain, but their later rechecks are Inconclusive and no mode is confirmed for every selected interface. LCID 0 versus `0x0400` has not been isolated, and Range SAFEARRAY runtime shape remains blocked on Rust workbook creation. See [unresolved.md](../../../knowledge/excel-object-model/generated/client-implementations/unresolved.md).

## 25. Validation

The tool’s deterministic unit suite covers source manifests, stable IDs, wrapper modes, flags, named puts, optional metadata, typelib joins, report determinism, portable paths, and Prompt 05B preservation. Its `check` command regenerated and byte-compared every evidence/report file. Repository-wide validation is recorded with the PR.

## 26. Prompt 05D handoff

Review this knowledge base and backlog before changing the Rust invocation kernel. Prompt 05D should implement only confirmed parity work behind bounded diagnostics, first resolve the activation/LCID comparison, then resume Prompt 05 only after a workbook and Range smoke test succeed.

Prompt 05E later compared the generated wrapper with a lower Rust ABI path and
a native C ABI control without rewriting these source findings. The resulting
version matrix and remaining local/`0x0400` sequence blocker are documented in
[the native ABI differential](05e-native-abi-differential.md).
