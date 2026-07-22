# Microsoft C++ automation sample port (Prompt 06A)

## 1. Scope

This research starts from `origin/master` `4fe3c04c7eb75f937b1a2e9587d74d425cbfc28a`. It adds an independently written native C++ control, an independent `windows-sys` Rust port, structured evidence, and a narrowly scoped internal-codec refactor. It does not add a public Excel API or a high-level Excel wrapper.

## 2. Microsoft reference

The reference is Microsoft's previous-versions article, [How to automate Excel from C++ without using MFC or #import](https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/automate-excel-from-c). The page was retrieved on 2026-07-22 and preserved outside the repository at a temporary research location; its SHA-256 was `71E1575D2A43930EB941138CAD2C3EBB0B4F52B54B9316293636758A5DAC51C0`. The downloaded source is not committed because the page identifies Microsoft copyright.

The control retains the article's relevant calling convention: `CoInitialize(NULL)`, `CLSCTX_LOCAL_SERVER`, `IID_IDispatch`, `LOCALE_USER_DEFAULT` (1024) for `GetIDsOfNames`, `LOCALE_SYSTEM_DEFAULT` (2048) for `Invoke`, and `DISPATCH_PROPERTYGET` for `Workbooks.Add`.

## 3. Environment

The evidence records Windows 11 Enterprise build 26200, 64-bit Microsoft 365 Apps / Excel 16.0.20131, the Visual Studio 2022 MSVC 19.40.33812 x64 compiler, Rust 1.97.1, and `windows-sys` 0.61.2. The controls were run with no pre-existing `EXCEL.EXE` process. Each owned Excel process received `Quit` and its natural exit was verified; no process was force-terminated.

## 4. Native C++ control

`tools/excel-com-microsoft-sample/cpp/microsoft_excel_sample.cpp` is a native C++ implementation of the reference's COM sequence, adapted only to emit structured JSON lines and to avoid message boxes. It activates Excel locally, makes it visible, obtains `Workbooks`, invokes `Add` as a property get, writes the 15 by 15 `SAFEARRAY(VARIANT)` matrix, reads A1, B3, and O15, sets `Saved`, calls `Quit`, releases interfaces, clears variants, and finally calls `CoUninitialize`.

Twenty fresh native C++ runs completed. Every run read back A1=1, B3=6, and O15=225 and naturally exited. The recorded call trace has fourteen calls per run and exactly one `Quit` call. The source and evidence omit raw pointer values and PIDs.

## 5. Independent Rust port

`tools/excel-com-microsoft-sample` is a standalone `windows-sys` 0.61.2 tool. Its manual `IDispatch` vtable invocation and `AutoWrap`-style helper are independent from `tools/excel-com-range-probe` and preserve the C++ control's field values and ownership boundaries. It explicitly sets `Visible = true`, as the C++ control does.

Twenty fresh faithful Rust runs also completed with the same read-backs, one `Quit` trace per run, and verified natural exit. The C++/Rust comparison covers initialization, activation, requested IID, LCIDs, flags, zero-argument and property-put `DISPPARAMS`, argument ordering, result initialization, output-pointer mode, interface ownership, `SAFEARRAY` cleanup, and `CoUninitialize` timing. The recorded baseline fields are identical.

## 6. Controlled differentials

The faithful Rust baseline uses `DISPATCH_PROPERTYGET` for `Workbooks.Add`. A method-only control (5/5), a method-plus-property-get control (5/5), the property-get baseline (20/20), project-LCID invocation control (5/5), apartment initialization control (5/5), output `EXCEPINFO` control (1/1), output `puArgErr` control (1/1), combined observational control (1/1), no-visible control (5/5), and the current pre-Add sequence controls for `Hwnd`, `Version`, and their combination (5/5) all completed and naturally exited.

Consequently, no individual Microsoft-field difference, nor the combined current pre-Add field behavior, is established as the cause of the earlier raw `Workbooks.Add` error. The current raw transport call fields are intentionally unchanged.

## 7. Current-kernel and semantic live results

The historical Prompt 06 cold-boot observation remains intact: the raw IDispatch `Workbooks.Add` path returned inner SCODE `0x800A03EC` in that environment. After the independent C++ and Rust controls, the unchanged raw scalar smoke and the complete semantic suite succeeded. The state transition is not explained, so this is evidence of an environment-sensitive condition rather than proof of a transport repair.

An evidence-isolated observer captures the semantic suite without modifying Prompt 06 historical evidence. Ten required cases were runtime-observed and compatible under the documented normalizations: number, Unicode, `#N/A`, `Value` and `Value2` dates, currency, mixed 2 by 3 array, error 2 by 2 array, row vector, and column vector. The optional arbitrary error SCODE `0x81234567` was not accepted by Excel and read back as Empty. All observations naturally exited.

## 8. Internal value-layer follow-up

The automation module is decomposed into focused value, array, date, currency, error, conversion-error, policy, argument, encode, decode, and evidence modules. The module-wide dead-code allowance was removed; only narrow test-only allowances remain.

Generic BSTR conversion no longer applies Excel's 32,767-character cell limit. The range-write boundary applies that Excel-specific validation instead. Embedded NUL remains rejected by the existing BSTR allocation policy. This preserves a generic COM conversion boundary while enforcing worksheet limits at the worksheet boundary.

## 9. Evidence and validation

Structured source evidence is in `knowledge/excel-object-model/microsoft-cpp-port/`; generated reports are in `knowledge/excel-object-model/generated/microsoft-cpp-port/`. The standalone tool checks evidence shape, required source files, report presence, LF line endings, and pointer-redaction assertions. Deterministic range-probe tests and Clippy checks pass, as do the Microsoft control's tests and Clippy checks.

## 10. Conclusion and remaining risks

The native C++ reference and independent Rust control establish that direct external Excel Automation can presently create, populate, read, and cleanly close a local Excel process in this environment. The unchanged raw semantic suite subsequently demonstrates the same capability. That supports continued research use of the raw external-automation transport, with fresh-process and natural-exit checks retained.

It does not establish a durable explanation for the earlier `0x800A03EC` condition. The optional arbitrary Excel-error write is also not supported. Neither limitation warrants a speculative transport change or a public automation API. Prompt 07 may consume the internal codecs only behind a deliberately narrow, non-public transport facade.
