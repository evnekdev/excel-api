# Prompt 05E — native C++ versus Rust raw-COM ABI differential

## Scope and baseline

This record starts from `origin/master` `91ab21df735d38928e27775aea8f4ecad3499821`.
It preserves the Prompt 05B–05D evidence: raw Rust historically failed at
`Workbooks.Add` and `Workbooks.Open`, while isolated Python controls succeeded.
It adds only unpublished research controls; it does not introduce a production
Excel COM API or an Excel-specific handwritten vtable.

The observed environment was Windows 10 Enterprise 25H2 build 26200.8875,
Excel `16.0.20131.20154` 64-bit, x64 MSVC `19.40.33812.0`, CMake `3.30.2`,
Ninja `1.12.1`, and Rust `1.97.1`. The first temporary CMake configuration
selected MinGW and was rejected; the clean build was reconfigured for x64 MSVC.
No Office, COM registration, bitness, or security settings were changed.

## Controls and ABI audit

`tools/excel-com-native-control/` contains an SDK-only, x64 MSVC C++ control
and a narrow C ABI DLL. It uses `CoInitializeEx`, `CLSIDFromProgID`,
`CoCreateInstance`/`CoCreateInstanceEx`, `IDispatch::GetIDsOfNames`/`Invoke`,
`VariantInit`/`VariantClear`, and BSTR cleanup. The Rust C ABI caller receives
only fixed-width copied diagnostics; it receives no pointer or BSTR ownership.

The lower Rust path uses `windows-sys` activation plus the SDK's generic
`IUnknown`/`IDispatch` vtable order. It has no Excel-specific vtable. The
generated `windows 0.62.2` source was inspected for `IDispatch::Invoke`,
`IDispatch_Vtbl::Invoke`, `VARIANT`/its nested union, `DISPPARAMS`,
`EXCEPINFO`, `CoCreateInstance`, and `CoCreateInstanceEx`.

| ABI item | C++ SDK | Rust |
| --- | ---: | ---: |
| pointer width | 64 | 64 |
| `VARIANT` size/alignment | 24 / 8 | 24 / 8 |
| `DISPPARAMS` size/alignment | 24 / 8 | 24 / 8 |
| `EXCEPINFO` size/alignment | 64 / 8 | 64 / 8 |
| `GUID` size/alignment | 16 / 4 | 16 / 4 |
| fixed C result size/alignment | 88 / 4 | 88 / 4 |

All audited field offsets match: the `VARIANT` tag is at 0 and its used union
members at 8; `DISPPARAMS` is 0/8/16/20; and `EXCEPINFO` is
0/8/16/24/32/40/48/56. The ABI evidence and exact generated signatures are in
the generated reports linked below.

## Live differential

A Python control created the temporary fixture (9,302 bytes,
SHA-256 `D346736623BB0157C92AA86E6954699A25A47E50A581F79100FA0A560889D429`).
Its path is not recorded in committed evidence.

The C++ implementation through the Rust-to-native C ABI shim and the
`windows-sys` generic-`IDispatch` path both succeeded for `Application.Workbooks`,
`Count`, `Add`, fixture `Open`, close, `Quit`, the three lifetime sequences,
and owned-process exit in all three required activation modes. The full
high-level Rust harness failed only in
`CoCreateInstance(CLSCTX_LOCAL_SERVER)` with LCID `0x0400`:
`DISP_E_EXCEPTION` (`0x80020009`) and copied `scode` `0x800A03EC`.
It succeeded in the two LCID-0 server modes.

The final separately executed native C++ runner returned the same error in its
local/`0x0400` row even though the same C++ source invoked through the narrow
C ABI succeeded in all modes. That direct-runner inconsistency remains an
explicit unresolved control discrepancy; it prevents classifying this as a
clean Case D result. No raw pointer, HWND, fixture path, or unrelated process
data is persisted.

## Minimal version differential

The smallest standalone high-level reproduction is
`tools/excel-com-native-abi/repro/`. It performs only activation, `Workbooks`
property get, zero-argument `Add`, and `Quit`. It succeeded with:

| Test | windows | windows-core | windows-result | windows-strings | windows-sys | `Add` |
| --- | --- | --- | --- | --- | --- | --- |
| current/newest released | 0.62.2 | 0.62.2 | 0.4.1 | 0.5.1 | 0.61.2 | `S_OK` |
| immediately preceding | 0.62.1 | 0.62.1 | 0.4.0 | 0.5.0 | 0.61.1 | `S_OK` |

An isolated current `windows-rs` checkout at
`447078ea771a97277b710de1e3149c5146af1dc8` could not compile this released
Win32-feature reproduction: its source tree no longer exposes the released
`Win32_*` feature names. This is a source-head compatibility blocker, not a
runtime result.

Therefore **no windows-rs regression is confirmed and no upstream issue draft
is warranted**. The passing minimal high-level reproduction also rules out a
minimal standalone ABI failure in the current crate.

## Classification and Prompt 05 status

The evidence supports a narrow finding: the larger high-level local/`0x0400`
research sequence is sensitive to wrapper/ownership/session sequencing, while
the minimal high-level call and the lower-level path work. The native direct
runner versus native-shim discrepancy remains unresolved, so the required
matrix classification is **inconclusive**, not a claim of a crate regression
or a general raw-`IDispatch` failure.

Prompt 05 remains **blocked**: the production range probe has not been changed
or rerun to create a Rust-owned workbook, so no Range smoke test was resumed.
The next bounded task is to bisect the local/`0x0400` high-level pre-`Add`
sequence against the passing minimal reproduction, then repair and validate the
production probe.

## Evidence

- [source manifest](../../../knowledge/excel-object-model/native-abi/SOURCE_MANIFEST.toml)
- [operation specifications](../../../knowledge/excel-object-model/native-abi/operation-specs.jsonl)
- [ABI comparison](../../../knowledge/excel-object-model/generated/native-abi/abi-layout-comparison.md)
- [operation matrix](../../../knowledge/excel-object-model/generated/native-abi/native-vs-rust-operation-matrix.md)
- [lifetime comparison](../../../knowledge/excel-object-model/generated/native-abi/interface-lifetime-comparison.md)
- [remaining blockers](../../../knowledge/excel-object-model/generated/native-abi/remaining-blockers.md)

## Prompt 05F follow-up

Prompt 05F preserved a cold-session baseline, tested the full pre-`Add`
prefix sequence in fresh child processes, and compared later current-state
controls. The prefix matrix did not isolate a causal operation, while the
high-level path continued to show a path-sensitive failure boundary and the
lower-level/native controls succeeded. See [the 05F delta investigation](05f-pre-add-sequence-delta.md).
