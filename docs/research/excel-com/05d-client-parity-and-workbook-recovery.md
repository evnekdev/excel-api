# Client parity and workbook recovery research

**Status:** Prompt 05 is still blocked before a Rust-owned workbook is returned.
**Date:** 2026-07-21
**Baseline:** `origin/master` `c28defb8981d219ffa3c425693d0e4554c003e72`

## 1. Scope and evidence boundary

This Prompt 05D work adds source-derived, research-only parity configurations to
`tools/excel-com-range-probe` and deterministic evidence reports. It is not a
production Excel API, wrapper generator, apartment abstraction, or retry
framework. Client source, installed typelib data, and runtime observations
remain separate evidence categories.

## 2. Starting baseline

The branch `research/excel-com-05d-client-parity` starts from the reviewed
`origin/master` baseline `c28defb8981d219ffa3c425693d0e4554c003e72` (`Build
Excel client implementation knowledge base`). The Prompt 05, 05B, and 05C
documents, client knowledge base, runtime evidence, and both standalone tools
were present and their deterministic checks passed before modification.

## 3. Prompt 05C corrections

The two malformed arrow sequences in the Prompt 05C research document were
corrected at their source. Both knowledge-base generators now reject known
mojibake patterns before accepting generated output. The client generator also
normalizes CRLF only for deterministic comparison with its LF output; generated
files themselves retain LF and a final newline.

## 4. Python environment setup

Three isolated Python 3.11.7 x64 environments were used. Environment A used
pywin32 311 from its x64 wheel. Environment B used pywin32 312 from its x64
wheel after an isolated source build of inspected 312.1 exceeded its bounded
diagnostic interval. Environment C used pure-Python comtypes 1.4.16 on the x64
interpreter. Generated-wrapper caches were isolated per client mode; their
paths are not committed.

The source/environment records are generated under
[`client-implementations`](../../../knowledge/excel-object-model/client-implementations/).
No production dependency, Office bitness, COM registration, or security setting
was changed.

## 5. pywin32 version reconciliation

The installed 311 control no longer relies on an unqualified comparison with
the Prompt 05C 312.1 reference. Exact upstream tag `b311`
`8b328dffac71b7afaf2d72f47c4048f27a32f6c8` was inspected and compared with
the retained 312.1 reference
`a992023bd2d2ef57f8b605b43c1bcc29cdc619e9`. Released tag `b312`
`2a277cb5552756c2b4d42b524dc36d25e0bb6354` was also compared with the
reference.

`dynamic.CDispatch` is identical. `PyIDispatch` Invoke/InvokeTypes, error
handling, and cache mechanics are semantically equivalent for the selected
Automation spine. Record SAFEARRAY and typed `VT_INT` changes in `oleargs.cpp`
are material but outside the bounded Range corpus. The released 312 comparison
has no material selected Automation-path difference from 312.1. This supports
source parity for the selected 311 path, but does not assign a cause to an
intermittent Excel result. See the generated
[pywin32 comparison](../../../knowledge/excel-object-model/generated/client-implementations/pywin32-311-vs-312.md)
and [reconciliation report](../../../knowledge/excel-object-model/generated/client-implementations/source-version-reconciliation.md).

## 6. comtypes environment

Environment C used comtypes 1.4.16, matching the inspected source version.
Wheel metadata does not establish an exact upstream commit, so that boundary is
retained. Dynamic controls selected `comtypes.client.lazybind.Dispatch` and
generated controls selected the generated `_Application`, `Workbooks`, and
`_Workbook` pointer wrappers.

## 7. Control harness

[`client_control_harness.py`](../../../tools/excel-com-client-kb/scripts/client_control_harness.py)
creates a new Excel instance in a separate client process per run, captures the
bounded state, invokes `Workbooks.Add`, closes an owned workbook without
saving, calls `Quit`, and checks owned-process exit. A known-good temporary
workbook was produced by the pywin32 generated control for the Rust Open
comparison; the fixture itself and its path are not committed.

| Control | Result | Wrapper distinction | Cleanup |
| --- | --- | --- | --- |
| pywin32 311 dynamic | `Book1` created | `dynamic.CDispatch` | owned process exited |
| pywin32 311 generated | `Book1` created | makepy generated classes | owned process exited |
| pywin32 312 dynamic | `Book1` created | `dynamic.CDispatch` | owned process exited |
| pywin32 312 generated | `Book1` created | makepy generated classes | owned process exited |
| comtypes dynamic | `Book1` created | `lazybind.Dispatch` | owned process exited |
| comtypes generated | `Book1` created | generated dual-interface pointers | owned process exited |

These are bounded control observations, not proof that every corpus member is
runtime-equivalent.

## 8. Excel session-state observations

Each Rust operation activated a fresh owned Excel 16.0 session on 64-bit Office
under Windows 10 Enterprise 25H2 build 26200.8875. Before `Add` and `Open`, the
sessions reported zero workbooks, hidden visibility, false user-control, true
interactive/ready state, automation security `1`, and display alerts enabled.
Calculation returned an Excel error value and is therefore not treated as a
controlled setting. PID and start identity were recorded for owned-process
verification; raw HWND values and paths were not persisted. The generated
[environment stability matrix](../../../knowledge/excel-object-model/generated/runtime/environment-stability-matrix.md)
contains the normalized per-operation state.

## 9. Rust parity configurations

Five serialized research modes were implemented: `rust-baseline`,
`pywin32-dynamic`, `pywin32-generated`, `comtypes-dynamic`, and
`comtypes-generated`. Each records activation, CLSCTX, IID, STA initialization,
separate name-resolution and invocation LCIDs, DISPID source, argument policy,
`puArgErr` policy, result ownership, type-info handling, and the vtable boundary.
See [implemented parity](../../../knowledge/excel-object-model/generated/client-implementations/implemented-rust-parity.md).

## 10. Activation parity

The baseline retains `CoCreateInstance(CLSCTX_LOCAL_SERVER, IID_IDispatch)`.
The pywin32 dynamic configuration uses source-matched
`CoCreateInstanceEx(CLSCTX_SERVER, IID_IDispatch)`; the remaining client modes
use `CoCreateInstance(CLSCTX_SERVER, IID_IDispatch)`. This is a bounded
activation experiment, not a public default. All five configurations activated
an owned Excel session successfully.

## 11. COM initialization parity

Rust explicitly calls `CoInitializeEx(COINIT_APARTMENTTHREADED)`. pywin32 and
comtypes initialize their importing main thread from `sys.coinit_flags` or an
STA default. This comparison records the difference but does not choose a
future public apartment abstraction.

## 12. LCID parity

The preserved baseline uses `0x0400` for `GetIDsOfNames` and `Invoke`. The four
client-derived modes use source-supported LCID `0` for both. LCID is serialized
per observation and is not made a global production policy.

## 13. Type-info and interface parity

Every configuration records `GetTypeInfoCount` and `GetTypeInfo` availability.
The source records distinguish pywin32 generated descriptors and comtypes
lazybind from raw `IDispatch`. The comtypes generated control uses dual-interface
wrappers, but the Rust tool has no safely generated Excel bindings; it records a
raw dispatch fallback and does not hand-write a vtable layout.

## 14. Optional-argument parity

The frame builder keeps zero omitted arguments, explicit
`VT_ERROR/DISP_E_PARAMNOTFOUND`, `VT_EMPTY`, and `VT_NULL` distinct. Unit tests
cover their VARTYPE separation. The `Workbooks.Add` matrix uses the exact
zero-argument form rather than treating any explicit marker as equivalent.

## 15. Result and ownership parity

Every invocation starts with an initialized result `VARIANT`. A returned
`VT_DISPATCH` is cloned before `VariantClear`; argument backing and named DISPID
storage outlive `Invoke`; each result and temporary BSTR is cleaned exactly once.
No raw pointer value is recorded.

## 16. Error-handling parity

Parity modes initialize `puArgErr` to `UINT_MAX`, retain its raw value, and map
it only for bad-argument HRESULTs. For zero-argument `DISP_E_EXCEPTION`, the
record says that a returned raw value is not a source parameter. Deferred
EXCEPINFO fill-in is honored, BSTRs are copied then released, and both outer
`0x80020009` and inner `0x800A03EC` remain recorded.

## 17. `Workbooks.Add` matrix

All five Rust configurations resolved audited/runtime DISPID 181 and used a
zero-argument null-pointer frame. All returned `0x80020009` with Excel
EXCEPINFO `0x800A03EC`; none returned an owned workbook. Every owned Excel
process exited after `Quit`, with no forced termination. The complete generated
[Add matrix](../../../knowledge/excel-object-model/generated/runtime/workbooks-add-parity-matrix.md)
is the authoritative normalized result.

## 18. `Workbooks.Open` matrix

To separate template/new-workbook behaviour from broader workbook access, every
Rust mode also invoked `Workbooks.Open` for the known-good temporary fixture.
All resolved audited/runtime DISPID 1923, supplied one redacted BSTR argument,
and returned the same outer/inner error pair. Since both `Add` and `Open` fail,
this bounded observation does not support a new-workbook/template-only cause.
See the generated [Open matrix](../../../knowledge/excel-object-model/generated/runtime/workbook-open-parity-matrix.md).

## 19. Minimal Range smoke test

No Rust mode received a workbook from either operation, so the required
`A1.Value2 = 42`, read-back, and `ClearContents` smoke test was not entered.
There is consequently no VARTYPE/value claim and no Range runtime matrix claim.
The generated [smoke-test report](../../../knowledge/excel-object-model/generated/runtime/range-smoke-test.md)
records this explicitly.

## 20. Prompt 05 resumed work

Prompt 05 was not resumed. Its scalar, rectangular, mixed-type, formula, and
SAFEARRAY matrix remains incomplete because the prerequisite Rust-owned
workbook and smoke test did not succeed.

## 21. Remaining intermittent behaviour

The six isolated Python controls succeeded while all five raw Rust modes failed
under normalized fresh-session observations. The historic Prompt 05B pywin32
success and later 05C failures remain preserved. This is an inconclusive client
invocation/session boundary, not proof of a Rust source-level fault, a package
version difference, or a template issue. The generated
[blocker report](../../../knowledge/excel-object-model/generated/runtime/remaining-blockers.md)
keeps per-mode classifications.

## 22. Explicit non-decisions

This work does not decide a public API, crate layout, value enum, generated
wrapper strategy, apartment abstraction, retry/message-filter architecture,
event model, supported Excel-version range, or Office x86 policy.

## 23. Validation

The standalone client knowledge-base tests passed (13 tests), range-probe tests
passed (21 tests), and both standalone tools regenerated and checked their
evidence deterministically. Final repository-wide formatting, lint, test, and
documentation checks are recorded with the PR validation results. Committed
evidence was scanned for raw local paths, raw HWND values, pointers, virtual
environments, generated caches, mojibake, and missing deterministic outputs.

## 24. Recommended next step

Do not begin Prompt 06. First establish a safely generated typed interface path
for the bounded Application/Workbooks/Add spine, or isolate an external Excel
session-state condition with a controlled raw-Invoke versus generated-wrapper
comparison. Resume Prompt 05 only after a Rust parity mode reliably creates or
opens a workbook and the minimal Range smoke test succeeds.

Prompt 05E performed that bounded ABI comparison. It preserves this parity
record, finds that the minimal high-level reproduction succeeds on both tested
windows-rs releases, and leaves the production local/`0x0400` sequence as the
next bounded repair; see [the native ABI differential](05e-native-abi-differential.md).

Prompt 05F then tested that pre-`Add` sequence under a preserved cold baseline,
fresh-process prefix matrix, session transitions, and independent controls. It
found no causal prefix and retained a high-level runtime/path-sensitive failure
boundary; see [the pre-`Add` delta investigation](05f-pre-add-sequence-delta.md).
