# Excel `Value`/`Value2` and SAFEARRAY runtime matrix (Prompt 05H)

## 1. Scope and baseline

This research-only matrix starts from `origin/master` commit
`4a46f72004bc1d79c24d979c31d70b77acba0179` on branch
`research/excel-com-05h-value-safearray-matrix`. It records raw generic
`IDispatch` observations only; it introduces no public Rust API and no Excel
dual-interface vtable.

The exact environment is recorded in
[`environments.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/environments.jsonl).

## 2. Transport prerequisite

The rebooted controls-first gate passed with zero pre-existing Excel processes.
The earlier session-transition failure is preserved as historical evidence and
is not assigned a causal explanation. Every current live command refuses to
start when an Excel process exists and confirms its owned process exits.

## 3. Raw-kernel refactor

The raw kernel is split into apartment, COM pointer, BSTR, dispatch,
EXCEPINFO, process, VARIANT, SAFEARRAY, observation, Excel, and matrix
modules. The original 05G 30-run kernel evidence remains unchanged. The 05H
post-refactor gate recorded ten further successful fresh raw smokes.

## 4. Observation value model

`ObservedVariant` retains the physical VARTYPE and distinguishes Empty, Null,
Boolean, integral widths, floating-point bits, currency scaling, date bits,
UTF-16 BSTRs, errors, arrays, dispatch values, and unsupported tags. It does
not persist interface pointers, HWNDs, PIDs, or local paths.

## 5. SAFEARRAY inspection method

Inspection uses the SDK APIs `SafeArrayGetDim`, `SafeArrayGetLBound`,
`SafeArrayGetUBound`, `SafeArrayGetVartype`, `SafeArrayAccessData`,
`SafeArrayUnaccessData`, and `SafeArrayGetElement`. Construction uses
`SafeArrayCreate`, `SafeArrayCreateVector`, and `SafeArrayPutElement`; owned
arrays transfer to `VARIANT` cleanup exactly once.

## 6. Scalar `Value2`

The 28 required rows are recorded in
[`scalar-value2-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/scalar-value2-observations.jsonl).
Each row was repeated as one fresh owned Excel process. Accepted integral inputs
read back as `VT_R8`; `VT_NULL` and empty BSTR become an empty cell; non-finite
floating-point attempts become Excel error values; and an embedded-NUL BSTR is
stored only through its prefix.

## 7. Scalar `Value`

The matching 28 fresh-process rows are in
[`scalar-value-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/scalar-value-observations.jsonl).
The raw data retains both `Value` and `Value2` read-backs rather than reducing
them to a display-string comparison.

## 8. Empty/null/blank semantics

Never-written blank, `VT_EMPTY`, `VT_NULL`, empty BSTR, `ClearContents`, and a
formula returning an empty string are separate rows in
[`blank-null-empty-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/blank-null-empty-observations.jsonl).

## 9. Dates and currency

OA date and `VT_CY` inputs are written through `Value` with explicit
`NumberFormat` values and read through `Value`, `Value2`, `Formula`, and
`Formula2`. The observations retain the VARTYPE, scaling, format, and both
member reads in
[`date-currency-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/date-currency-observations.jsonl).

## 10. Excel errors

Controlled formulas create the supported core Excel errors. Raw evidence keeps
their `VT_ERROR`/`scode` values rather than converting errors to strings; the
display identities are only labels for the controlled inputs. See
[`error-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/error-observations.jsonl).

## 11. Scalar formulas

Arithmetic, reference, text, Boolean, error, date, empty-string, and Unicode
formulas are captured with Formula and Formula2 read-backs in
[`formula-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/formula-observations.jsonl).

## 12. Rectangular reads

All required shapes from `1×1` through `3×4` were populated cell-by-cell with
`row * 1000 + column` markers and read through both members. `1×1` is scalar;
other tested multi-cell ranges return `VT_ARRAY|VT_VARIANT`, rank 2, with
one-based bounds. See
[`rectangular-read-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/rectangular-read-observations.jsonl).

## 13. SAFEARRAY dimension mapping

`SafeArrayGetElement` marker traversal proves physical dimension 1 maps to
Excel rows and physical dimension 2 maps to columns in this environment. The
same source data records the bounds and per-element VARTYPEs.

## 14. Rectangular writes

SDK-created `SAFEARRAY(VARIANT)` values round-trip for the required shapes and
for one-, zero-, and non-zero-based lower-bound variants. Evidence includes the
input metadata, whole-range read-back, and individual-cell reads in
[`rectangular-write-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/rectangular-write-observations.jsonl).

## 15. Mixed arrays

A mixed variant array includes Empty, Null, Boolean, integer, floating point,
BSTR, error, date, and currency elements. The result is retained without
coercing it to a homogeneous Rust representation in
[`mixed-array-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/mixed-array-observations.jsonl).

## 16. Precision boundaries

The precision evidence covers the `2^53` boundary, integer limits, many-digit
decimal input, normal and subnormal small values, a large finite value, and
negative zero. It preserves bit patterns in
[`precision-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/precision-observations.jsonl).

## 17. String edge cases

The string rows retain UTF-16 lengths and a bounded preview only where safe.
They cover whitespace, controls, embedded NUL, Unicode, combining text,
formula-like input, apostrophe text, and the Excel cell-length boundary. See
[`string-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/string-observations.jsonl).

## 18. Formula2 and dynamic arrays

`=SEQUENCE(2,3)` is attempted through `Formula2`; capability and spill-range
result are recorded as observed rather than presumed in
[`dynamic-array-observations.jsonl`](../../../knowledge/excel-object-model/value-safearray-runtime/dynamic-array-observations.jsonl).

## 19. Locale comparison

The canonical semantic matrix uses L (`CLSCTX_LOCAL_SERVER`, LCID `0x0400`).
S and X are restricted to the representative stability subset, so this prompt
does not make a general locale conclusion.

## 20. Stability

The stability file contains fresh-process reruns and records every outcome;
the report deliberately exposes the subset boundary instead of claiming that
all semantic rows were multiplied across modes.

## 21. Value-model requirements

The generated [requirements report](../../../knowledge/excel-object-model/generated/value-safearray-runtime/value-model-requirements.md)
derives distinctions a future safe layer must preserve without declaring that
future API.

## 22. Explicit non-decisions

No public enum, conversion policy, production COM abstraction, Office setting,
registration change, or process-name termination is introduced.

## 23. Validation

The standalone `value-matrix-check` validates JSONL completeness, LF endings,
absence of prohibited machine-specific fields, and report determinism. The
workspace validation commands remain part of final handoff.

## 24. Remaining blockers

The required mismatch-array variants and the full multi-mode rectangular and
formula stability subset remain explicit follow-up work if they are not present
in the generated remaining-blockers report.

## 25. Recommended next step

Use this evidence to design, but not yet freeze, the safe value-model policy
and its explicit conversion boundaries.

## 26. Prompt 05I follow-on

Prompt 05I preserves this Prompt 05H corpus and records a separate
raw-kernel/Python-client differential in
[`python-client-differential`](../../../knowledge/excel-object-model/python-client-differential/).
It narrows the original mixed-array anomaly with a fixed-position 3x3 control:
in the observed environment, Empty, Null, I4, Date, and Currency replacements
complete, while the `VT_ERROR(2042)` replacement fails repeatedly with retained
HRESULT and EXCEPINFO. It also distinguishes negative `VT_DATE` through
`Value` from the same negative OA number through `Value2`.

The pywin32 and comtypes results in that follow-on are explicitly
client-visible post-conversion observations. They are not substituted for the
physical VARTYPE, SAFEARRAY, HRESULT, or EXCEPINFO observations in this matrix.
