# Python client differential for Excel `Value`/`Value2` (Prompt 05I)

## Scope and baseline

This research-only follow-on starts at reviewed `origin/master`
`f00c0bbe1b6440b89c98cb7ca4dc58989ce4f92b` on branch
`research/excel-com-05i-python-differential`. It does not introduce a public
Rust API, a production COM abstraction, or an Excel-specific vtable. The raw
reference remains the generic SDK `IDispatch` kernel from Prompt 05H.

The evidence is stored in
[`python-client-differential`](../../../knowledge/excel-object-model/python-client-differential/).
Generated reports are deterministic and are kept under
[`generated/python-client-differential`](../../../knowledge/excel-object-model/generated/python-client-differential/).

## Safety and environment

Every live run required zero pre-existing `EXCEL.EXE` processes, made one
hidden workbook in its own instance, requested `Quit`, and recorded a natural
exit. No process-name termination was used for the accepted evidence. The
captured environments are Python 3.11 x64 with pywin32 311, pywin32 312, and
comtypes 1.4.16. Generated-wrapper caches were isolated outside the repository.

The run used pywin32 311 dynamic and generated wrappers, pywin32 312 dynamic
as the pinned-current comparison, and comtypes 1.4.16 dynamic and generated
wrappers. This is a client differential, not a Python implementation port.

## Source boundary

The source audit records the inspected installed package symbols, versions,
provenance, and licenses. Its conclusions are deliberately narrow:

- pywin32 exposes a public input `VARIANT` helper and separate dynamic and
  `EnsureDispatch` generated-wrapper paths.
- comtypes converts natural `None`, `datetime`, `Decimal`, and sequences while
  preparing a call; its dynamic and generated modes are measured separately.
- ordinary property-get results from either client are post-conversion Python
  values. They do not prove a returned physical VARTYPE.

The raw `windows-sys` observations therefore remain the reference for
VARTYPE, SAFEARRAY bounds and elements, HRESULT, EXCEPINFO, and `puArgErr`.

## Mixed SAFEARRAY result

Prompt 05H's single mixed 3x3 write is narrowed by a new raw fixed-position
matrix. Every non-target cell is a stable R8, BSTR, or BOOL control. In L
mode, replacing the target with Empty, Null, I4, Date, or Currency completed.
Replacing it with `VT_ERROR(2042)` failed with `0x80020009` and EXCEPINFO
`0x8007000E`; three immediate re-runs produced the same result.

The Python results add a conversion comparison, not a VARTYPE claim. The
tested pywin32 variants completed the recorded client-side mixed inputs.
Comtypes dynamic and generated variants rejected their `None`-based mixed
inputs with `0x80020009`, while their date, currency, and error candidates
completed. The source audit explains why these are client conversion paths and
why they cannot replace the raw fixed-position evidence.

## Date boundary

Raw `VT_DATE(-1.0)` and `VT_DATE(-0.5)` writes through `Value` failed with
`0x80020009` and EXCEPINFO `0x800A03EC`, regardless of General or date number
format. The same OA doubles through `Value2` completed. With a date format,
`Value` returned a date-shaped raw result while `Value2` retained the numeric
form; with General, both reads remained numeric.

The Python date attempts are separately recorded. Negative datetime-based
`Value` calls fail at the client/Automation boundary, while the `Value2` float
controls complete. The evidence does not collapse that distinction into a
single date conversion rule.

## Shapes, rank, and Formula2

The raw controls cover target/input pairs 1x2 to 1x3, 1x3 to 1x2, 2x2 to 2x3,
2x3 to 2x2, rank-1 row and column inputs, and rank-3 input. The observed
rank-3 value is rejected with `0x80020009` and `0x80020005`; the other listed
shape controls complete in this environment. Python natural sequence controls
are captured alongside those raw cases without assigning them a physical array
rank.

Formula2 `SEQUENCE(2,3)`, a text-producing spill, and a blocked spill are
captured through the raw L, S, and X paths and through each Python client mode.
The Formula2 property write itself completes in every raw mode; the spill
read-backs, including the blocked case, remain in the JSONL rather than being
reduced to a capability assumption.

## Cross-mode stability

Raw S and X each ran the mixed-array, shape/rank, direct 2x3 rectangular
read/write, and Formula2 groups twice in fresh owned Excel processes. The
JSONL preserves every result by run ID. The S/X records do not claim a general
locale conclusion; they establish only the repeated activation-path behavior
captured here.

## Internal requirements

Future internal value-model work must retain raw physical facts separately from
client-visible converted values. It must preserve input and output member
(`Value` or `Value2`), source VARTYPE when known, NumberFormat, SAFEARRAY
rank/bounds, exact error diagnostics, and post-failure reads. It must not infer
a physical VARTYPE solely from a Python `None`, float, tuple, datetime, or
Decimal.

No public conversion policy or API is finalized by this research. The
remaining explicit limitations are retained in `unresolved.jsonl`.

## Prompt 05J correction

Prompt 05J resolves this follow-on's raw `VT_ERROR(2042)` anomaly without
changing its client-visible corpus: `2042` is the Excel/CVErr number, while
formula returns expose the physical signed SCODE `0x800A07FA`. The separate
05J matrix verifies full-SCODE and raw-copy scalar and SAFEARRAY writes.

## Validation

`python-differential-check` verifies the required JSONL/TOML files, forbids
machine-specific pointer and PID fields, enforces LF endings, and compares all
generated reports deterministically. Standard workspace tests do not launch
Excel or Python clients; live collection is opt-in through the explicit
commands only.
