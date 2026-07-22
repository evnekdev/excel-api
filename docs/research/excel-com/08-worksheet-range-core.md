# Worksheet and Range core wrapper (Prompt 08)

## Scope

`excel-com` now implements the bounded path `Application -> Workbooks ->
Workbook -> Worksheets -> Worksheet -> Range`. It remains an experimental,
Windows-only wrapper over generic `IDispatch`; it does not attach to existing
Excel sessions, marshal wrappers across threads, expose raw pointers, or claim
the broader object model as supported.

## Wrapper contract

`Workbook::worksheets` returns `Worksheets`. The collection exposes Count,
one-based Item by index, Item by name, and `add(WorksheetsAddOptions)`. Add
encodes `Before`, `After`, `Count`, and `Type` in logical order, uses explicit
`Missing` markers for omitted positions, reverses only in the dispatch layer,
rejects Before plus After and zero Count before COM, and deliberately keeps
Type missing.

`Worksheet` exposes Name, Index, Visible, Range, and UsedRange. `Range`
exposes address and dimensions, Value/Value2 and Formula/Formula2 reads and
writes, and ClearContents. The descriptor registry is the only wrapper-to-COM
member mapping and uses the extracted member IDs. `_ClearContents` is the
type-library name corresponding to the selected `excel.range.clearcontents`
inventory entry.

## Automation values

The public values remain pointer-free. `AutomationValue` preserves Empty,
Null, Bool, Number, Text, Error, OaDate, Currency, and rectangular arrays.
Range array conversion uses an owned rank-two `SAFEARRAY(VARIANT)`, SDK
SafeArray APIs, and one-based physical indices that normalize to zero-based
row-major `AutomationArray` values. Error values retain their exact signed
SCODE; `ExcelError::NOT_AVAILABLE` is the full `0x800A07FA` value.

`Value` encodes `OaDate` as `VT_DATE`; `Value2` encodes it as a numeric serial.
Every Range setter first obtains the target dimensions. Scalars are valid only
for 1x1 targets, and arrays require exact rows and columns. A mismatch returns
`ExcelComError::Conversion(ConversionError::ShapeMismatch)` before a COM
setter is invoked.

## Runtime evidence

The opt-in visible live test covers worksheet collection navigation and add,
worksheet identity, scalar values, text, exact #N/A SCODE, date, negative
Value2 serial, currency, 2x3 and error-array round trips, Formula, Formula2
dynamic-array spill, a pre-COM shape rejection, and ClearContents. It passed
with zero pre-existing Excel processes and verified natural exit after Quit.
The normalized record is
`knowledge/excel-object-model/worksheet-range-core/runtime-observations.jsonl`.

## Rustdoc

The crate root denies missing public documentation and broken intra-doc links,
and warns on private intra-doc links. Public wrappers, collection options,
visibility enum, values, array construction, error representation, and
apartment ownership are documented. The `no_run` crate example covers the main
Application through Range flow. Rustdoc is distinct from the generated
type-library inventory.

Validation used `cargo test -p excel-com --doc` and
`RUSTDOCFLAGS="-D warnings" cargo doc -p excel-com --all-features --no-deps`.
No scoped missing-docs allowance remains.
