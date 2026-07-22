# Workbook file lifecycle and positional optional arguments (Prompt 09)

## Scope

The experimental `excel-com` slice now covers `Application.DisplayAlerts`,
`Workbooks.Open`, and workbook identity, save, copy, SaveAs, and close
operations. It remains Windows-only, apartment-bound, pointer-free at the
public boundary, and intentionally does not attach to existing Excel sessions.

## Optional-argument kernel

`PositionalArguments` stores values in the logical Excel signature order. It
encodes every omitted position as `VT_ERROR` / `DISP_E_PARAMNOTFOUND`, retains
interior and trailing missing values, supports a borrowed wrapper object as an
owned `VT_DISPATCH` call argument, and maps pre-call conversion failures to the
zero-based logical position. `IDispatch::Invoke` is the only reversal point.
No trailing-missing optimization is applied.

The registered Excel 1.9 type library records `Workbooks.Open` as DISPID 1923
with 15 positions and the selected modern `Workbook.SaveAs` as DISPID 3174
with 13 positions. The extracted member identifiers include their DISPIDs
because the inventory reconciles duplicate logical names structurally.

## Paths and options

Open, SaveAs, SaveCopyAs, and Close consume `Path`/`OsStr` values using
Windows `encode_wide` units. Inputs are passed exactly as supplied: there is no
canonicalization or `to_string_lossy`; embedded NUL is rejected without
including the caller path in an error. `WorkbookOpenOptions` covers the
registered UpdateLinks through CorruptLoad positions. `WorkbookSaveAsOptions`
covers FileFormat through Local, and the final WorkIdentity position remains
Missing. Password fields borrow `&str` and are redacted in `Debug` output.

`XlFileFormat`, `XlSaveAsAccessMode`, `XlSaveConflictResolution`,
`XlUpdateLinks`, `XlPlatform`, `XlCorruptLoad`, and `WorkbookOpenFormat` are
transparent numeric newtypes with curated constants, retaining raw values
where Excel can return a newer enum member.

Current runtime evidence returned Workbook.FileFormat as exact integral
`VT_R8` even though the registered type library describes the enum as an
integer. The getter accepts exact finite `VT_I4` and `VT_R8` representations
and keeps the public result as `XlFileFormat`.

## Alerts, identity, and close

`DisplayAlertsGuard` reads the previous value, sets the requested value, and
best-effort restores it on drop. `restore` reports an explicit restoration
failure and disarms the guard after success. Workbook properties return Excel's
FullName, Path, FileFormat, and ReadOnly values. `SaveCopyAs` does not change
the wrapper's current identity. `Workbook::close` consumes the wrapper;
`SaveChanges::Prompt` uses Missing, `Save` uses true, and `Discard` uses false.

## Error and inventory policy

Invocation errors retain the target wrapper type, member, resolved DISPID,
dispatch flags, HRESULT, EXCEPINFO SCODE, and COM argument index without
addresses. The inventory now records independent `surface_class` and
`roadmap_class` fields and type-library flags. IUnknown and IDispatch entries
have explicit inherited origins and remain structurally visible, but only
declared Excel members contribute to human coverage counts.

## Tests and runtime evidence

Unit tests cover all-Missing/interior/trailing argument positions, the single
COM reversal, conversion position attribution, Open and SaveAs layouts,
password redaction, Unicode/spaced path construction, and NUL rejection. The
opt-in visible `workbook_file_live` test owns a unique temporary directory,
creates and saves an `.xlsx`, verifies identity and format, saves a copy,
reopens read-only, checks values and a dynamic-array spill, records the
read-only Save result (the tested Excel completed it without a reported
error), verifies alert restoration after a local failure, then
quits Excel naturally and removes only its own temporary directory.

The normalized runtime record is
`knowledge/excel-object-model/workbook-file-lifecycle/runtime-observations.jsonl`.
It intentionally omits user paths, process identifiers, handles, pointer
values, and passwords.

## Deferred scope

OpenText, CSV import/export policy, file dialogs, workbook protection,
external-link operations, templates, legacy SaveAs overloads, routing,
existing-session attachment, and generic Automation invocation remain outside
this bounded wrapper. Prompt 10 should extend a proven wrapper surface rather
than treating this experimental API as stable.
