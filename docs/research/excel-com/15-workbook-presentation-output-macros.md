# Prompt 15 — worksheet lifecycle, presentation, output, and macro safety

## Scope and baseline

Implemented the Prompt 15 slice on
`implementation/excel-com-15-workbook-presentation-output-macros`, based on
`origin/master` commit `7249f02c39c7f73f72344afad4c186836a31b4c5`.

The required minimal baseline ran `workbooks_add_baseline_live` three times in
separate processes. Every run began with zero `EXCEL.EXE` processes and created
an Application, but `Workbooks.Add` returned `DISP_E_EXCEPTION` / Excel SCODE
`0x800A03EC` before a workbook existed. All test-created Excel processes exited
naturally. This blocks the new live suite at setup; it does not justify further
host diagnosis or a claim that later operations were exercised.

## Safe object model

`Sheets` is heterogeneous and returns `Sheet::Worksheet(Worksheet)` or
`Sheet::Other(SheetObject)`; neither variant exposes raw dispatch pointers.
`WorksheetAddOptions` supplies typed `Before`, `After`, count, and sheet type,
while the worksheet-only collection rejects a non-worksheet type instead of
returning a false `Worksheet`. `SheetVisibility` and the other Excel enum-like
types are transparent `i32` newtypes, preserving future Excel values.

The implemented navigation path covers active workbook/sheet/cell/range/window,
safe Windows and Sheets collections, selection navigation through `GoTo`, and
worksheet activate/select/copy/move/delete. Worksheet Copy uses the documented
mutually exclusive Before/After positions; omitting both deliberately retains
Excel's new-workbook behavior. [Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.copy)

## Views, layout, and protection

`Window` covers view, zoom, gridlines, headings, zeros, scrolling, splits, and
freeze panes. `PageSetup` covers orientation, paper size, margins, gridlines,
centering, black-and-white, draft, quality, error rendering, print areas/title
ranges, headers/footers, and `PageZoom`/`PageFit`. `PageZoom::Automatic`
encodes Excel's `False` value; numeric zoom remains a percentage, matching the
documented bool-or-number contract. [Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.zoom)

Range merge, layout, reading order, grouping, outline, locked-cell, and
hidden-formula APIs preserve mixed formatting results. Worksheet and workbook
protection accept explicit option structures whose custom `Debug`
implementations redact passwords. Page-break wrappers use range locations;
print options also redact printer and print-file names.

## Output and macros

Workbook, worksheet, and range wrappers expose print preview, nine-position
`PrintOut`, and fixed-format PDF/XPS export. The wrapper validates local
one-based page values and leaves printer/output validation to Excel.

`AutomationSecurityGuard` saves and restores Excel's process-global setting.
`Workbooks::open_safely` forces `msoAutomationSecurityForceDisable` only for its
open call and restores the prior setting on both success and error paths. The
three documented Office values are Low=1, ByUI=2, and ForceDisable=3.
[Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/office.msoautomationsecurity)

`Application::run_macro` accepts a macro string and at most thirty positional
Automation values, preserves their positional order, decodes scalar returns,
and rejects returned dispatch objects rather than leaking an untyped COM
pointer. This follows the documented `Macro, Arg1..Arg30` signature and its
positional-only rule. [Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run)

## Inventory, tests, and limits

The object-model inventory now classifies Sheets, Windows/Window, PageSetup,
Tab, Outline, and horizontal/vertical page-break wrappers as implemented
surfaces and registers every invoked member ID. Unit tests cover raw enum
preservation, positional print arguments, and secret redaction. The new ignored
live tests are `workbook_presentation_live`, `fixed_format_export_live`, and
`macro_runtime_live`; each is intentionally single-threaded when invoked.

No chart, PivotTable, query, external-data, add-in, raw `IDispatch`, or generic
macro-object API was added. The large Prompt 14 table module remains coherent
as one object-family implementation, so no unrelated table refactor was made.
