# Prompt 14 — structured data operations

## 1. Scope and baseline

Implemented the bounded table, filtering, sorting, validation, duplicate-removal, and structural Range surface on branch `implementation/excel-com-14-structured-data`, based on `origin/master` commit `e75cd335ba42af0dd83d63ff5fccbb2bce10e8eb`.

## 2. Structured-data object graph

The exposed path is `Worksheet -> ListObjects -> ListObject -> {ListColumns, ListRows, AutoFilter, Sort}` and `Range -> {AutoFilter, Validation, structural operations}`. All wrappers retain the crate's apartment-bound COM ownership model.

## 3. ListObjects

`ListObjects` supports Count, one-based and string Item lookup, and fallible `_NewEnum` iteration. Embedded-NUL names fail before COM.

## 4. Table creation

`ListObjects::add_from_range` calls `ListObjects.Add` with `xlSrcRange`, the Range dispatch value, missing LinkSource, header mode, optional destination, and optional style in preserved positions. Multi-area source Ranges are rejected locally.

## 5. ListObject

The wrapper covers table name/display name, full Range, visible state, style, resize, delete, unlist, collections, filter, and sort. Excel remains authoritative for names and source/overlap validation.

## 6. Table Range components

`HeaderRowRange`, `DataBodyRange`, `TotalsRowRange`, and `InsertRowRange` are `Option<Range>`. The live header-and-data table returned no InsertRowRange; an empty DataBodyRange remains supported as no object.

## 7. ListColumns and ListColumn

`ListColumns` supports Count, both Item lookup forms, Add, and enumeration. `ListColumn` exposes index, name, full Range, optional data/total ranges, totals calculation, and consuming delete.

## 8. Calculated columns

Calculated formulas use the column DataBodyRange directly with Formula or Formula2. The dedicated internal call intentionally permits Excel's calculated-column propagation, while ordinary scalar multi-cell Formula assignment remains protected.

## 9. Totals rows

`TotalsCalculation` preserves raw `XlTotalsCalculation` integers and exposes common constants. Totals row visibility and a column SUM total were live-tested.

## 10. ListRows and ListRow

`ListRows` supports Count, one-based Item, Add with optional position/AlwaysInsert, and enumeration. A ListRow Range covers table columns only; delete consumes its wrapper.

## 11. Structured references

The live test set `=[@Quantity]*[@Price]` as a calculated-column formula. Structured-reference syntax and any rename/unlist text rewriting are delegated to Excel, not parsed by Rust.

## 12. AutoFilter

Range AutoFilter accepts explicit one-based field and bounded criteria. `AutoFilter` returns its Range and typed Filters collection.

## 13. Filters and criteria

Scalar criteria use the existing scalar VARIANT encoder. `FilterCriterion::Values` uses a rank-one, zero-based `SAFEARRAY(VARIANT)`; `RemoveDuplicates` uses the same physical family for its column vector. The live scalar table filter reported `Filter.On = true`.

## 14. Range sorting

`RangeSortOptions` preserves the 15 logical legacy Range.Sort positions, including PivotTable-only holes. Sort modifies the receiver in place and Excel validates key containment.

## 15. Sort and SortFields

Persistent Sort exposes SortFields, header, case, orientation, and Apply. The live ListObject sort successfully cleared fields, added a Quantity key, and applied ascending order. On this Excel build, `SetRange` on a ListObject-owned Sort returned Excel error `0x800A03EC`; the table-owned Sort is already bound.

## 16. Data validation

Validation Add retains all five positions. Validation formulas are passed as Excel syntax; a list validation with `yes,no` and input title round-tripped live.

## 17. RemoveDuplicates

Columns are nonempty, unique, one-based receiver-relative integers. The runtime accepted the one-dimensional `SAFEARRAY(VARIANT)` column vector and compacted duplicate rows in place.

## 18. CurrentRegion

`Range::current_region` delegates to Excel. It is bounded by blank rows and columns and is not treated as a dataset abstraction.

## 19. UsedRange

`Worksheet::used_range` remains a direct Excel property. It may include formerly formatted cells and accessing it can update Excel's own used-range state.

## 20. Hidden dimensions

`Range::hidden` returns `MixedValue<bool>` and `set_hidden` passes the exact Range through. The live test hid and restored an EntireColumn; arbitrary cells are never silently expanded.

## 21. Insert and Delete

`Range::insert` preserves optional Shift and CopyOrigin. The typelib return is Variant rather than a usable Range, so Rust returns unit; delete consumes its wrapper and preserves Excel shifting behavior.

## 22. Clear operations

Clear, ClearFormats, ClearComments, and ClearHyperlinks map to separate installed typelib members. Existing ClearContents remains unchanged.

## 23. Copy and Cut

Copy and Cut accept safe optional destination Ranges. A missing destination intentionally leaves Excel in cut/copy mode; the live test used both explicit destination and Excel-controlled copy mode.

## 24. PasteSpecial

PasteSpecial has explicit paste type, operation, skip-blanks, and transpose positions. The live values-only paste succeeded without inspecting arbitrary OS clipboard contents.

## 25. Persistence

The live test used an unsaved transient workbook and closed it without saving. Persistence of tables, validation, and filter state across SaveAs/Open is deferred to a focused lifecycle test.

## 26. Physical VARIANT observations

Optional Range wrappers accept `VT_EMPTY`, `VT_NULL`, and null `VT_DISPATCH` as no Range. The passing live calls establish rank-one `SAFEARRAY(VARIANT)` acceptance for RemoveDuplicates; scalar filter criteria use BSTR values through the existing encoder.

## 27. Version-dependent behaviour

The registered Office type library is version 1.9. `InsertRowRange` availability and ListObject Sort.SetRange behavior vary with UI/table state and are therefore not normalized by the wrapper.

## 28. Live results

`cargo test -p excel-com --test structured_data_live --offline -- --ignored --test-threads=1` passed once after a zero-process precondition and normal Quit/close cleanup. It covered table creation, calculated column, totals, filtering, sort, validation, duplicate removal, region/used range, copy/paste, insert/clear, hidden column, and unlist. Subsequent attempts, including unchanged historical live tests, were blocked before workbook creation by Excel's `Workbooks.Add` `0x800A03EC` error; their transient Excel processes passively exited without intervention. PowerShell Automation also failed on the read-only fixture's `Workbooks.Open` path in that state.

After a user-performed reboot, the `cold-boot-no-prior-excel` baseline recorded boot time `2026-07-22T20:16:42+10:00`, test start `2026-07-22T20:21:04+10:00`, and zero pre-existing `EXCEL.EXE` processes. The independent Microsoft-style raw control, an isolated minimal high-level reproduction, the full high-level control, and a generic `windows-sys` `IDispatch` control all reached `Workbooks.Add`, received `DISP_E_EXCEPTION` / `0x800A03EC`, called `Quit`, and naturally left zero processes after passive waits. Native C++ direct and C-ABI-shim controls were not reached because CMake configuration stalled before compilation or an Excel launch. This cold result blocks further live calls; it does not alter the earlier passing structured-data observation.

The user then opened and closed Excel interactively. A new zero-process structured-data run still failed at the unchanged `Workbooks.Add` invocation with the same `0x800A03EC`, then naturally left zero processes. Interactive opening and closing is not an observed recovery for this condition.

## 29. Explicit non-decisions

No PivotTables, query/external/model/XML table creation, table style-object wrappers, color/icon/custom-list sorting, data-type SubFields, Power Query, charts, or clipboard inspection were added.

## 30. Rustdoc

Every public Prompt 14 type and operation is documented under the crate's `deny(missing_docs)` policy. Crate-level docs explain optional data bodies, one-based fields, stateful filtering, structural changes, and clipboard implications.

## 31. Inventory

The inventory now maps every implemented member, promotes table/filter/sort/validation wrappers, records typed collection metadata, structured-data capabilities, and the requested member-level behavior flags.

## 32. Validation

Unit tests cover positional holes and local argument rejection; the new live integration test verifies the supported end-to-end path. Workspace formatting, Clippy, tests, docs, inventory extraction/generation/check, and KB check are required final gates.

## 33. Remaining blockers

The current Excel host must return to accepting `Workbooks.Add` or `Workbooks.Open` before the required full live regression sweep, table persistence matrix, and remaining live cases can be completed. This is an external host-state blocker: the structured wrapper does not modify the unchanged Workbooks.Add descriptor or Office registration. The failed Microsoft-style raw, minimal high-level, full high-level, and generic `windows-sys` controls on a cold reboot, together with the failed PowerShell `Open` control, exclude stale Excel processes and a Rust-only dispatch path as sufficient explanations. Persistence and broad version/locale matrices remain deliberate follow-up work rather than implied compatibility guarantees.

## 34. Recommended Prompt 15 scope

Prioritize workbook/worksheet protection and selected non-table data operations only after a new exact type-library and live-evidence plan. Keep table query, pivot, chart, event, and external-data expansion separate.
