# Typed collections, COM identity, and Range navigation (Prompt 10)

## 1. Scope and baseline

This Prompt 10 slice starts from `origin/master` `5ac5d477398cc31934403e623db7499f280c62d9` and retains the Windows-only, apartment-bound, pointer-free experimental `excel-com` boundary.

## 2. Typelib collection patterns

The registered Excel 1.9 typelib records Workbooks, Worksheets, and Areas as `Count`/`Item`/`_NewEnum` collections. `_NewEnum` is DISPID -4 and returns `VT_UNKNOWN`; Item is DISPID 170.

## 3. `IEnumVARIANT` ownership

The private owner holds one `IEnumVARIANT` reference, is neither cloneable nor transferable across threads, and releases it through RAII.

## 4. `_NewEnum` acquisition

The dispatch result is validated as `VT_UNKNOWN` or `VT_DISPATCH`, transferred out of its owned VARIANT once, then queried for `IID_IEnumVARIANT`.

## 5. `Next` semantics

Each call requests one item with initialized output storage. `S_OK` requires exactly one fetched item; `S_FALSE` with zero is end-of-sequence; every other result is a structured enumeration error.

## 6. Typed iterator design

WorkbooksIter, WorksheetsIter, and AreasIter implement `Iterator<Item = Result<_, ExcelComError>>` and convert only object-bearing VARIANT values into typed wrappers.

## 7. Terminal-error policy

A COM or conversion error fuses an iterator. Early drop simply releases the owned enumerator.

## 8. Canonical COM identity

Identity uses `QueryInterface(IID_IUnknown)` on both wrappers, compares only the internal canonical interfaces, then releases both. Public APIs are fallible methods, not `PartialEq`.

## 9. Workbooks collection

Workbooks supports count, one-based or string Item lookup, and `_NewEnum` iteration. The live test covers two unsaved workbooks and explicit close.

## 10. Worksheets collection

Worksheets supports the same typed pattern. Excel's default Add inserts before the active sheet in the observed run, so iterator order is verified against indexed Excel order rather than creation order.

## 11. Areas collection

Areas is a third concrete collection with Count, one-based Item, and iteration. It is produced by Range.Areas.

## 12. Range Cells and Item

Cells remains a Range wrapper. Item and cell use strict one-based checked indices and never synthesize addresses.

## 13. Offset and Resize

Offset preserves signed checked `i32` offsets. Resize requires two nonzero checked dimensions and uses Excel's property-get classification.

## 14. Rows and Columns

Rows and Columns return Range wrappers; legacy row_count and column_count retain their prior behavior through those properties.

## 15. EntireRow and EntireColumn

These return Range wrappers and are validated through addresses only, avoiding full-row or full-column value reads.

## 16. Multi-area Range construction

The bounded Application.union2 helper passes two borrowed Range dispatch arguments in Excel logical order and supports the live Areas case without a generic variadic API.

## 17. Collection metadata

Metadata now has a structured collection block with element type, exact member IDs, controlled index kinds, and iterator status. The generated dashboard lists every structurally detected collection.

## 18. Live observations

The visible, fresh-process control passed Workbooks, Worksheets, Areas, navigation, derived writes, one-based rejection, and natural shutdown. Normalized records are under `knowledge/excel-object-model/collections-identity-navigation/`.

## 19. Early-drop cleanup

A Worksheets iterator was dropped after one item; subsequent Excel calls completed and the owned application exited naturally.

## 20. Explicit non-decisions

This does not add formatting, Names, Charts, Shapes, ListObjects, generic public late binding, cross-thread marshaling, events, or mutation-during-iteration guarantees.

## 21. Rustdoc

Public wrappers, iterators, identity methods, Range navigation, and union2 document apartment affinity, fallibility, one-based indexing, and the intentional absence of PartialEq.

## 22. Validation

The workspace tests, doctests, strict rustdoc, inventory regeneration/check/diff, knowledge checks, and all three visible Excel controls are run before handoff.

## 23. Remaining blockers

Excel may materialize equivalent Range or Areas references as distinct canonical COM objects. The API accurately reports `false` rather than substituting address equality; this is recorded as runtime evidence, not a kernel failure.

## 24. Recommended Prompt 11 scope

Build on this private collection kernel with one small additional typed collection or bounded workbook/worksheet navigation feature, while preserving the evidence-per-case and visible-process discipline.
