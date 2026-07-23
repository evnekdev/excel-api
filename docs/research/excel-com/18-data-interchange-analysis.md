# Text interchange, data transformation, and analysis

## 1. Scope and baseline

Prompt 18 adds source-derived, apartment-bound wrappers for Excel-owned text interchange, data transformation, what-if analysis, scenarios, and controlled external-link operations.

## 2. Prompt 17 runtime status

Prompt 17 observations remain historical. Prompt 18 obtained its own fresh-process baseline.

## 3. Drawing-module split

The drawing facade remains separate from the new `excel::data` subsystem; data operations were not added to drawing.

## 4. Data module architecture

`excel::data` separates text import/export, transformations, analysis, scenarios, links, shared types, and shared encoding helpers.

## 5. OpenText

`Workbooks::open_text` invokes Excel `OpenText`, which creates a workbook rather than altering an existing one.

## 6. FieldInfo representation

Field information is encoded as a one-based rank-two `VT_ARRAY|VT_VARIANT` with `[start, TextColumnType]` rows for delimited input; fixed-width starts use Excel's zero-based character positions.

## 7. Delimiters and qualifiers

Delimiter switches and quote qualifiers are explicit. A custom delimiter cannot be NUL.

## 8. Text and CSV export

`Workbook::save_as_text` delegates to the existing `SaveAs` implementation and exact Excel format values.

## 9. Locale behaviour

Excel owns numeric, date, separators, qualifiers, and `Local` behaviour; Rust does not alter OS regional settings.

## 10. Text to Columns

`Range::text_to_columns` preserves all optional positions and uses the shared FieldInfo encoding.

## 11. Fill operations

Fill directions delegate relative-formula adjustment and formatting behaviour to Excel.

## 12. AutoFill

AutoFill takes a destination Range; Excel validates its required containment geometry.

## 13. DataSeries

Data series validates non-finite Rust inputs before invoking Excel.

## 14. Flash Fill

Flash Fill is exposed as a version-dependent, heuristic operation, not a deterministic parser.

## 15. Transpose helpers

Transposition uses Excel copy state and `PasteSpecial`, then attempts `CutCopyMode` cleanup without reading the OS clipboard.

## 16. Advanced Filter

Copy-mode filtering requires a destination; in-place filtering can change row visibility.

## 17. Subtotal

Subtotal validates one-based group and total columns. Data should normally be sorted first.

## 18. Consolidate

Range sources use Excel-generated external addresses; `create_links` can create external-reference formulas.

## 19. Goal Seek

Goal Seek is Excel's numerical solver and returns Excel's Boolean success result.

## 20. Scenarios

Scenarios are typed Excel-owned collection and object wrappers; no detached Rust scenario engine is created.

## 21. Scenario summaries

`CreateSummary` returns the Worksheet reported by the registered Excel contract and mutates workbook structure.

## 22. What-if Data Tables

`Range::create_data_table` models Excel What-if Data Tables, not ListObjects; calculation can be expensive.

## 23. Row and column differences

The wrappers return Excel's potentially multi-area result and let Excel validate the comparison Range.

## 24. SpecialCells conveniences

Formula, constants, blanks, and visible-cell convenience methods map `SpecialCells` no-match EXCEPINFO `0x800A03EC` to `None` only in that narrow context.

## 25. External links

Link sources preserve Excel's exact names. Operations require explicitly named source strings and never resolve arbitrary paths.

## 26. Safe link-update policy

`AskToUpdateLinksGuard` captures, sets, restores explicitly, and makes best-effort drop restoration. Link prompts and macro security remain separate controls.

## 27. Persistence

Text export and link-persistence semantics require controlled live observation and are not claimed as verified.

## 28. Runtime observations

The fresh `Workbooks.Add` baseline failed once with EXCEPINFO `0x800A03EC`. No Excel process remained and no retry was made.

## 29. Version and locale dependence

Flash Fill is version-dependent; parsing, separators, and text encoding are locale-dependent Excel behaviour.

## 30. Explicit non-decisions

Power Query, QueryTables, ODBC/OLE DB, PivotTables, Solver, Analysis ToolPak, events, and arbitrary network links remain out of scope.

## 31. Rustdoc

Public data types and methods document Excel ownership, mutation, locale sensitivity, destination requirements, and destructive `BreakLink` semantics.

## 32. Inventory

The registry and generated object-model inventory map every implemented Prompt 18 member and record data-utility capability metadata.

## 33. Validation

The structural crate check passes; full workspace formatting, tests, Clippy, inventory, and docs checks remain required before merge.

## 34. Remaining blockers

All workbook-dependent Prompt 18 live tests are runtime-blocked because workbook creation failed and no documented owned blank fixture is present.

## 35. Recommended Prompt 19 scope

Resume controlled live confirmation after a stable owned-workbook baseline, then address external-data connectivity separately under explicit safety controls.
