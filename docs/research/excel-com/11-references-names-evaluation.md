# Prompt 11: References, Names, and Evaluation

## 1. Scope and baseline

Started from `66a8f29f5287ec6ee478ed5ca60a45f2252565f1`. The implementation adds a bounded, apartment-bound reference, defined-name, formula-conversion, and typed-evaluation slice without exposing raw COM interfaces or a formula parser.

## 2. Reference-style representations

`ReferenceStyle` is a forward-compatible `i32` newtype: `A1 = 1` and `R1C1 = -4150`. `ReferenceAbsoluteMode` preserves the four `XlReferenceType` values `1..=4` and unknown values.

## 3. Public Range selection API

`Worksheet::range(&str)` replaces the former `AutomationArgument`/optional-second-argument signature. `range_between`, `range_r1c1`, `cell`, and `range_from_cells` make selection intent explicit.

## 4. A1 selection

Excel accepted `A1:X10`, full-column `A:A`, full-row `1:1`, and sheet-qualified `'Sheet Name'!A1`. Excel remains the syntax authority; embedded NUL is rejected before COM.

## 5. R1C1 selection

Direct `Worksheet.Range("R1C1")` failed with Excel's invocation error under both observed global styles. `Worksheet::range_r1c1` therefore uses `Application::convert_formula` from R1C1 to A1, then calls `Worksheet.Range` with the converted A1 text. No global state is mutated.

## 6. Numeric selection

Indices are one-based and must fit `i32`; zero and overflow fail before COM. The implementation obtains `Worksheet.Cells`, derives both cells as Ranges, and passes those Range objects to `Worksheet.Range`.

## 7. Range.Address

`RangeAddressOptions` maps exactly to `RowAbsolute`, `ColumnAbsolute`, `ReferenceStyle`, `External`, and `RelativeTo`, retaining `Missing` for omitted optional values. Observed absolute output was `$A$1:$X$10` (A1) and `R1C1:R10C24` (R1C1).

## 8. Relative and external addresses

Observed relative R1C1 output was `R[1]C[1]` for B2 from A1, `R[-1]C[-1]` for A1 from B2, and `RC` for the same cell. For an unsaved workbook, external A1 output was `'[Book1]Sheet Name'!$A$1:$X$10`; it is context text, not a path or identity. Excel quoted spaces and escaped an apostrophe as `O''Brien`.

## 9. Application.ReferenceStyle

`Application::reference_style` reads the global setting. `reference_style_guard` reads, changes, explicitly restores, and attempts restoration on drop without panicking. The live test restored its original A1 setting.

## 10. Formula conversion

`Application::convert_formula` preserves all five logical `ConvertFormula` positions and uses Excel's engine. It converted `$A$1:$X$10` to `R1C1:R10C24`, converted the reverse direction, converted `=SUM(A1:X10)`, and converted a relative R1C1 expression using `RelativeTo`.

## 11. Workbook Names

`Workbook::names` exposes Excel's workbook Names collection. A workbook Range name `InputRange` increased collection count, supported index/name lookup and iteration, and resolved through `Name::range`.

## 12. Worksheet-local Names

`Worksheet::names` exposes the local collection. Excel returned the local name as `'Sheet Name'!LocalInput`, rather than the unqualified creation string.

## 13. Names collection and iteration

`Names` uses the existing safe collection helper and `_NewEnum` `IEnumVARIANT` implementation. It has one-based numeric lookup, string-key lookup, a fallible fused iterator, and early enumerator release through normal drop.

## 14. Name creation

`NameAddOptions` supports Range, A1, R1C1, and formula inputs. Range targets are represented by Excel's own external A1 address. A1/formula use `RefersTo`; R1C1 uses the positionally correct `RefersToR1C1` slot.

## 15. RefersTo and RefersToR1C1

For `InputRange`, Excel returned `='Sheet Name'!$B$2:$B$3` and `='Sheet Name'!R2C2:R3C2`. Both include the leading equals sign.

## 16. RefersToRange

`Name::range` uses `RefersToRange` and preserves Excel errors. The Range-backed name resolved with the same external address but a different canonical COM identity. A constant `=42` name retained `RefersTo` but produced Excel's structured invocation error for `RefersToRange`.

## 17. Excel-backed evaluation

`Application` and `Worksheet` expose separate `evaluate_value` and `evaluate_range` methods. `SUM(InputRange)` evaluated to the Automation number `5`; `InputRange` evaluated as a Range. Category mismatches return structured `Unsupported` errors directing the caller to the other method.

## 18. Scope collision behaviour

Excel permitted workbook `InputRange` and worksheet `'Sheet Name'!InputRange`. In this observation, both Application and Worksheet simple-name evaluation resolved to `'[Book1]Sheet Name'!$A$1`; this is recorded as Excel-defined rather than a crate conflict policy.

## 19. COM identity versus address equivalence

`Range::is_same_object` uses canonical `IUnknown`. A Range resolved from `InputRange` had the same external address as its source Range but returned `false` for canonical identity. The crate adds no `PartialEq`.

## 20. Live observations

`reference_names_live` starts only when no Excel process exists, creates a visible fresh instance, discards its unsaved workbook, explicitly restores `ReferenceStyle`, deletes created names, calls `Quit`, and waits for natural process exit.

## 21. Explicit non-decisions

This change does not add FormulaLocal assignment, FormulaR1C1 assignment, array-formula APIs, dependency tracing, a formula tokenizer, generic object evaluation, events, formatting, tables, charts, or a name-conflict policy.

## 22. Rustdoc

All public reference, address, conversion, Names, Name, and guard items have Rustdoc. The crate documentation has a selecting/naming/converting/evaluating API table and compiling `no_run` examples.

## 23. Inventory

The generated inventory maps every implemented member to its registered metadata ID. `Names` is recorded as a `Name` collection with one-based/string indexing and implemented iteration. Application, Worksheet, and Range carry controlled reference/evaluation capability metadata.

## 24. Validation

The final validation runs formatting, clippy, workspace tests, doctests, strict Rustdoc, generated documentation checks, knowledge-base validation, and all four ignored `excel-com` live tests.

## 25. Remaining blockers

No implementation blocker remains. Runtime strings and scope collision behavior are environment observations, not an Excel compatibility guarantee.

## 26. Recommended Prompt 12 scope

Build on this typed surface with narrowly selected workbook/worksheet operations or formula-writing APIs only after their typelib signatures and live semantics are independently verified.
