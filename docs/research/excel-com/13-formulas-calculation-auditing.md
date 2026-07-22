# Formula models, calculation, auditing, and search

## 1. Scope and baseline

This work started from `origin/master` commit `80cfea786fb2da562b2d5b1583750c174843f551`. It extends the experimental Windows-only `excel-com` slice with Excel-owned formula authoring, calculation, dependency inspection, SpecialCells, Find, and Replace. The tested environment is Windows 10 Enterprise 25H2 build 26200.8875, Excel `16.0.20131.20154` x64, typelib `{00020813-0000-0000-c000-000000000046}` version 1.9, Rust 1.97.1, and `windows-sys` 0.61.2.

The registered typelib, current Excel runtime, and official Microsoft documentation are authoritative. Microsoft’s raw C++ sample was also reviewed as a low-level reference for `IDispatch::Invoke`, property puts, reversed positional arguments, and a 1-based two-dimensional `SAFEARRAY(VARIANT)`: [How to automate Excel from C++ without using MFC or #import](https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/automate-excel-from-c).

## 2. Formatting-module refactor

`formatting.rs` was inspected but not split. It remains a coherent implementation of the Prompt 12 formatting primitives, has no formula dependency, and splitting it would create unrelated churn without improving this formula-focused change. Its public behavior and prior tests are retained.

## 3. Formula result model

`FormulaValue` preserves the semantic shapes Excel exposes: `Text(String)`, `Array(AutomationArray)`, `Mixed`, and `Empty`. It deliberately distinguishes an empty formula result from an empty-string formula and rejects dispatch results or unsupported scalar variants. Formula text is never parsed or normalized in Rust.

## 4. Formula and Formula2 consistency

`Formula` and `Formula2` now share the same result model and exact-shape policy. Scalar setters only accept a 1x1 Range; array setters accept only a same-shaped rectangular `AutomationArray` of text formulas. `Formula2` is passed directly to Excel and therefore retains dynamic-array and implicit-intersection semantics described in [Range.Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2).

## 5. FormulaR1C1

`Range::formula_r1c1`, `set_formula_r1c1`, and `set_formula_r1c1_array` use `Range.FormulaR1C1` directly. A live `=RC[-3]*2` case read back unchanged and calculated to `4`; relative-reference interpretation stays in Excel.

## 6. Formula2R1C1

The installed typelib contains `Formula2R1C1`, and the wrapper exposes getter, scalar setter, and exact-shape array setter. A live `=SEQUENCE(2,2)` case spilled to a 2x2 range. This member is marked version-sensitive because dynamic-array support depends on the installed Excel build.

## 7. Locale formula members

`FormulaLocal` and `FormulaR1C1Local` are exposed as clearly locale-dependent scalar APIs. In the tested `en-AU` culture with `en-GB` UI culture, separator-free `=1+1` read back identically. Rust does not translate function names, decimal separators, list separators, or reference text; broader locale syntax remains intentionally unclaimed.

## 8. Formula presence

`Range::has_formula` returns `MixedValue<bool>`: uniform `true` for formula-only input, uniform `false` for constants or blanks, and `Mixed` for heterogeneous cells. Excel `Null` remains `Mixed`; no default Boolean is invented.

## 9. Legacy array formulas

`has_array`, `current_array`, `formula_array`, and `set_formula_array` are transparent wrappers for Excel’s legacy array-formula surface. The live `=A6:A8*B6:B8` array filled `C6:C8`, reported `HasArray`, returned `CurrentArray`, and calculated `10`, `40`, and `90`. `FormulaArray` is distinct from ordinary rectangular Formula assignments and retains Excel’s documented legacy 255-character limit.

## 10. Dynamic-array spills

`has_spill`, `spilling_to_range`, and `spill_parent` delegate to the installed dynamic-array members. `=SEQUENCE(2,2)` reported origin `F1`, spill range `F1:G2`, and spill-parent `F1` for a child. A deliberately occupied output cell produced an Excel error value; after clearing the blocker, `Range::calculate` restored the spill. No spill range is inferred from address arithmetic. See [Range.SpillParent](https://learn.microsoft.com/en-us/office/vba/api/excel.range.spillparent).

## 11. Calculation mode

`CalculationMode(i32)` is a transparent, forward-compatible `XlCalculation` representation with Automatic, Manual, and Semiautomatic constants. The getter rejects an unexpected `VT_ERROR` result instead of treating an Excel error SCODE as an enum. In this environment that error occurred before a workbook was added, so the controlled test obtains calculation state after `Workbooks.Add`.

## 12. Calculation state

`CalculationState(i32)` preserves known and unknown `XlCalculationState` values. The small controlled workbook observed `Done` after Application calculation, full calculation, and full rebuild. This is an observation for the tested workbook, not a promise of synchronous completion in all Excel workloads.

## 13. Calculation guards

`Application::calculation_mode_guard` follows the established alert/reference-style guard pattern: it captures the prior global mode, applies the requested mode, provides fallible explicit `restore`, best-effort restores in `Drop`, and disarms after restoration. It never panics from `Drop`.

## 14. Calculation scopes

`Application::calculate`, `calculate_full`, and `calculate_full_rebuild`, plus `Worksheet::calculate` and `Range::calculate`, invoke the respective Excel operations. In Manual mode, a changed dependent remained stale at `20` until `Range::calculate` produced `24`; Worksheet and Application calculation then produced `26` and `28`. Full calculation and full rebuild run once each because rebuild can be expensive. [Application.CalculateFullRebuild](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild) documents the rebuild operation.

## 15. Dirty cells

`Range::mark_dirty` wraps `Range.Dirty`. It marks cells for Excel recalculation and does not promise to calculate them itself. The live Manual-mode case established that it is callable; the subsequent Application calculation produced the expected dependent value.

## 16. Precedents and dependents

`direct_precedents`, `direct_dependents`, `precedents`, and `dependents` return typed Range wrappers and preserve Excel errors. For a same-sheet chain, direct precedents were `A12`, transitive precedents `A12:B12`, direct dependents `B12`, and transitive dependents `B12:C12`, compared as qualified external addresses rather than COM identity. A cross-sheet `DirectPrecedents` call returned structured `DISP_E_EXCEPTION` / `0x800A03EC`; the wrapper preserves this version-sensitive result rather than attempting graph construction.

## 17. SpecialCells

`SpecialCellType` and bitwise `SpecialCellValueMask` expose the curated typelib constants. `Range::special_cells` preserves a multi-area Range and preserves Excel’s no-match error. The live test found formulas, text-or-number constants, blanks, and visible cells. A blank multi-cell no-match receiver returned `DISP_E_EXCEPTION` / `0x800A03EC`; a multi-cell receiver is used because Excel can expand one-cell `SpecialCells` searches beyond that cell.

## 18. Find

`FindLookIn`, `FindMatchMode`, `SearchOrder`, and `SearchDirection` are transparent forward-compatible types. `FindOptions::default` sends concrete values for every remembered Find setting, avoiding accidental inheritance from Excel UI state. The API accepts text and finite numeric values, rejects arrays and objects before COM, and maps no match to `Ok(None)`.

## 19. FindNext and FindPrevious

`find_next` and `find_previous` preserve Excel’s stateful search model and optional `After` Range. The live test exercised both after a text Find. Callers should start with `find`; these methods intentionally do not pretend that their state is independent.

## 20. Find-all iterator

`RangeFindIter` starts one typed Find, advances with `FindNext`, records normalized external addresses, stops before a duplicate address is emitted, and fuses on terminal failure. This avoids relying on canonical COM identity, which is unsuitable for equivalent Excel Range objects. Two text matches were emitted and wraparound terminated safely.

## 21. Replace

`ReplaceOptions::default` sends concrete search and replacement-format controls while retaining the trailing FormulaVersion position as Missing. `Range::replace` accepts text, number, Boolean, and Excel-error scalars, returns Excel’s Boolean result, and rejects arrays/objects before COM. The live test replaced text values and changed formula text from `SUM` to `AVERAGE`.

## 22. Physical VARIANT observations

The public formula layer deliberately exposes `FormulaValue`, not raw `VARIANT`. Its decoder handles BSTR text, `VT_ARRAY | VT_VARIANT` arrays, `VT_NULL` mixed values, and `VT_EMPTY`; dispatch and unrelated scalar values are rejected. Functional live coverage validates scalar and rectangular semantic results. A dedicated raw runtime VARTYPE capture is intentionally deferred rather than leaking private Automation representation into the public API or evidence.

## 23. Version-dependent behaviour

`Formula2`, `Formula2R1C1`, `HasSpill`, `SpillingToRange`, and `SpillParent` are marked version-sensitive in object-model metadata. Cross-sheet dependency results and locale-specific formula syntax are also treated as environment dependent. No late-bound fallback is fabricated if a registered typelib lacks a member.

## 24. Live results

The ignored visible test creates a fresh unsaved workbook only after confirming no `EXCEL.EXE` process exists. It passes Formula/Formula2, R1C1, local separator-free formulas, exact-shape rejection, formula presence, legacy arrays, dynamic spills and blocked-spill recovery, Manual calculation and all scopes, Dirty, same-sheet auditing, cross-sheet structured error handling, SpecialCells, Find/FindNext/FindPrevious/find-all, Replace, state restoration, close-without-saving, explicit Quit, and natural process exit with no remaining Excel process.

## 25. Explicit non-decisions

This work does not add a formula parser, AST, calculation engine, dependency graph, events, tables, charts, generic late binding, external-link management, query waiting, or raw COM pointers. It does not attach to an existing Excel session, alter Office registration/security/bitness, or make wrappers `Send`/`Sync`.

## 26. Rustdoc

All new public formula, calculation, SpecialCells, search, replacement, and iterator types have Rustdoc. Crate documentation now includes a “Formula, calculation, and auditing” section and compiling `no_run` examples covering Formula2 spills, scoped calculation, SpecialCells, and Find iteration.

## 27. Inventory

The inventory maps each implemented calculation and Range member to its exact normalized ID, typelib DISPID, and selected invocation kind. Range records formula and auditing/search capability blocks; Application, Worksheet, and Range record calculation capability blocks. Metadata marks mixed formula values, version-sensitive dynamic-array members, nullable Find results, and stateful search members.

## 28. Validation

The final validation set passed formatting, Clippy, workspace tests and doctests, strict Rustdoc, extraction/generation/check/diff for the inventory, the knowledge-base checker, all prior ignored Excel tests, and the new formula/calculation live test. The generated inventory was regenerated from the registered typelib after the metadata changes.

## 29. Remaining blockers

There is no implementation blocker. Deliberate evidence boundaries remain: no raw runtime formula VARTYPE capture, no cross-workbook dependency traversal, and no claim about locale-specific function names or separators beyond the tested separator-free literal. Cross-sheet DirectPrecedents is documented as an Excel structured-error result in this environment.

## 30. Recommended Prompt 14 scope

Prefer a bounded next slice that builds on the now-typed formula/search foundation, such as conditional formatting or data validation, only after separately inspecting typelib support and Excel runtime behavior. Keep calculation, dependency, and locale behavior Excel-owned; do not broaden this crate into a formula engine or generic Automation framework.
