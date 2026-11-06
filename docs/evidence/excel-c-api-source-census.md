# Excel C API source-census working report

**Status: provisional; independent review and exhaustive coverage NOT ASSESSED.** Owner: `evnekdev/excel-api`, [#117](https://github.com/evnekdev/excel-api/issues/117); validator work [#118](https://github.com/evnekdev/excel-api/issues/118). Inspected 2026-11-06. This records source declarations, not Rust implementation, callback legality, ABI layout or real-Excel behavior.

## Source and method

The source is Microsoft Excel 2013 XLL SDK, header 15.0, [`tools/Excel2013XLLSDK/INCLUDE/XLCALL.H`](https://github.com/evnekdev/excel-api/blob/6c8354d6536e18b58559f9cc3134f25a9e58d478/tools/Excel2013XLLSDK/INCLUDE/XLCALL.H), commit `6c8354d6536e18b58559f9cc3134f25a9e58d478`, Git blob SHA-1 `4eb50e8fd3aea40d283181b7dcc74871ba88d854`, content SHA-256 `d69f860ca5f7be9779f69423276808845ebb2f9648e38aadc99908277a710c74`. The blob contains **1,480 physical LF-terminated lines**, not the 1,481 stated in the task brief.

`source_census.py` reads the Git blob without consulting the catalogue. It extracts the supported object-like macros, typedef statements and prototypes, records exact line spans and evaluates macro expressions with a restricted AST walker. It reports an error for unsupported constructs or expressions rather than silently skipping them. `audit_source.py` uses a separate comment state machine, line/statement scan and recursive-descent expression evaluator to cross-check the extracted names, spans and values. It found 1,094 macro directives, 16 typedef statements, six prototype statements, no unaccounted non-comment source lines, and zero lexical/value discrepancies. Its first pass incorrectly chose a nested `sref` as the outer typedef of both XLOPER structs; anchoring the outer terminator fixed that discrepancy. These are **complementary mechanical methods authored in the same run**, not independent human review and not evidence of exhaustive C semantic correctness.

Reproduce with the commands in the [tool README](../../tools/excel-c-api-inventory/README.md). The generated JSON is a source-candidate census without CAPI IDs, not a second identity registry. Its `complete` flag requests a comparison of the full extracted set; it cannot certify source completeness. The existing reconciler returns exit 3 even when it matches.

## Provisional accounting

A *candidate* is a distinct source-visible top-level macro, typedef name, struct tag or function prototype, with one CAPI identity per `(path, kind, symbol)`. Ten aggregate typedef statements each supply a typedef name and a distinct struct tag; six additionally supply a named pointer typedef. Nested struct/union fields, including variable-length trailing-array members, are part of their enclosing aggregate and are not counted as separate top-level candidates. The XLOPER C/C++ conditional members are recorded as variants of the enclosing type. C and C++ function linkage is recorded separately from calling conventions. A pointer typedef is not merged with the pointee.

| Candidate category | Source candidates | Matched catalogue records | Excluded | Parser unresolved |
| --- | ---: | ---: | ---: | ---: |
| Scalar typedef | 5 | 5 | 0 | 0 |
| Pointer typedef | 6 | 6 | 0 | 0 |
| Aggregate typedef | 10 | 10 | 0 | 0 |
| Struct tag | 10 | 10 | 0 | 0 |
| Function-pointer typedef | 1 | 1 | 0 | 0 |
| Function prototype | 6 | 6 | 0 | 0 |
| Value-type flag | 12 | 12 | 0 | 0 |
| Ownership flag | 2 | 2 | 0 | 0 |
| Error code | 8 | 8 | 0 | 0 |
| Flow code | 5 | 5 | 0 | 0 |
| Return code | 11 | 11 | 0 | 0 |
| Event ID | 2 | 2 | 0 | 0 |
| Function-number modifier | 4 | 4 | 0 | 0 |
| Special callback ID | 20 | 20 | 0 | 0 |
| Worksheet/XLM function ID | 579 | 579 | 0 | 0 |
| XLM macro command ID | 403 | 403 | 0 | 0 |
| Other constants | 48 | 48 | 0 | 0 |
| **Total** | **1,132** | **1,132** | **0** | **0** |

The equation is candidate names = matched records + excluded candidate names + parser-unresolved candidate names, within this provisional scanner scope. Zero parser-unresolved means the implemented narrow syntax was handled; independent source and category review is still pending for the set as a whole. Of the 1,132 IDs, 15 retain their original IDs and anchors and 1,117 are newly assigned. The 1,094 macro candidates comprise the flag, code, ID and other-constant categories. In particular, the five `xlflow*` codes are not `xlf*` worksheet functions even though their spelling starts with `xlf`.

Structural preprocessing directives (`#pragma once`, three `#ifdef`, one `#else`, three `#endif`) are accounted for but excluded from declaration-name counts. Comments and nested member names are not top-level candidates. The `#ifdef __cplusplus` branches for XLOPER's `WORD xbool` and `WORD bool` are recorded; function prototypes retain conditional C linkage. No macro is inside a conditional branch in this pinned header. `BOOL`, `XCHAR`, `RW`, `COL` and `IDSHEET` are header typedefs whose underlying names refer to external platform types; no independent Windows ABI width is inferred. `FP`/`FP12`, `XLREF`/`XLREF12`, `XLMREF`/`XLMREF12`, and `XLOPER`/`XLOPER12` remain separate declarations.

No direct single-name macro aliases were found. The 1,094 macros form 46 groups with repeated numeric values, containing 129 names across those groups; equal numbers alone are **not** aliases. The composite `xltypeBigData` and IDs containing `xlSpecial`/`xlCommand` retain their source expressions and checked integer results. Signedness beyond the written declaration and C implementation details are not inferred.

## Unresolved review and limits

- A reviewer independent of the extractor author must review source sections, candidate categories, exclusions, the counting convention, conditional configurations, signatures and values. The current comparison is not such a review. No full-coverage percentage or closure of #117 is justified.
- The parser supports the forms in this pinned source; `#if`/`#elif`, function-like or multiline macros, repeated definitions, recursive/unknown names and unsupported arithmetic fail explicitly and need review. It does not prove full aggregate field layouts, packing, ABI, all possible preprocessor configurations or externally provided Windows type definitions.
- A reviewed source census format and verifiable review provenance are not yet enforced by the validator. An editable metadata field or `complete: true` cannot establish independence. Reconciliation remains **NOT VERIFIED** even when the provisional sets match.
- Rust assessment and real-Excel testing remain separate; all records remain `unassessed`/`not-run`. Empty EDNA links on new raw declarations mean no relationship was asserted, not absence of a relationship.
- SDK mirror rights and legacy CLR/COM-header provenance remain under [#120](https://github.com/evnekdev/excel-api/issues/120), a native child of #90; #114 is closed. No vendor payload is changed by this inventory. Dalton evidence remains separate under #115.
