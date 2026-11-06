# Excel C API source-census working report

**Status: independent agent technical review performed; authenticated completeness acceptance remains NOT VERIFIED.** Owner: `evnekdev/excel-api`, [#117](https://github.com/evnekdev/excel-api/issues/117); validator work [#118](https://github.com/evnekdev/excel-api/issues/118). Inspected 2026-11-06. This records source declarations, not Rust implementation, callback legality, ABI layout or real-Excel behavior.

The original implementation report below is retained as working history. Its outstanding independent technical inspection was subsequently performed in the [review addendum](#independent-agent-technical-review-2026-11-06); its claims about universal rejection of unsupported constructs were corrected there. Neither report is human sign-off or an authenticated CLI completeness attestation.

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

These were the limitations recorded at implementation commit `2747f97`; see the dated addendum for their current disposition.

- A reviewer independent of the extractor author must review source sections, candidate categories, exclusions, the counting convention, conditional configurations, signatures and values. The current comparison is not such a review. No full-coverage percentage or closure of #117 is justified.
- The parser supports the forms in this pinned source; `#if`/`#elif`, function-like or multiline macros, repeated definitions, recursive/unknown names and unsupported arithmetic fail explicitly and need review. It does not prove full aggregate field layouts, packing, ABI, all possible preprocessor configurations or externally provided Windows type definitions.
- A reviewed source census format and verifiable review provenance are not yet enforced by the validator. An editable metadata field or `complete: true` cannot establish independence. Reconciliation remains **NOT VERIFIED** even when the provisional sets match.
- Rust assessment and real-Excel testing remain separate; all records remain `unassessed`/`not-run`. Empty EDNA links on new raw declarations mean no relationship was asserted, not absence of a relationship.
- SDK mirror rights and legacy CLR/COM-header provenance remain under [#120](https://github.com/evnekdev/excel-api/issues/120), a native child of #90; #114 is closed. No vendor payload is changed by this inventory. Dalton evidence remains separate under #115.

## Independent agent technical review, 2026-11-06

Reviewed Terra branch `feature/edna02-exhaustive-c-api-inventory` at `2747f97a3636789ba6af1780d71bfb39047c1b9a`, from fetched `origin/master` `d825636b48abbc8ff55413fc7556bba548fa51cf`. Corrections are on separate attached branch `review/edna02-independent-census-audit`; neither original branch nor master is rewritten.

### Method and independence provenance

The coordinating chat already contained Terra implementation history. A fresh agent context (`independent_source_review`) received only the task boundaries and pinned source locator, without that history or expected totals. It inspected the full header, independently accounted for directives/declaration statements, evaluated macros, and recorded initial expectations and nested-member observations in temporary artifacts **before reading the catalogue, extractor, scanner, generator or tests**. After freezing those expectations it was permitted to compare them. It then read the tool README and the first 100 extractor lines while locating the catalogue, before completing comparison; this sequencing deviation is disclosed rather than claiming a wholly blinded comparison. The initial expected set was already fixed.

The retained [review reproducer](../../tools/excel-c-api-inventory/review_source.py) uses source **comment section boundaries** for macro classification, its own linear declaration accounting and recursive-descent integer evaluator, and imports none of Terra's tooling. It compares directly with the JSONL catalogue, including a Git snapshot option. Its method was developed from the initial source observations and retained after the comparison; it is not a cryptographically timestamped or authenticated attestation. Temporary raw-header/member artifacts are not republished. This is an **independent agent technical review**, not independent human certification.

Reproduction against both original and corrected catalogues found no missing/extra identity, wrong kind, source-span, scalar/pointer/tag relationship, macro expression/value, callback signature or linkage. Source hashes above were recomputed (including Git's blob framing), confirming 41,681 bytes and 1,480 LF-terminated lines. A trailing empty split element is not another source line.

### Counting convention and all 38 non-macro identities

Independently observed: **16 typedef statements**, introducing **22 typedef names** (five scalar, ten aggregate, six pointer, one callback); **ten struct tags** and **six prototypes**. Thus 38 non-macro names plus 1,094 macros = **1,132** top-level identities. Tags and typedefs use distinct C namespaces and source spellings; separate records are reasonable for a name inventory, but do not represent distinct storage layouts. Anonymous nested aggregates and named members remain properties of the enclosing type. No standalone enum or named union declaration occurs.

Every row below is **PASS for source identity, kind and anchor**; no identity correction is required. “Aggregate typedef” is the schema's retained `struct` kind. All line numbers refer to the immutable header linked above. Signature relationships describe written types and conventions, not binary ABI equivalence.

| Name | Kind | Lines | Reviewed relationship / source meaning |
| --- | --- | --- | --- |
| BOOL | scalar-typedef | 24 | Underlying INT32; Boolean role |
| XCHAR | scalar-typedef | 25 | Underlying WCHAR; character role |
| RW | scalar-typedef | 26 | Underlying INT32; Excel 12 row |
| COL | scalar-typedef | 27 | Underlying INT32; Excel 12 column |
| IDSHEET | scalar-typedef | 28 | Underlying DWORD_PTR; sheet identifier |
| XLREF | aggregate typedef | 36–42 | Tag xlref; legacy rectangular reference |
| xlref | struct-tag | 36–42 | Tag of XLREF |
| LPXLREF | pointer-typedef | 36–42 | Pointer to XLREF |
| XLMREF | aggregate typedef | 53–57 | Tag xlmref; counted XLREF table |
| xlmref | struct-tag | 53–57 | Tag of XLMREF |
| LPXLMREF | pointer-typedef | 53–57 | Pointer to XLMREF |
| XLREF12 | aggregate typedef | 66–72 | Tag xlref12; RW/COL rectangular reference |
| xlref12 | struct-tag | 66–72 | Tag of XLREF12 |
| LPXLREF12 | pointer-typedef | 66–72 | Pointer to XLREF12 |
| XLMREF12 | aggregate typedef | 83–87 | Tag xlmref12; counted XLREF12 table |
| xlmref12 | struct-tag | 83–87 | Tag of XLMREF12 |
| LPXLMREF12 | pointer-typedef | 83–87 | Pointer to XLMREF12 |
| FP | aggregate typedef | 96–101 | Tag _FP; unsigned-short dimensions and double array |
| _FP | struct-tag | 96–101 | Tag of FP; no pointer typedef here |
| FP12 | aggregate typedef | 109–114 | Tag _FP12; INT32 dimensions and double array |
| _FP12 | struct-tag | 109–114 | Tag of FP12; no pointer typedef here |
| XLOPER | aggregate typedef | 125–177 | Tag xloper; WORD discriminator; conditional Boolean spelling |
| xloper | struct-tag | 125–177 | Tag of XLOPER; self-pointer in array member |
| LPXLOPER | pointer-typedef | 125–177 | Pointer to XLOPER; outer terminator is line 177 |
| XLOPER12 | aggregate typedef | 187–235 | Tag xloper12; DWORD discriminator; unconditional BOOL xbool |
| xloper12 | struct-tag | 187–235 | Tag of XLOPER12; self-pointer in array member |
| LPXLOPER12 | pointer-typedef | 187–235 | Pointer to XLOPER12; outer terminator is line 235 |
| Excel4 | function-prototype | 329 | int, _cdecl; int function number, LPXLOPER result, int count, variadic operands |
| Excel4v | function-prototype | 332 | int, pascal; int function number, LPXLOPER result, int count, LPXLOPER operand array |
| XLCallVer | function-prototype | 334 | int, pascal; explicit void parameter list |
| LPenHelper | function-prototype | 336 | long, pascal; int code then VOID pointer |
| Excel12 | function-prototype | 338 | int, _cdecl; int function number, LPXLOPER12 result, int count, variadic operands |
| Excel12v | function-prototype | 341 | int, pascal; int function number, LPXLOPER12 result, int count, LPXLOPER12 operand array |
| PXL_HPC_ASYNC_CALLBACK | function-pointer-typedef | 352 | int, CALLBACK; LPXLOPER12 handle then LPXLOPER12 result; outside preceding extern-C block |
| FMLAINFO | aggregate typedef | 456–464 | Tag _fmlainfo; edit mode, formula buffer, selection/caret positions |
| _fmlainfo | struct-tag | 456–464 | Tag of FMLAINFO; no pointer typedef here |
| MOUSEINFO | aggregate typedef | 466–477 | Tag _mouseinfo; HWND/POINT inputs and four int outputs |
| _mouseinfo | struct-tag | 466–477 | Tag of MOUSEINFO; no pointer typedef here |

### Aggregate members and conditional semantics

All ten aggregates were inspected, including nested unions/structures, member order, pointers and arrays:

- **XLREF / XLREF12:** first/last rows then first/last columns; WORD/WORD/BYTE/BYTE becomes RW/RW/COL/COL (36–42, 66–72).
- **XLMREF / XLMREF12:** WORD count precedes the matching reference table (53–57, 83–87). The written extent is `[1]`; comments describe count-dependent storage. This is not a C flexible-array declaration.
- **FP / FP12:** rows then columns then double array (96–114). Dimensions change from unsigned short int to INT32. The written trailing extent stays `[1]`; comments describe rows-by-columns storage.
- **XLOPER / XLOPER12:** `val` union precedes `xltype`; discriminator changes WORD to DWORD. Scalar members are double `num`, LPSTR / XCHAR-pointer `str`, WORD / BOOL Boolean, WORD / int `err`, short-int / int `w` (127–137, 189–196). Only legacy Boolean spelling branches: WORD `xbool` under `__cplusplus`, WORD `bool` otherwise (131–135). This is compilation syntax, not Excel-version availability.
- **Single reference:** WORD count, followed by XLREF / XLREF12 `ref`; comment states count always 1 (138–142, 197–200). **Multi-reference:** XLMREF / XLMREF12 pointer then IDSHEET (143–147, 201–205). **Array:** self-struct pointer then WORD/WORD dimensions or RW/COL (148–153, 206–211). Nested `sref` is a member, never the outer typedef terminator; both full spans above were independently checked.
- **Flow:** nested `valflow` union holds level/control as short-int / int and IDSHEET; then WORD/BYTE coordinates become RW/COL, followed by BYTE flow discriminator (154–165, 212–223).
- **Big data:** nested `h` union holds BYTE pointer (input comment) or HANDLE (output comment), followed by long byte count in both types (166–174, 224–232). No platform width or free protocol is inferred.
- **FMLAINFO:** int edit mode and character count, char pointer, then three int offsets for selection start/end and caret (456–464). The buffer is read-only by comment, not a `const` type; mode zero leaves the remaining fields undefined per comment. **MOUSEINFO:** HWND and POINT inputs, then int document type, hit-test code, row and column outputs (466–477); comments use -1 for unavailable coordinates.

The header requires Windows definitions by comment at 9–10; it does not include them itself. Windows INT32, WCHAR, DWORD_PTR, WORD, BYTE, DWORD, LPSTR, HANDLE, HWND, POINT, VOID and calling-convention macros are external contracts. Their applicable compiler/SDK definitions, packing, target sizes, offsets and x86/x64 ABI remain unverified. The catalogue's exact blob/span preserves the source; this compact review supplies member semantics without copying whole SDK declarations into another registry or adding a speculative general C schema. Automated general aggregate-field validation remains a separate limitation.

The six prototypes have conditional C linkage at 325–345; none is excluded in C compilation. Calling convention spelling is preserved separately from linkage. The callback typedef at 352 is a pointer type, not a callback number or exported function.

External contract pointer, inspected 2026-11-06: Microsoft's [Windows Data Types](https://learn.microsoft.com/en-us/windows/win32/winprog/windows-data-types) identifies INT32 in BaseTsd.h, DWORD_PTR as a pointer-precision type, WCHAR in WinNT.h, and CALLBACK in WinDef.h. This identifies the relevant Windows SDK authority; it does not pin the compiler/Windows SDK configuration used to build the 2013 XLL header or prove an aggregate layout. Those target-specific checks remain separate.

### Macro families, expressions and exclusions

Every macro name, exact source line, expression and effective integer was compared: **1,094 checked, zero mismatches**. Categories in the original table were confirmed independently from source comment boundaries. The 48 `other-constant` identities now have source-checked subfamilies without changing `(path, kind, symbol)`:

| Subfamily | Count | Source lines | Interpretation |
| --- | ---: | --- | --- |
| cluster-return-code | 3 | 359–361 | Cluster connector results, distinct from xlret callback returns |
| edit-mode | 4 | 404–407 | Ready/enter/edit/point modes |
| document-type | 5 | 410–415 | Document/page discriminators, including dtNil's special comment |
| mouse-hit-test | 35 | 418–454 | Mouse-query hit areas; includes internal-use codes |
| user-defined-function-id | 1 | 487 | UDF dispatch number; first argument is a function reference |

The 20 retained `special-function-id` records are auxiliary **selectors**, not uniformly Excel4/Excel12-callable operations. Lines 395–397 restrict xlGetFmlaInfo/xlGetMouseInfo to LPenHelper. xlGetInst's 32-bit limitation and xlGetInstPtr's 32/64-bit description are source comments, not preprocessor guards. `worksheet-function-id` is retained as a stable historical kind but means the combined Excel/**XLM** table (497–1075), not a claim that every entry is a modern cell function. `xlflow*` remains separate from `xlf*`; no other prefix collision was found in this pinned set.

There are **424 OR expressions**: xltypeBigData, 20 auxiliary selectors and 403 commands. All remaining expressions are supported numeric literals, including the two negative cluster results. Examples from distinct families include xltypeBigData = 2050, xlCommand = 32768, xlFree = 16384, xlcOpen = 32769, xlHpcRetCallFailed = -2, dtNil = 127 and htEdit = 34. All expressions, not only these examples, were independently evaluated. There are **no direct single-name aliases**, **46 repeated-value groups containing 129 names**, and no same-family collision under the independent source-comment taxonomy. Numeric collisions across families are not aliases. No gaps were filled with invented symbols.

No active function-like/multiline/duplicate-name macro, disabled declaration, macro conditional, enum, include/undef directive or unknown top-level syntax occurs in this source. Explicit non-name exclusions: pragma once (18); three ifdef, one else and three endif directives (131–135, 325–327, 343–345); conditional linkage braces; comments; nested member identities; externally defined Windows declarations. No in-scope top-level candidate is excluded or unresolved. This exhaustively reconciles the **defined pinned name-accounting scope as technical review evidence**, not all C semantics or an authenticated CLI completeness result.

### Substantiated corrections and regression evidence

Pre-fix regressions reproduced on Terra's code before corrections:

1. Appended extern/enum/multiword typedef declarations, include/undef/packing directives and a second declaration on a typedef line were silently omitted. Character-span accounting now rejects unaccounted top-level source. The reviewer also found malformed `#define` directives were skipped; these now fail explicitly. Aggregate bodies remain source-anchored and manually reviewed, not a general C grammar.
2. An added conditional around BOOL lost its guard. Supported guards now accompany typedef/tag/pointer/callback/prototype candidates as well as macros. A multiline scalar typedef also anchored only its final name line; the regression now requires its entire declaration span. Multiline start/end handling is corrected for supported typedef/prototype forms without changing any pinned-source anchor.
3. An exact but unapproved commit containing the same header blob passed. An explicit extendable `APPROVED_SOURCES` tuple allowlist now rejects it. A missing approved Git object returns NOT VERIFIED / exit 3; an unapproved tuple is invalid / exit 1. Future versions require a reviewed source tuple and regression evidence; no permanent single-version assumption or ID ledger is introduced.
4. Python-only numeral spellings (`1_000`, `0b10`, `0o10`) passed the AST evaluator. The source literal grammar now rejects them before safe AST interpretation. The pinned source needs only decimal/hexadecimal literals, unary signs and OR; signedness and overflow semantics remain outside this evaluator.
5. Contradictory top-level census source metadata was ignored. Supplied repository/revision/path/blob and SHA-256 now undergo consistency checks. Consistent metadata still cannot authenticate a reviewer or establish completeness.

All 1,117 former generic descriptions are replaced with source-grounded descriptions through `populate.py`; the original 15 descriptions and every original field are preserved. Precise types, return/flow/event roles, auxiliary restrictions and all five miscellaneous subfamilies are explained. Historical function/command descriptions remain conservative about availability and argument semantics. The header's inconsistent row/column wording for htRwShow/htColShow is explicitly retained as uncertainty, not silently corrected as behavior. Regeneration upgrades only the exact previous generic template; any curated description, EDNA list or evidence survives. No CAPI ID, kind or source anchor changes; no schema migration is needed. `subfamily` is additive optional v2 detail and any asserted value is source-checked.

### EDNA assessment and remaining acceptance

The 15 original EDNA mappings are preserved. New empty lists still mean **unassessed relationship**, not “unrelated.” The architecture register supports these concrete mapping candidates: value/FP types → EDNA-02/03; reference types → EDNA-02/05; Excel4/12 invocation → EDNA-02/06; ownership bits/xlFree → EDNA-04; xlfRegister/registration identifiers → EDNA-07; native async return → EDNA-11; event identifiers/registration → EDNA-14; cluster callback/results → EDNA-10 subject to scope review. These are relationship recommendations, not assertions of Excel-DNA usage or Rust implementation. No blanket per-function or per-command mapping is applied. A follow-up should record each selected relationship with the architecture section and, when asserting actual upstream usage, a pinned Excel-DNA call site; assess mappings in cohesive families and preserve curated lists during regeneration.

| Verification dimension | Result |
| --- | --- |
| Source integrity and all name anchors | PASS against approved pinned Git source |
| Mechanical enumeration | PASS for defined top-level names and supported numeric expressions |
| Independent technical inspection | PERFORMED by fresh agent context, with sequencing caveat above; not human sign-off |
| Claimed literal typedef/prototype/conditional relationships | PASS within the documented source scope |
| Full C semantics / aggregate ABI / external platform contracts | NOT VERIFIED |
| Authenticated exhaustive census acceptance in CLI | NOT VERIFIED; matching reconciliation remains exit 3 |
| Rust implementation / real Excel / modern availability | NOT ASSESSED / NOT RUN |

**Merge assessment: READY AS PROVISIONAL INFRASTRUCTURE.** No observed name/number/identity blocker remains in the defined source scope. Keep #117 open until the maintainer accepts this agent review provenance/counting scope or records any additionally required independent review and completeness attestation method. Keep #118 open for structured host evidence and assessed-record freshness; #120 remains open for source rights. ABI/runtime expansion is not a blocker to source-only infrastructure, and no payload or production Rust modification is proposed. #90 and #78 remain open.

### Executed verification

Windows, Python 3.11.7; `PYTHONDONTWRITEBYTECODE=1`, `PYTHONIOENCODING=utf-8`. `<temp-census>` below was `%TEMP%/excel-c-api-independent-review-census.json`, not a committed artifact. An initial subprocess-output capture used the wrong encoding for Windows punctuation; it was rerun with explicit UTF-8. This was a capture issue, not a catalogue/CLI failure.

| Exact command / check | Result |
| --- | --- |
| `python -m unittest discover -s tools/excel-c-api-inventory -p "test*.py" -v` | PASS: 50 tests (32 retained, 18 new) |
| `python -m unittest discover -s tools/excel-c-api-inventory -p test_review_regressions.py -v` | PASS: 9 targeted regressions, including CLI provenance/spoofed-review checks |
| `python tools/excel-c-api-inventory/validate.py check` | PASS: 1,132 records, exit 0 |
| `python tools/excel-c-api-inventory/validate.py freshness` | Integrity PASS, exit 0; completeness/Rust/host/EDNA assessments NOT ASSESSED |
| `python tools/excel-c-api-inventory/validate.py reconcile` | NOT ASSESSED, exit 3 |
| `python tools/excel-c-api-inventory/source_census.py --output <temp-census>` | PASS, exit 0 |
| `python tools/excel-c-api-inventory/audit_source.py <temp-census>` | Mechanical check PASS, exit 0; no source-only lines or differences |
| `python tools/excel-c-api-inventory/validate.py reconcile --census <temp-census>` | Matching supplied census; SOURCE COMPLETENESS NOT VERIFIED, exit 3 |
| `python tools/excel-c-api-inventory/review_source.py --catalogue-revision 2747f97` | PASS, exit 0; zero differences against frozen original catalogue |
| `python tools/excel-c-api-inventory/review_source.py --require-subfamilies` | PASS, exit 0; zero differences against enriched catalogue |
| `python tools/excel-c-api-inventory/populate.py --write` | PASS; 1,132 records, zero new IDs; subsequent byte-for-byte build comparison unchanged |
| `git diff --check` | PASS |
| `git diff --exit-code 2747f97 -- tools/Excel2013XLLSDK docs/evidence/native-xll-header-provenance.md` | PASS: source payload and prior header inventory unchanged |

Synthetic CLI fixtures returned 3 for complete/unverified and partial matches; 2 for contradictory source provenance, omitted census entry (extra catalogue) and extra eligible census entry (missing catalogue). Existing negative tests verify invalid catalogue exit 1, missing history exit 3, duplicate/reassigned/removed/reused IDs, source anchors, values, EDNA references and host claims. New tests exercise independent missing/extra names, type/tag/pointer/callback/signature/conditional/classification mutations, false aliases, unsupported source/literals, missing approved source, curated-text preservation and spoofed reviewer metadata. Each newly corrected failure was reproduced before its fix where described above.

All 1,132 identities and source objects match the original branch; only 1,117 descriptions and 48 additive subfamilies differ in JSONL. The original 15 complete records are unchanged. Local Markdown file links and full scoped diff were checked. CI, Cargo, native Excel, UI automation, installer/product builds and ABI acceptance: **NOT RUN**, explicitly outside this source-review task.
