# Internal Excel Automation value and conversion layer (Prompt 06)

## 1. Scope and baseline

This implementation branch starts at `origin/master` `8627b81077d796c4e879d0a7bfd6fdbc15f488c0`. It adds crate-internal codecs to `tools/excel-com-range-probe`; it does not add a stable Excel API or wrapper types.

## 2. Research findings consumed

Prompt 05H established scalar normalization and rank-two `SAFEARRAY(VARIANT)` reads. Prompt 05I bounded the client-wrapper observations. Prompt 05J established that a physical Excel error is a full signed SCODE, not a short worksheet error number.

## 3. Layer boundaries

The semantic layer depends on the raw `windows-sys` owners. The raw layer contains no semantic import. The only added transport seam accepts and returns raw owned variants, so the semantic live suite can encode and decode without an Excel-specific interface declaration.

## 4. Semantic scalar model

`AutomationValue` preserves Empty, Null, Bool, Number, Text, Error, OaDate, Currency, and Array. Missing is `AutomationArgument::Missing`, deliberately outside worksheet values. Blank cells, empty strings, and cleared cells are not collapsed into an `Option`.

## 5. Excel error representation

`ExcelError(i32)` is lossless. The traditional errors use their signed `0x800Axxxx` SCODEs; unknown negative SCODEs also round-trip. Encoding rejects positive short error numbers such as 2042, preventing the invalid direct-raw representation identified in Prompt 05J.

## 6. OA date representation

`OaDate` stores a finite serial exactly as an `f64`. `DateVariant` writes non-negative values as `VT_DATE`; negative values fail conversion. `Value2Serial` writes `VT_R8`, the explicit supported route for negative OA numeric storage.

## 7. Currency representation

`Currency(i64)` is the COM `CY` scaled integer with fixed scale 10,000. Decimal construction uses checked arithmetic, and no conversion through `f64` is implicit.

## 8. Rectangular arrays

`AutomationArray` has explicit rows, columns, and row-major storage. Rank-one input has to be constructed as `row` or `column`; flat input has no inferred orientation. Decoding maps observed SAFEARRAY physical dimension 1 to rows and dimension 2 to columns, regardless of lower bounds.

## 9. Conversion policy

The strict default rejects non-finite values, embedded NUL, loss of integer precision, unsupported arrays, and shape mismatch. `ShapePolicy::Exact` is the only behavior exposed.

## 10. Conversion errors

Pre-COM conversion errors are structured and separate from activation, IDispatch, EXCEPINFO, Excel, and process-cleanup failures. Array errors include a row and column when construction or decoding fails.

## 11. Raw-to-safe decoding

The decoder supports `VT_EMPTY`, `VT_NULL`, `VT_BOOL`, integer and float numeric tags, `VT_BSTR`, `VT_ERROR`, `VT_DATE`, `VT_CY`, and rank-two `VT_ARRAY|VT_VARIANT`. BSTR uses its SDK length and is converted from UTF-16 without exposing pointers.

## 12. Safe-to-raw encoding

The encoder creates initialized `VARIANT`s with exact SCODE, `VT_DATE`, `VT_CY`, or `VT_R8` policy behavior. Text rejects embedded NUL and a cell-length overrun. `Missing` becomes `VT_ERROR / DISP_E_PARAMNOTFOUND`, distinct from `#N/A`.

## 13. SAFEARRAY mapping

Encoding creates a rank-two one-based `SAFEARRAY(VARIANT)` and inserts every initialized element through SDK APIs. Decoding observes bounds, checks count overflow, reads each copied element through the raw owner, then normalizes to zero-based semantic indexing.

## 14. Missing arguments

Missing has no `AutomationValue` variant. This prevents an optional-invocation marker from becoming a worksheet error or an array element.

## 15. Deterministic tests

The range-probe crate has deterministic scalar and array codec coverage, including all standard SCODEs, an arbitrary unknown SCODE, Unicode, non-one bounds, precision failures, negative-date policy, exact CY values, and RAII cleanup on partial array creation. Small property-style loops sample finite bit patterns and currency extrema without opening Excel.

## 16. Live compatibility

`automation-value-live` is opt-in and only uses raw L-mode. It refuses a nonzero pre-existing Excel process count, writes through the semantic encoder, decodes the read-back value, clears, closes, quits, and checks natural exit. Before Office repair, the zero-process suite reached `Workbooks.Add` for every case but Excel returned `DISP_E_EXCEPTION` (`0x80020009`) before any write; the detailed scalar control supplied inner SCODE `0x800A03EC`. After repair and cold reboot, PowerShell Automation successfully created a workbook, while direct raw L/S and standalone high-level Rust IDispatch controls still returned the same Excel SCODE. Every owned process exited naturally.

## 17. Explicit non-decisions

This work does not add public Application, Workbook, Worksheet, or Range wrappers; generated Excel bindings; event sinks; retry policy; existing-session attachment; calendar interpretation; or permissive array coercion.

## 18. Validation

The deterministic range-probe crate test and Clippy checks pass after the implementation. Full workspace and cross-tool validation remains part of final Prompt 06 verification.

## 19. Remaining blockers

The current L-mode raw smoke and full semantic suite are blocked at `Workbooks.Add` by inner Excel SCODE `0x800A03EC`, so neither can establish live codec compatibility. The Office repair corrected the prior general COM document-operation failure: PowerShell Automation now creates a workbook successfully. The remaining post-repair condition is narrower—direct raw and high-level Rust IDispatch controls fail while PowerShell's COM binder succeeds. All attempts began with zero Excel processes and exited without a remaining process. The next step is a bounded direct-IDispatch binding differential, not another Office repair or a semantic-codec change.

## 20. Prompt 06A follow-up (historical evidence retained)

Prompt 06A retains the historical observations above and records later evidence separately in `knowledge/excel-object-model/microsoft-cpp-port/`. The independent native C++ and `windows-sys` Rust controls each completed 20 fresh visible Excel runs, then the unchanged raw scalar smoke and isolated semantic live suite completed. This establishes a later successful state but does not explain or erase the earlier `0x800A03EC` observation.

Generic BSTR conversion now permits text above Excel's 32,767-character cell limit; the Excel range-write boundary alone enforces that limit. Embedded NUL remains rejected. The live observer used for Prompt 06A writes only its new evidence file and does not modify this Prompt 06 historical record.

## 21. Recommended Prompt 07 scope

Prompt 07 can consume these internal codecs through a narrowly designed, non-public transport façade, after reviewing live compatibility evidence and deciding the public-API boundary independently.
