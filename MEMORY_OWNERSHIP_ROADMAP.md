# Memory and Ownership Implementation Roadmap

## Objective

Build and validate the translation layer between `XLOPER12` and Rust without leaks, double frees, use-after-free, invalid union access, incorrect allocator use, or callback-lifetime violations.

## Phase 1 — ABI verification

Deliverables:

- compare supported structures with official SDK headers;
- add C-versus-Rust size, alignment, and offset checks;
- verify tags, ownership masks, calling conventions, and supported target ABI;
- document constant/header provenance.

Exit: every supported raw definition is verified; placeholders are not used by safe code.

## Phase 2 — Borrowed inputs

Deliverables:

- `RawExcelValue<'call>` and `ExcelValueRef<'call>`;
- borrowed UTF-16 string and flat-array views;
- checked tag, pointer, length, dimension, and overflow validation;
- explicit safety limits;
- unit tests and parser fuzzing.

Exit: supported callback inputs parse without allocation and cannot escape in safe Rust.

## Phase 3 — Owned semantic values

Deliverables:

- complete `ExcelValue` and `ExcelArray` invariants;
- deep-copy conversions;
- strict/lossy UTF-16 APIs;
- exact missing/empty handling;
- worker-safe owned input flow.

Exit: copied values remain valid after the callback and may safely move to workers.

## Phase 4 — Logical returns and planning

Deliverables:

- `ExcelReturnValue`;
- `IntoExcel` producing logical values, not pointers;
- `ReturnPlan` with checked counts and byte sizes;
- rejection of unsupported/nested outputs;
- aggregate-memory limits.

Exit: every supported result is fully validated before raw allocation begins.

## Phase 5 — Stable allocation

Deliverables:

- offset-zero `ReturnAllocation` root;
- stable boxed string and array buffers;
- pointer patching after final addresses exist;
- `ExcelReturn` RAII owner;
- failure injection for each allocation stage.

Exit: dropping any unhanded return frees all backing storage; no published pointer can move.

## Phase 6 — Handoff and cleanup

Deliverables:

- consuming `into_raw_for_excel`;
- centralized DLL-free flag application;
- minimal panic-safe `xlAutoFree12`;
- debug header and live-allocation counters;
- tests for exactly-once handoff/free.

Exit: every handed-off complex return has one matching cleanup, including panic/error stress paths.

## Phase 7 — Excel-owned API results

Deliverables:

- `ExcelOwnedValue` and explicit release policies;
- per-call ownership documentation;
- correct release on success, conversion failure, and panic;
- thread restrictions where required.

Exit: repeated API calls leak nothing and never use the XLL-return allocator.

## Phase 8 — Integration and stress

Exercise:

- repeated scalar, string, and mixed-array recalculation;
- empty and maximum supported strings;
- workbook open/close and add-in load/unload;
- multi-threaded recalculation;
- conversion errors and deliberate panics;
- cancellation and shutdown;
- high-volume handoff/free;
- malformed parser fuzz corpus.

Use atomic allocation accounting, failure injection, Miri for pure-Rust helpers, supported sanitizers, Windows heap diagnostics, and real Excel stress workbooks.

Exit: zero live allocations after cleanup, no invalid-heap diagnostics, and all ownership transitions reviewed.

## Suggested PR sequence

1. `verify/xloper12-abi`
2. `feature/borrowed-excel-values`
3. `feature/owned-excel-values`
4. `feature/return-planning`
5. `feature/stable-return-allocation`
6. `feature/xlautofree12`
7. `feature/excel-owned-api-results`
8. `test/memory-stress-harness`
