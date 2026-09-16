# `excel-com` migration support contract

Status: architecture note and implementation target. This document defines the reusable Excel Automation capabilities required by workbook-migration consumers such as PyroApp. It does **not** make PyroApp migration semantics part of `excel-com`.

`excel-com` remains a generic Windows desktop Excel COM Automation crate. The motivating migration use case is important because it exercises workbook fidelity, dynamic formulas, names, legacy arrays, and controlled private Excel sessions, but the public API added here must remain generally useful.

## Responsibility boundary

`excel-com` owns:

- COM apartment initialization and ownership;
- a crate-owned/private Excel application instance;
- typed Excel object-model wrappers;
- `IDispatch` invocation policy;
- BSTR, `VARIANT`, `SAFEARRAY`, optional/missing arguments, HRESULT, and EXCEPINFO handling;
- Excel-version compatibility behavior at the wrapper boundary;
- explicit and testable Excel shutdown semantics;
- real desktop Excel evidence for supported migration-critical members.

A migration consumer owns:

- application-specific formula interpretation;
- application-specific rewrite rules;
- application-specific validation policy;
- user-facing migration UI and reporting;
- decisions about which workbook constructs are supported or rejected.

A consumer should not need to construct raw `VARIANT`s, know DISPIDs, call `IDispatch::Invoke` directly, or implement Excel COM lifetime rules.

## Current baseline

The existing crate already supports the bounded path:

```text
Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range
```

and already provides typed workbook open/save/close behavior, Automation values, rectangular arrays, apartment-bound wrappers, and explicit private application/session ownership.

Migration support should extend this design rather than create a parallel automation layer.

## Hard rules

1. **Private Excel means private Excel.** Crate-owned sessions must activate a new local `Excel.Application` instance. They must not attach to an arbitrary user's running Excel through the Running Object Table or `GetActiveObject`.
2. **No hidden raw escape hatch for normal consumers.** Migration-critical operations must be expressible through the ordinary public wrapper API.
3. **Desktop Excel is the runtime authority.** Dispatch metadata and mocked behavior are not sufficient evidence for migration-critical features.
4. **Excel remains the serializer.** The crate exposes reliable workbook open/save/SaveAs operations; it does not replace Excel with direct OOXML serialization.
5. **Version-sensitive behavior is centralized.** When Excel versions differ, compatibility handling belongs in `excel-com`, not duplicated in every consumer.
6. **Apartment ownership remains explicit.** Wrappers remain apartment-bound and neither `Send` nor `Sync` unless a future design with real COM marshaling explicitly changes that rule.
7. **Shutdown is explicit.** Dropping wrappers releases COM references but must not silently reinterpret application ownership. Crate-owned Excel shutdown remains an explicit operation with testable cleanup.

## Migration-critical capability inventory

The first implementation task is to compare this list against the current crate and classify each item as already supported, partially supported, or missing.

### Application and session

Required generic capabilities:

- create a new local/private `Excel.Application` instance;
- `Visible` read/write;
- `UserControl` read/write where supported/needed;
- `DisplayAlerts` read/write and scoped restoration guard;
- `AskToUpdateLinks` read/write;
- `AutomationSecurity` with a typed option corresponding to force-disable macros;
- calculation mode/state operations if a consumer needs deterministic verification;
- explicit `Quit` and deterministic release behavior.

A migration application opening untrusted/legacy `.xlsm` input must be able to force-disable macros before opening the workbook.

### Workbooks and workbook lifecycle

Required generic capabilities:

- `Workbooks.Open` with complete positional optional-argument fidelity;
- typed options for read-only operation and link-update behavior;
- workbook `Path` / `FullName` where available;
- worksheet collection access;
- workbook-level names access;
- `Workbook.SaveAs` with complete positional optional-argument fidelity;
- typed file-format enum for common formats, including at least `.xlsx` and `.xlsm`;
- explicit close with save/discard/prompt semantics.

Suggested typed file-format names should avoid magic integers in consumers, e.g. conceptually:

```text
OpenXmlWorkbook
OpenXmlMacroEnabledWorkbook
```

Exact Rust naming should follow existing crate conventions.

### Worksheet and range navigation

Required generic capabilities:

- worksheet name;
- worksheet lookup;
- `Range` access by address;
- `Cells` access/indexing;
- `UsedRange`;
- `SpecialCells` with typed cell-type selection, especially formula cells;
- range rows/columns/cells counts and indexing;
- `Range.Address` with useful absolute/relative options;
- range's parent worksheet.

`SpecialCells` must preserve the normal Excel behavior where no matching cells may be returned as a COM error; the wrapper should expose that state in a way callers can handle deliberately.

### Values and formulas

Required generic capabilities:

- `Value2` read/write;
- `HasFormula`;
- `Formula` read/write;
- `Formula2` read/write;
- a migration-friendly dynamic-formula abstraction if needed to centralize compatibility behavior.

The important semantic distinction is that `Formula2` uses Excel's dynamic-array formula model while legacy `Formula` can apply implicit-intersection semantics and manufacture a leading `@`.

Consumers should not each reproduce version/fallback logic. A suitable public abstraction may be something conceptually like:

```text
Range::dynamic_formula()
Range::set_dynamic_formula(...)
```

backed by `Formula2` when supported and a deliberate compatibility path when necessary. The exact API should be decided from live Excel evidence, not guessed.

### Legacy CSE array formulas

Add first-class support for legacy array-formula identity and whole-array operations:

- `HasArray`;
- `CurrentArray`;
- `FormulaArray`;
- whole-array formula replacement where Excel permits it.

The wrapper should make it practical for a consumer to discover that one cell belongs to a CSE array and then treat the entire `CurrentArray` as the migration unit. Consumers should not mutate individual cells inside an existing CSE array accidentally.

Live tests must cover this behavior on real desktop Excel.

### Names

Required generic capabilities:

- workbook `Names` collection;
- worksheet-local `Names` collection;
- lookup by exact name;
- `Name.RefersTo` where useful;
- `Name.RefersToRange`.

Both workbook-scoped and worksheet-scoped names need live coverage because migration consumers may resolve formula arguments against either scope.

### Automation values and optional arguments

Continue using typed wrapper boundaries for:

- numeric, Boolean, string, empty/null/error Automation values;
- rectangular SAFEARRAY values;
- optional/missing arguments represented as Excel expects (`VT_ERROR` / `DISP_E_PARAMNOTFOUND` at the private dispatch boundary);
- Windows path values without lossy UTF-8 conversion;
- typed Excel enums instead of public magic integers where practical.

The consumer should never have to know that positional COM arguments are reversed at the private dispatch layer.

## Error model

Migration-critical APIs should retain enough context to distinguish:

```text
member not supported by this Excel version
ordinary Excel object-model error
no matching SpecialCells result
invalid caller input
COM activation failure
Excel exception/HRESULT
```

Do not collapse every Excel automation failure into a string-only error.

Continue to avoid exposing raw pointer addresses or COM internals that make errors unstable or unsafe to log.

## Real Excel verification requirements

Every migration-critical capability added to `excel-com` requires a real desktop Excel live test in addition to unit/dispatch tests.

At minimum the live suite should prove:

### Private-instance ownership

- create the crate-owned Excel process;
- do not attach to an already running user Excel instance;
- closing/quitting the owned instance does not close a pre-existing user instance.

### Workbook lifecycle

- open a controlled `.xlsm` read-only;
- suppress prompts deliberately;
- SaveAs a separate `.xlsx`;
- close and reopen the `.xlsx`;
- verify expected workbook content through Excel.

### Formula2

- write a dynamic-array formula through the intended API;
- read it back without unwanted implicit-intersection `@` transformation;
- verify the actual Excel spill/formula behavior where available.

### CSE arrays

- create or open a legacy array formula fixture;
- identify `HasArray`;
- obtain `CurrentArray`;
- read `FormulaArray`;
- prove whole-array handling works and individual-cell misuse is rejected or otherwise safely surfaced.

### Names

- resolve workbook-scoped names;
- resolve worksheet-local names;
- obtain `RefersToRange` for each.

### UsedRange / SpecialCells

- enumerate formula cells from a controlled worksheet;
- handle the no-formula-cell case predictably.

### Macro security

- set the automation security mode before opening a macro-enabled workbook;
- prove the property is accepted by real Excel;
- where a safe controlled fixture can demonstrate it, prove workbook-open VBA is not executed.

## Consumer contract

A migration consumer is considered `excel-com`-clean only if its Excel automation layer contains no direct raw Excel COM dispatch implementation.

A practical source gate for a compiled migrator should reject migration-side use of constructs such as:

```text
IDispatch::Invoke
raw DISPID constants/member lookup
manual Excel VARIANT construction
Excel CLSID/ProgID activation
GetActiveObject / ROT attachment
manual SAFEARRAY ownership
```

Those belong in `excel-com`.

This rule does **not** forbid an application from using ordinary OS APIs for its own GUI, files, process launch, progress, or reporting.

## Recommended implementation sequence

### E0 - audit

Inventory every COM operation used by the existing migration reference implementation and map it to `excel-com`.

Output a table:

```text
operation | current excel-com API | status | live evidence | required action
```

No migration consumer rewrite should begin until this is complete.

### E1 - session/security gaps

Implement and live-test any missing:

- private instance semantics;
- `UserControl`;
- `AskToUpdateLinks`;
- automation security;
- calculation state/mode if required.

### E2 - worksheet/range discovery gaps

Implement and live-test:

- `UsedRange`;
- typed `SpecialCells`;
- address/parent/indexing gaps.

### E3 - formula and CSE gaps

Implement and live-test:

- `Formula2`;
- migration-friendly dynamic-formula behavior if justified by evidence;
- `HasArray`;
- `CurrentArray`;
- `FormulaArray`;
- safe whole-array update behavior.

### E4 - names and remaining workbook gaps

Implement and live-test:

- workbook names;
- worksheet-local names;
- `RefersTo` / `RefersToRange`;
- any typed SaveAs/file-format improvements found by the audit.

### E5 - migration-complete gate

Run the complete live suite serially against real desktop Excel and record the tested Excel version/bitness.

The gate passes only when every COM operation required by the migration consumer is available through public `excel-com` API with real Excel evidence.

## Cross-repository handoff

When E5 passes, the consuming repository may begin/continue the compiled migrator implementation.

The handoff should record an immutable `excel-api` commit or crate version so the migration executable does not silently depend on a moving COM surface.

The consumer should then pin that reviewed revision/version until an intentional update is made.

## Scope discipline

This migration-support effort should not be used as justification to wrap the entire Excel type library before progress can continue.

The target is:

> complete support for every generic Excel COM operation required by the migration workflow, with architecture that makes subsequent additions straightforward.

Unrelated Excel object-model expansion can proceed independently.

## Coordination

This document is additive and intentionally avoids changing the current `excel-com` implementation while other work may be in flight elsewhere.

Before implementation:

1. fetch fresh `origin/master`;
2. review current `excel-com` public API and live-test evidence rather than assuming this note is still current;
3. preserve valid post-note improvements;
4. make small ownership-local commits;
5. keep live Excel evidence alongside the API capability it validates.
