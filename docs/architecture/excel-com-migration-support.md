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

### Audited PyroApp reference surface (2026-09-16)

This matrix is the completed E0/M0 audit of
`pyroapprs/tools/workbook-migrator/Invoke-PyroAppWorkbookMigration.ps1` and
`tools/workbook-migrator/tests/WorkbookMigration.Tests.ps1` against
`excel-com` at `c0dd9a0db1286226875030c2003e7e4db37a5df4`. The PowerShell test fixture
uses a few construction calls that the production migration path does not use;
they remain listed because the compiled replacement needs equivalent synthetic
acceptance fixtures.

Status meanings are deliberately literal:

- **SUPPORTED BY excel-com**: a public typed API exists;
- **PARTIALLY SUPPORTED**: the underlying operation exists, but the public
  migration-facing surface or focused live evidence is incomplete;
- **MISSING FROM excel-com**: generic Excel behavior must be added here before
  the compiled migrator may depend on it;
- **NOT COM — migration-specific**: belongs in the PyroApp application;
- **NOT COM — UI/application concern**: belongs in the executable shell or OS
  integration.

| Reference operation / behavior | Public `excel-com` API | Status at audit | Existing real-Excel evidence | Required action |
|---|---|---|---|---|
| Create a new private `Excel.Application` | `ComApartment::sta`, `OwnedApplication::new` | SUPPORTED BY excel-com | `fixture_open_live`, `release_smoke_live`, session diagnostics | Keep owned activation as the only migrator construction path; never use `AttachedApplication`. |
| `Application.Visible` | `visible`, `set_visible` | SUPPORTED BY excel-com | release/session live suites | None. |
| `Application.UserControl` | none | MISSING FROM excel-com | Earlier runtime probes only; no public-wrapper live test | Add typed read/write access and live-test private-session behavior. |
| `Application.DisplayAlerts` | `display_alerts`, `set_display_alerts`, `display_alerts_guard` | SUPPORTED BY excel-com | release/presentation live suites | Use the restoring guard in the migrator. |
| `Application.AskToUpdateLinks` | `ask_to_update_links`, `ask_to_update_links_guard` | SUPPORTED BY excel-com | external-link suite is environment-blocked; property participates in safe-open code | Add focused migration-support live evidence on the available desktop Excel host. |
| `Application.EnableEvents` | none | MISSING FROM excel-com | Installed type-library evidence only | Add typed read/write access plus a restoring guard and live test. |
| `Application.AutomationSecurity = ForceDisable` | `AutomationSecurity`, `automation_security_guard`, `Workbooks::open_safely` | SUPPORTED BY excel-com | `macro_runtime_live` verifies guard restoration | Exercise `open_safely` against the controlled `.xlsm` fixture in the migration suite; do not enable macro execution. |
| `Application.Calculation = Manual` | `calculation_mode`, `calculation_mode_guard`, `CalculationMode::MANUAL` | SUPPORTED BY excel-com | `formula_calculation_live` | Use the guard rather than an un-restored scalar assignment. |
| `CalculationState`, `Calculate`, `CalculateFull`, `Quit` | typed state/method APIs and `OwnedApplication::quit_and_wait` | SUPPORTED BY excel-com | calculation and release smoke suites | Include exact-process exit evidence in the migration lifecycle test. |
| `Application.Workbooks` and `Workbooks.Add` (fixture only) | `Application::workbooks`, `Workbooks::add` | SUPPORTED BY excel-com | multiple live suites | Production migration must open a copied fixture; `Add` remains test-only. |
| `Workbooks.Open(path, UpdateLinks=0, ReadOnly=true)` | `Workbooks::open` with `WorkbookOpenOptions`; `open_safely` | SUPPORTED BY excel-com | fixture/file live suites | Use typed options and safe open; preserve all optional positions inside the crate. |
| `Workbook.Worksheets` / indexed and named lookup | `Workbook::worksheets`, `Worksheets::item_by_index`, `item_by_name`, `iter` | SUPPORTED BY excel-com | worksheet/collection live suites | None. |
| `Workbook.Names`, `Names.Item`, `Names.Add` (fixture) | `Workbook::names`, `Names::item_by_name`, `Names::add` | SUPPORTED BY excel-com | `reference_names_live` | Extend live evidence to same textual name in workbook and worksheet scopes. |
| `Workbook.SaveAs(..., 51/52)` | `Workbook::save_as`, `WorkbookSaveAsOptions`, `XlFileFormat::{OPEN_XML_WORKBOOK, OPEN_XML_WORKBOOK_MACRO_ENABLED}` | SUPPORTED BY excel-com | workbook-file and release smoke suites | Use typed formats; add `.xlsm`→`.xlsx` reopen verification. |
| `Workbook.Save`, `Path`, `FullName`, `Close` | `save`, `path`, `full_name`, consuming `close` | SUPPORTED BY excel-com | workbook-file live suite | The migrator must not call `save` on its source. Close explicitly with discard after SaveAs. |
| `Workbook.HasVBProject` for output verification | `has_vb_project` | SUPPORTED BY excel-com | presentation live coverage is not migration-focused | Verify the reopened `.xlsx` reports no VBA project. |
| `Worksheet.Name`, `Range`, `Cells`, `UsedRange`, `Names` | `name`, `range`, `cell`/`range_from_cells`, `used_range`, `names` | SUPPORTED BY excel-com | worksheet/range, structured-data, reference-name suites | None beyond consolidated migration evidence. |
| `Range.Value2`, `Formula`, `Formula2`, `HasFormula` | `value2`/`set_value2`, formula APIs, `has_formula` | SUPPORTED BY excel-com | worksheet/range and formula-calculation suites | Migrator uses `Formula2` directly; an older-version fallback is not enabled without contrary live evidence. |
| Formula2 compatibility fallback to Formula | no silent fallback by design | SUPPORTED BY excel-com | installed Excel 16.0 accepts Formula2 | Keep version-sensitive failure inside `excel-com`; do not let the migrator catch arbitrary Formula2 errors and retry Formula. |
| `Range.HasArray`, `CurrentArray`, `FormulaArray`, whole-array set | `has_array`, `current_array`, `formula_array`, `set_formula_array` | SUPPORTED BY excel-com | `formula_calculation_live` uses a genuine legacy CSE array and checks partial-edit safety | The PowerShell reference does not protect CSE members; the compiled migrator must group by `CurrentArray` and migrate each array once. |
| `Range.Address`, rows, columns, cells, areas and indexed cell access | typed address/count/navigation APIs and `Areas` | SUPPORTED BY excel-com | worksheet/range and formula-calculation suites | None. |
| `Range.SpecialCells(xlCellTypeFormulas)` | `special_cells` / `try_formula_cells`, `SpecialCellType::FORMULAS` | SUPPORTED BY excel-com | `formula_calculation_live` covers formula discovery and no-match handling | Prefer `try_formula_cells` so “no formula cells” is `Ok(None)`. |
| `Range.Worksheet` | no public method (only private internal dispatch use) | MISSING FROM excel-com | registry metadata exists; no public live test | Add `Range::worksheet` and verify identity/name on real Excel. |
| `Name.Name`, `RefersTo`, `RefersToRange` | `name`, `refers_to`, `range` | PARTIALLY SUPPORTED | `reference_names_live` covers global/local ranges | Add the explicit `refers_to_range` spelling as a compatibility-preserving alias and focused scope-collision evidence. |
| Explicit COM release / zombie-process prevention | apartment-bound wrappers, consuming close, `quit_and_wait` exact PID | SUPPORTED BY excel-com | `release_smoke_live` | Add a migration-like traversal test that drops every temporary wrapper before quit. |
| Formula recognition/translation and unsupported-family rejection | not applicable | NOT COM — migration-specific | PowerShell pure tests | Port to pure Rust tests without Excel. |
| `CA_CALCULATE` parsing and three-row header normalization | not applicable | NOT COM — migration-specific | PowerShell pure/live tests | Preserve behavior, then add CSE-aware planning. |
| File selection, save dialog, progress, cancellation, report and error dialogs | not applicable | NOT COM — UI/application concern | PowerShell UI/static tests | Implement in the standalone Windows executable. |
| Source-closed preflight, isolated staging copy, output collision policy and partial-output cleanup | not applicable | NOT COM — UI/application concern | PowerShell integration behavior | Implement in the executable with ordinary filesystem APIs. |

The audit also identified two behavioral gaps in the reference implementation,
not missing COM members: it rewrites formula cells individually without
`CurrentArray` grouping, and it treats a successful `SaveAs` as completion
without closing and reopening the output. The compiled implementation must
correct both, while preserving the reference formula and header semantics.

**E0/M0 gate result:** complete. Every current COM interaction and every
additional COM operation required by the target CSE/output-verification design
has an owner. E1-E5 must close the four public/evidence gaps above before the
compiled migrator starts.

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

### E5 outcome (2026-09-16)

The migration-support gate is complete at the implementation represented by
this workstream. The E0 gaps were closed as follows:

| Gap | Final public API / policy | Real-Excel evidence |
|---|---|---|
| `Application.UserControl` | `user_control`, `set_user_control` | A crate-owned Excel 16.0 session accepted and round-tripped `false`. |
| `Application.EnableEvents` | `enable_events`, `set_enable_events`, `enable_events_guard` | Excel 16.0 accepted the temporary value and explicit restoration. |
| `Range.Worksheet` | `Range::worksheet` | The returned worksheet matched the source worksheet by canonical COM identity. |
| explicit `RefersToRange` and name-scope safety | `Name::refers_to_range`; scope-verifying `Names::item_by_name` | A workbook-global and worksheet-local `MigrationHeader` resolved independently before and after SaveAs/reopen. The first live run caught Excel's ambiguous raw lookup and drove the scope fix. |
| `Workbooks.Open(UpdateLinks := 0)` | `XlUpdateLinks::DO_NOT_UPDATE` | The copied `.xlsm` opened read-only through `open_safely` with link prompts disabled. |
| consolidated lifecycle evidence | `tests/migration_support_live.rs` | Excel 16.0: safe `.xlsm` open, Formula2, genuine CSE `CurrentArray`, UsedRange/SpecialCells, scoped names, `.xlsx` SaveAs/reopen, no VBA project, and exact owned-process exit all passed. |
| private-instance isolation | second test in `migration_support_live.rs` | Two `OwnedApplication::new` calls produced different observed process IDs; quitting the newer instance left the pre-existing instance callable, and both exited naturally. |

The installed Automation surface reports Excel version `16.0`. Office bitness
is not authoritative through the available object model and remains `None`
rather than being guessed. This host can reject restoration of
`Application.Calculation` with Excel error `0x800A03EC`; the reference migrator
already treats manual calculation as a best-effort performance setting. No
migration correctness rule depends on changing calculation mode, while the
typed guard continues to expose the restoration result.

There are no remaining unsupported migration-critical COM operations. The
consumer must still group legacy CSE members by `CurrentArray`, reopen and
verify its output, and contain no raw COM implementation.

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
