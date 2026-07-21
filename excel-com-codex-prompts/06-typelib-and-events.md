# Codex Prompt 06 — Audit Excel Events and the Installed Type Library

## Objective

Inspect the installed Excel COM type library and document metadata required for dynamic calls, typed wrappers, constants, default members, and future event support.

This remains research work. Do not generate a complete production binding crate.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
research/excel-com-06-typelib-events
```

## Required preparation

Read all prior Excel COM research and the existing RTD type-library audit. Apply the same evidence discipline used for RTD ABI work.

## Deliverables

Create:

```text
docs/research/excel-com/06-excel-typelib-and-events.md
tools/excel-com-typelib-audit/
```

The audit tool must be clearly marked as non-production developer tooling.

## Type-library inspection

Use official COM type-information APIs such as:

- `LoadTypeLibEx`;
- `ITypeLib`;
- `ITypeInfo`;
- type attributes;
- function and variable descriptions;
- documentation strings;
- referenced type descriptions.

Record:

- Excel type-library GUID, version, and LCID;
- major coclasses;
- default interfaces and source interfaces;
- interface IIDs and coclass CLSIDs;
- DISPIDs;
- member names and invocation kinds;
- argument names and order;
- optional/default flags;
- return types;
- default members;
- `_NewEnum`;
- hidden and restricted members;
- enum values and aliases.

Do not hand-copy an enormous undocumented binding surface. Produce reproducible audit output.

## Scope the audit

Prioritize:

- `Application`;
- `Workbooks` and `Workbook`;
- `Worksheets` and `Worksheet`;
- `Range`;
- `Names` and `Name`;
- `ListObjects` and `ListObject`;
- application, workbook, and worksheet events.

## Event model

Document:

- `IConnectionPointContainer`;
- `IConnectionPoint`;
- `FindConnectionPoint`;
- `Advise` and `Unadvise`;
- subscription cookies;
- outgoing dispatch interfaces;
- event DISPIDs and arguments;
- by-reference cancellation parameters;
- apartment of event delivery;
- reentrancy;
- panic containment;
- cleanup on Excel shutdown;
- subscription lifetime.

Recommend whether events belong in the initial crate, an opt-in feature, a separate crate, or a later milestone.

## Generated versus dynamic API

Compare:

### Fully dynamic

Resolve every member by name, cache DISPIDs, generate little code, maximize flexibility, and accept runtime errors.

### Generated metadata

Generate DISPIDs, names, constants, argument metadata, and enums while retaining dynamic invocation.

### Generated typed interfaces

Generate strongly typed wrappers, with larger surface and versioning/compatibility concerns.

Recommend a staged strategy.

## Reproducibility

The tool must:

- avoid machine-specific paths;
- accept an explicit type-library path or discover the registered library safely;
- report the selected version;
- produce deterministic reviewable output where possible;
- fail clearly if Excel is absent;
- never modify registration.

Commit a small relevant audit fixture rather than an enormous unreviewed dump.

## Validation and completion

Run all applicable checks. Open a draft PR summarizing type-library identity, dynamic-dispatch implications, event findings, generation strategy, and documentation/metadata discrepancies.

Do not merge. Stop after reporting the draft PR.
