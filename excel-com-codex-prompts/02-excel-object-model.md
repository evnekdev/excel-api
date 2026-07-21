# Codex Prompt 02 — Inventory the Excel COM Object Model

## Objective

Produce an implementation-oriented inventory of the Microsoft Excel Automation object model, focusing on the object hierarchy and navigation paths required for an xlwings-style Rust API.

This prompt is documentation and research only. Do not create production wrappers.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
docs/excel-com-02-object-model
```

Create the branch from the latest reviewed `master`.

## Required reading

Read all current Excel COM research, especially:

```text
docs/research/excel-com/01-com-automation-foundations.md
```

Also read `ARCHITECTURE.md`, `COM_ARCHITECTURE.md`, `OPTIONAL_INTEGRATIONS_ROADMAP.md`, and the existing RTD type-library research.

## Sources

Use the official Microsoft Excel VBA object-model reference as the primary source. Use Microsoft Office Primary Interop Assembly documentation as a secondary typed projection.

Clearly distinguish:

- documented Excel object-model behavior;
- .NET-specific projection details;
- information inferred from type-library metadata;
- behavior requiring experimentation.

## Required deliverable

Create:

```text
docs/research/excel-com/02-excel-object-hierarchy.md
```

## Required analysis

### Root model

Document `Application` as the Automation root and map its main navigation paths:

```text
Application
├── Workbooks
│   └── Workbook
│       ├── Worksheets
│       ├── Sheets
│       ├── Names
│       ├── Windows
│       ├── Connections
│       └── other major child collections
├── Windows
├── AddIns
├── Names
├── Calculation state/settings
└── Events
```

### Initial typed-wrapper scope

Analyze these objects in detail:

- `Application`;
- `Workbooks`;
- `Workbook`;
- `Sheets`;
- `Worksheets`;
- `Worksheet`;
- `Range`;
- `Names` and `Name`;
- `ListObjects` and `ListObject`;
- `ChartObjects`, `ChartObject`, and `Chart`;
- `Shapes` and `Shape`;
- `Windows` and `Window`.

For each object or collection record:

```text
Object name
Parent object
Default interface
Default member
Primary collection membership
Indexing rules
Name-based lookup
Important properties
Important methods
Returned object types
Optional parameters
Read/write behavior
Relevant events
Likely initial support level
Raw-dispatch escape requirements
```

### Collections

Determine and document:

- whether indexing is 1-based;
- whether `Item` accepts numeric indices, names, or both;
- how default members are exposed;
- whether `_NewEnum` is available;
- whether iteration should use COM enumeration or indexed access;
- differences between `Sheets` and `Worksheets`;
- heterogeneous collections;
- ordering guarantees;
- mutating collection methods.

Do not assume all collections behave identically.

### Optional arguments

Identify methods whose signatures materially affect the invocation layer, especially:

- `Workbooks.Open`;
- `Workbook.SaveAs`;
- `Worksheets.Add`;
- `Range.Find`;
- `Range.Sort`;
- `Application.Run`;
- export and print methods.

Record positional parameter order, named parameters, optionality, documented defaults, and methods with unusually large parameter lists.

### API layering proposal

Classify object-model coverage into:

#### Release 0.1

- `Application`;
- `Workbooks`;
- `Workbook`;
- `Worksheets`;
- `Worksheet`;
- `Range`.

#### Early follow-up

- names;
- tables;
- calculations;
- macro execution;
- basic formatting.

#### Later

- charts;
- shapes;
- pivots;
- connections;
- events;
- specialized analysis objects.

Explain the classification based on user value and implementation dependencies.

### Public naming questions

Discuss Rust naming and ergonomics without freezing the final API:

```rust
ExcelApp
Workbooks
Workbook
Worksheets
Worksheet
Range
```

Analyze whether collection wrappers are worthwhile or whether convenience methods should live directly on parents. Compare representative usage with VBA, pywin32, xlwings, and C# Interop without copying those APIs blindly.

## Capability matrix

Add a table:

```text
Excel capability | Required COM objects | Required value support | Required reliability support | Proposed milestone
```

Include create Excel, attach to Excel, open/add workbook, access sheet by name, scalar and rectangular values, formulas, names, tables, macro execution, charts, PDF export, and events.

## Validation and completion

Run the normal workspace formatting, Clippy, test, and documentation checks.

Commit, push, and open a draft PR. The PR must identify the proposed 0.1 boundary, collection/indexing hazards, optional-argument implications, and open questions requiring type-library inspection or experiments.

Do not merge. Stop after reporting the draft PR.
