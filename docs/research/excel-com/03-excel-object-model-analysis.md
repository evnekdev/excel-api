# Excel Automation object-model analysis

**Status:** documentation-derived analysis and an evidence backlog. No
`excel-com` crate, activation path, real Excel experiment, or type-library
inspection is introduced.

**Date:** 2026-07-21
**Baseline:** `origin/master` `347f82381aad8212b98ab5c400f20033daf3b891`

## 1. Scope and evidence boundary

This record consumes the pinned knowledge base from Prompt 02. It distinguishes
documentation extraction from design and experiment work:

- **Documentation-established:** an assertion is represented by a canonical
  record, stored relationship, or pinned source-coverage value.
- **Project classification:** a candidate grouping or priority for later work.
- **Type-library verification required:** interface identity, `DISPID`, invoke
  kind, optional/default encoding, aliases, default members, and event metadata
  are not documentation facts in this repository.
- **Runtime verification required:** application state, indexing, enumeration,
  omission behavior, range shape, coercion, and lifetime need real Excel.
- **Deferred:** no commitment is made in this prompt.

The analysis tool reads the canonical records and writes only deterministic
Markdown under `knowledge/excel-object-model/generated/analysis/`. It does not
modify `data/`, inspect an installed typelib, start Excel, activate COM, or
create a production API.

## 2. Knowledge-base baseline and check

The prerequisite was verified from the current `origin/master` tree, rather
than by requiring the obsolete Prompt 02 branch commit to be an ancestor. The
following paths exist at that baseline:

- `knowledge/excel-object-model/`
- `tools/excel-com-kb/`
- `docs/research/excel-com/02-object-model-knowledge-base.md`
- `excel-com-codex-prompts/02-object-model-knowledge-base.md`

Documentation-established: the offline check passes. The pinned corpus reports
5,928 selected source files, 5,156 parsed files, 268 objects, 4,632 members,
and 5,770 relationships. Its source remains
`MicrosoftDocs/VBA-Docs` at `b2cda886ea91e36c62eb1cb177133ad024ecd345`.

The new `excel-com-kb analyze` command derives six reports and `check` verifies
them when the analysis directory exists. This preserves the original Prompt 02
check behavior for a knowledge base that predates analysis output.

## 3. Automation root analysis

Documentation-established: `Excel.Application` is the root object and its
records describe `Visible`, `Version`, `Workbooks`, `ActiveWorkbook`,
`ActiveSheet`, `Calculation`, `DisplayAlerts`, `ScreenUpdating`, `EnableEvents`,
`Run`, and `Quit`. `ActiveWorkbook` and `ActiveSheet` documentation explicitly
allows no-current-object cases; no Rust null/error policy follows from that.

Project classification: these root members are a candidate dynamic entry
surface, not a public wrapper contract. Type-library verification required:
activation class/interface identity and every invocation member's dispatch
metadata. Runtime verification required: all availability and lifecycle
behavior.

## 4. Architectural spine

Project classification identifies the narrow navigation spine:

`Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range`.

Documentation-established canonical return relationships support the root,
workbook creation/opening, workbook-to-worksheets, worksheet-to-range, and
range navigation edges. The complete evidence table is in
[architectural spine](../../../knowledge/excel-object-model/generated/analysis/architectural-spine.md).
It lists only stored return edges; a missing edge is an audit target, not an
invented relation.

`Sheets`, `Windows`/`Window`, `Names`/`Name`, `ListObjects`/`ListObject`,
`ChartObjects`/`ChartObject`/`Chart`, and `Shapes`/`Shape` are
Documentation-established adjacent families. Project classification: all are
deferred outside candidate 0.1.

## 5. Collection analysis

The mandatory review set is `Workbooks`, `Worksheets`, `Sheets`, `Windows`,
`Names`, `ListObjects`, `ChartObjects`, and `Shapes`. Documentation-established
records retain the available `Count`, `Item`, and `Add` members; their object
summaries use collection language. The detailed matrix is in
[collection analysis](../../../knowledge/excel-object-model/generated/analysis/collection-analysis.md).

Project classification: `Sheets` is treated as potentially heterogeneous and
must not receive a typed item wrapper by inference. Type-library verification
required: default members, item return interfaces, `_NewEnum`, `DISPID`s, and
collection aliases. Runtime verification required: integer/name lookup,
index-base, enumeration, and failure behavior. No generic Rust collection API
is selected here.

## 6. Range capability decomposition

Documentation-established Range records group naturally into addressing and
selection (`Worksheet.Range`, `Cells`, `UsedRange`, `Range.Item`), navigation
(`Rows`, `Columns`, `Offset`, `Resize`), values/formulas (`Value`, `Value2`,
`Formula`, `Formula2`), mutation (`ClearContents`, `Clear`, `Delete`, `Copy`,
`PasteSpecial`), and inspection/discovery (`Address`, `Count`, `Find`, `Sort`).

Project classification: Range is the value-and-formula transport boundary, but
not yet a Rust wrapper. Deferred: scalar/array `VARIANT` and `SAFEARRAY`
representation, dimensions, bounds, formulas, errors, and write-back semantics
belong to Prompt 04. Runtime verification required for every operation's shape
and side effect.

## 7. Optional-argument risk analysis

The high-risk target list is `Workbooks.Open`, `Workbook.SaveAs`,
`Workbook.Close`, `Worksheets.Add`, `Range.Find`, `Range.Sort`,
`Application.Run`, `Workbook.ExportAsFixedFormat`, and
`Range.ExportAsFixedFormat`.

Documentation-established: the canonical records retain parameter order,
optionality, selected types, and enum relationships. The generated
[optional-argument targets](../../../knowledge/excel-object-model/generated/analysis/optional-argument-targets.md)
report gives the exact retained sequence. Deferred: complete source syntax and
default values are not stored, and are not reconstructed. Type-library
verification required for COM parameter direction/default encoding and invoke
kind; Runtime verification required for omission behavior and persistent state.

## 8. Candidate 0.1 surface

Project classification: the candidate inventory contains 50 currently present
canonical members across `Application`, `Workbooks`, `Workbook`, `Worksheets`,
`Worksheet`, and `Range`. It is a breadth-control device, not a frozen public
API, milestone, or promise of typed wrappers. The per-object member evidence is
in [candidate 0.1 members](../../../knowledge/excel-object-model/generated/analysis/candidate-0.1-members.md).

`Application` covers visibility, active objects, controls, macro invocation,
and quit; the middle objects cover identity, navigation, and basic persistence;
Range covers addressing, values/formulas, navigation, mutation, find, and sort.
All candidate members remain gated by type-library and runtime verification.

## 9. Deferred object groups

Deferred: charts, names, shapes, tables, pivots, queries, connections,
add-ins, window/UI families, and events. Their documentation records are useful
for later audit planning but do not authorize wrapper generation, event sinks,
or public types.

## 10. Type-library verification backlog

At the time of this documentation-derived analysis, no installed typelib had
been inspected. The prioritized
[typelib audit target](../../../knowledge/excel-object-model/generated/analysis/typelib-audit-targets.md)
sets cover the Application root, the 50 candidate members, collection mechanics,
Range transport and optional calls, event interfaces, then adjacent deferred
families. Required facts include coclass/CLSID, IID, `DISPID`, invoke kind,
parameter direction/type/default, return interface, default member, `_NewEnum`,
aliases, enum values, and connection points.

## 11. Runtime verification backlog

No Excel process was started. The
[runtime probe targets](../../../knowledge/excel-object-model/generated/analysis/runtime-probe-targets.md)
prioritize 18 probes covering availability, open/save/add, collection index and
enumeration, range selection/value/formula/mutation/navigation, `Find`, `Sort`,
application control restoration, macro invocation, export, and shutdown.
Runtime verification required applies to every target.

## 12. Knowledge-base quality assessment

Documentation-established: coverage records 918 unresolved return-object
references and 3,083 review entries. The parser deliberately does not guess
return objects from prose, dispatch facts, default members, collection
enumeration, or runtime behavior. This is an evidence boundary, not a defect to
paper over.

Project classification: `.gitattributes` now forces LF checkout endings for
canonical JSONL and generated reports. It fixes a Windows `core.autocrlf`
checkout false failure without changing canonical record content. The check now
verifies byte-deterministic LF/final-newline output locally as intended.

## 13. Architecture candidates, not decisions

Project classification: a later architecture may begin with a small dynamic
kernel, preserve a raw value boundary for Range transport, and add typed
wrappers only after type-library evidence. This analysis does **not** choose
dynamic versus typed wrappers, an ownership model, a threading model, an error
model, a collection abstraction, a lifecycle policy, or a public module shape.

## 14. Non-decisions and exclusions

No `excel-com` crate was created. No `IDispatch::Invoke`, activation,
apartment, COM ownership, installed typelib inspection, Excel process, event
sink, public API, dependency-graph change, ABI claim, or production test was
introduced. No canonical knowledge-base data was manually changed.

## 15. Next prompt

The core installed-typelib audit now precedes the later Range runtime work; see
[Prompt 04](04-core-excel-typelib-audit.md). The existing Range runtime prompt
must consume both the Range transport/runtime targets and the declared
type-library evidence without treating this candidate inventory as a public API.

## 16. Validations

The following passed after regeneration:

```text
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo fmt --manifest-path tools/excel-com-kb/Cargo.toml --check
cargo clippy --offline --manifest-path tools/excel-com-kb/Cargo.toml --all-targets -- -D warnings
cargo test --offline --manifest-path tools/excel-com-kb/Cargo.toml
cargo run --offline --manifest-path tools/excel-com-kb/Cargo.toml -- analyze --root knowledge/excel-object-model
cargo run --offline --manifest-path tools/excel-com-kb/Cargo.toml -- check --root knowledge/excel-object-model
```

The standalone test covers byte-deterministic analysis generation. The final
change review also checks Markdown links, schema validation through `check`,
absolute-path absence in canonical data, LF endings, deterministic regeneration,
and that no canonical record files changed.
