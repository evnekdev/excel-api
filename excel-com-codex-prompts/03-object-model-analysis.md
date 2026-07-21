# Codex Prompt 03 - Analyze the Excel Automation Object Model from the Knowledge Base

## Objective

Turn the Prompt 02 knowledge base into a reviewable, documentation-derived
object-model analysis. This is research only: do not introduce a production
`excel-com` crate, COM activation, a type-library inspection, real Excel, or a
public API.

## Prerequisite and branch

Verify the Prompt 02 knowledge-base paths from the current `origin/master` tree
and run the offline `excel-com-kb check`. Branch from that current tree as:

```text
research/excel-com-03-object-model-analysis
```

Do not require the former Prompt 02 feature commit to be an ancestor; a squash
merge with equivalent content is sufficient.

## Required preparation

Read the repository architecture, Prompt 01, Prompt 02, the knowledge-base
README, manifest, generated reports, unresolved review queue, and the current
prompt sequence. Treat canonical records as the sole source for object-model
claims. Every conclusion must be explicitly labelled as
Documentation-established, Project classification, Type-library verification
required, Runtime verification required, Architecture candidate, or Deferred.

## Deliverables

Create `docs/research/excel-com/03-excel-object-model-analysis.md` and generate
these deterministic reports under `knowledge/excel-object-model/generated/analysis/`:

- `architectural-spine.md`
- `candidate-0.1-members.md`
- `collection-analysis.md`
- `optional-argument-targets.md`
- `typelib-audit-targets.md`
- `runtime-probe-targets.md`

Extend `tools/excel-com-kb` only when needed. It must remain standalone,
offline, deterministic, and unable to modify canonical records during analysis.

## Required analysis

Analyze `Application -> Workbooks -> Workbook -> Worksheets -> Worksheet ->
Range`; adjacent sheets/windows/names/tables/charts/shapes; mandatory
collections; Range capabilities; the high-risk optional-argument set; a narrow
candidate 0.1; type-library targets; runtime probes; and knowledge-base quality.
Use only canonical object-return edges. Do not infer default members,
enumeration, aliases, DISPIDs, invocation flags, or runtime behavior.

## Required exclusions

Do not add `excel-com`, `IDispatch::Invoke`, activation, Excel experiments,
installed-type-library facts, ABI claims, dependency changes, or a frozen public
API. Do not manually edit canonical knowledge-base data.

## Validation and handoff

Run deterministic regeneration, the standalone tool's format/clippy/tests/check,
repository format/clippy/tests/doc tests, Markdown-link validation, schema and
absolute-path checks, and `git diff --check`. Open a draft PR reporting source
counts, candidate and target counts, validations, blockers, and the exclusions
above.
