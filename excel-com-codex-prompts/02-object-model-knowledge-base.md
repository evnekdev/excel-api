# Codex Prompt 02 — Build the Excel Object-Model Knowledge Base Foundation

## Objective

Build a reproducible, attributed, machine-readable foundation from the official
Microsoft Excel VBA documentation source. Normalize objects, collections,
properties, methods, events, enumerations, relationships, source provenance,
and independent documentation/type-library/runtime verification dimensions.

This is research infrastructure only. Do not create `excel-com`, activate
Excel, invoke `IDispatch`, inspect the installed type library, run real-Excel
experiments, freeze a Rust public API, mirror Microsoft Learn, or copy long
documentation passages/examples.

## Branch and preparation

Create `feature/excel-com-02-object-model-kb` from a clean, rebased
`origin/master` and record the starting SHA. Read the core/XLL architecture,
optional-integration boundary, existing RTD research, Prompt 01, and the
repository's tooling/CI/generated-file conventions. Keep the new tool outside
all core and example dependency graphs.

## Required deliverables

Create an excluded or non-default unpublished tool under `tools/excel-com-kb/`
and an attributed knowledge base under `knowledge/excel-object-model/`:

```text
SOURCE_MANIFEST.toml  ATTRIBUTION.md  README.md
schema/{object,member,relationship,enum,source}.schema.json
data/{objects,members,relationships,enums}.jsonl  data/source.json
generated/{object-index,member-kind-summary,source-coverage,unresolved-classification}.md
```

Create `docs/research/excel-com/02-object-model-knowledge-base.md` explaining
the source, licence evidence, selection, stable IDs, schema/provenance model,
ingestion, limitations, regeneration, upstream updates, and later use. Link to
Prompt 01 instead of duplicating generic COM analysis. Update the architecture
and this prompt index only enough to expose the new foundation.

## Source and provenance policy

Use a separately supplied checkout of `MicrosoftDocs/VBA-Docs`, pinned by an
exact Git commit in the manifest. Inspect and record its documentation and
sample-code licence files, repository URL, retrieval date, and verified include
and exclude patterns. Add a licence-appropriate attribution notice and identify
transformed/project-authored content.

Do not scrape rendered Learn HTML, use a mutable branch tip in ordinary
generation, vendor a clone/archive, add a submodule, or require network access
for normal validation. Do not copy VBA examples. Every canonical record must
retain a relative source path, source commit, extraction/field provenance,
documentation/typelib/runtime evidence, and implementation status.

## Data and tool requirements

Use deterministic UTF-8 JSON Lines with one sorted complete record per line,
stable IDs, LF/final newline, no absolute paths or clock data outside the
manifest, and schemas that reject bad IDs, unknown kinds, missing provenance,
invalid owner/relationship endpoints, invalid statuses, and impossible initial
verification values. Keep unresolvable classification visible in a report;
never guess a DISPID, default member, collection item, return object, or runtime
behavior.

Provide offline commands equivalent to `ingest`, `validate`, `generate`, and
`check`. `ingest` accepts a source checkout, manifest, and output directory;
`check` verifies canonical data, generated files, ordering, paths, endings, and
determinism without accessing the network.

Use compact project-authored fixtures and tests covering object, property,
method/optional parameter, event, collection, enum, malformed/missing metadata,
duplicate IDs/API names, unresolved references, capitalization, HTML, Unicode,
provenance, paths, and byte-identical regeneration.

## Completion

Run workspace and tool validation, schema validation, local Markdown-link
validation, determinism checks, and scans for absolute user paths, copied long
passages, and upstream cache/archive files. If the full source is unavailable,
record that blocker and do not claim full coverage. Otherwise commit the
normalized outputs and exact generated counts.

Open a draft PR against `master` describing source/licence evidence, data model,
coverage counts, deterministic evidence, limitations, validation, and the
non-effects on COM/XLL production behavior. Do not merge or begin the former
narrative inventory prompt.
