# Excel object-model knowledge-base foundation

**Status:** documentation-derived research infrastructure; no `excel-com`
crate or COM implementation is introduced.
**Date:** 2026-07-21

## Purpose and boundary

The repository needs a local, reviewable set of facts for later Excel
Automation work, not a duplicate documentation site. The
[knowledge base](../../../knowledge/excel-object-model/README.md) keeps only
stable identifiers, concise summaries, structure, relationships, and source
provenance. It intentionally does not decide a public Rust API, COM apartment,
DISPID, collection wrapper policy, range-array orientation, or application
lifecycle.

Generic COM ownership, dispatch, Automation values, and apartment contracts
remain in [Prompt 01](01-com-automation-foundations.md); this document does
not repeat them.

## Input, licence, and selection

The authoritative input is the public
[`MicrosoftDocs/VBA-Docs`](https://github.com/MicrosoftDocs/VBA-Docs) checkout
pinned in
[`SOURCE_MANIFEST.toml`](../../../knowledge/excel-object-model/SOURCE_MANIFEST.toml):
commit `b2cda886ea91e36c62eb1cb177133ad024ecd345`, retrieved 2026-07-21. The
repository `LICENSE` is Creative Commons Attribution 4.0 International for
documentation; `LICENSE-CODE` is MIT for sample code. The distinction and the
transformation notice are recorded in
[`ATTRIBUTION.md`](../../../knowledge/excel-object-model/ATTRIBUTION.md).

The selected corpus is the verified direct `api/Excel.*.md` corpus plus the
three direct `api/overview/Excel/*.md` files. `*-graph-*` pages are excluded:
they are adjacent graph-reference pages that share an `Excel` filename prefix
but are not the Excel VBA object-model corpus. The source is supplied as a
separate checkout; no clone, archive, submodule, rendered Learn HTML, or
network access is required by normal checks.

## Data model and provenance

Versioned JSON Schemas define objects, members, relationships, enumerations,
and source coverage. JSON Lines provides one deterministic UTF-8 record per
entity, with lexicographic stable-ID order, LF endings, and final newlines.
`data/source.json` holds coverage and unresolved-review metadata because it is
aggregate provenance rather than an entity collection.

Object and enum IDs use the normalized selected `api_name`; member IDs append
the member segment. IDs do not depend on page title case, heading position, or
filesystem order. Metadata/title disagreements are skipped with a source
coverage entry. A same-name event and method gains a `#kind` suffix so it stays
both stable and distinguishable; a same-kind collision fails ingestion.

Each record's `source` describes the repository, pin, relative path, and
extraction method. Its `provenance` map distinguishes source front matter,
source syntax/table/return text, short source text, project-authored summary,
and project classification. The summary maximum is 240 characters. No external
LLM, network service, remarks section, or full example is used to fill it.

`verification.documentation` is true for selected records. `typelib` and
`runtime` are false; `dispatch.dispid` and `invoke_kinds` are null/empty. Every
implementation status is `unplanned`. This separates a documented fact from
later type-library and real-Excel evidence.

## Ingestion pipeline

`tools/excel-com-kb` is a standalone, unpublished Rust utility, excluded like
the repository's other standalone tool. It has no path or production dependency
on `excel-api`, `excel-api-sys`, `excel-api-macros`, the minimal XLL, or the RTD
prototype. Its standalone lock file makes its offline tool/test resolution
reproducible without extending the core workspace dependency graph.

`ingest` reads a supplied pinned checkout, YAML front matter, title, short
description, syntax access marker, parameter table, return-value section,
example headings, and enum tables. It normalizes records and emits explicit
warnings/review entries for missing metadata, malformed tables, aliases,
ambiguous access, unknown return-object types, collection ambiguity, and
conflicting page metadata. `validate` enforces the checked-in schema contract,
ordering, unique IDs, source paths/commits, allowed statuses, owner/relationship
endpoints, initial verification policy, portable paths, and output endings.
`generate` writes the four deterministic Markdown reports. `check` validates
and verifies regenerated reports are current and repeatable.

## Current coverage and limitations

At the pinned revision, the full run selected **5,928** files and parsed
**5,156** pages into **268** objects, **3,337** properties, **1,186** methods,
**109** events, **256** enums, and **5,770** relationships. It skipped **772**
files, including 513 without `api_name`, 252 with titles outside the supported
entity taxonomy, and seven source metadata/title conflicts. The generated
[source coverage](../../../knowledge/excel-object-model/generated/source-coverage.md)
and [unresolved classification](../../../knowledge/excel-object-model/generated/unresolved-classification.md)
reports are authoritative for exact review counts.

The parser deliberately does not infer a return object from arbitrary prose,
derive DISPIDs, invent access flags, assume a default member, interpret
generated include child lists, or resolve aliases that alter meaning. Those
limitations leave a large, visible review queue—not a silent false claim of
complete semantic coverage.

Type-library inspection must later verify interface identity, member kinds in
collisions, DISPIDs, invocation flags, parameter directions/types, default
members, enum values, and aliases. Real Excel experiments must later verify
activation/lifetime, object availability, collection behavior and indexing,
optional-argument effects, events, locale/coercion, `SAFEARRAY` results, and
shutdown behavior.

## Regeneration and later use

Use the commands in the [knowledge-base README](../../../knowledge/excel-object-model/README.md)
to ingest a separately supplied checkout, validate, generate, and review a
zero-diff regeneration. An upstream update must change the manifest pin and
retrieval/licence evidence first; ordinary validation never follows a mutable
branch tip or network URL.

Later research should cite a canonical ID and source path, carry forward its
verification dimensions, and add type-library/runtime evidence in a reviewed
schema/tool update. Later prompts may use the records for object inventories,
implementation tracking, generated documentation, and compatibility audits;
they must not treat documentation extraction as a COM behavior experiment or a
frozen public wrapper design.
